// Pure shape-extraction helpers for the SharePoint news-pipeline data source
// (Site Pages, PromotedState=2). Kept separate from extractors.ts on purpose:
// the list-mode ThumbnailImage+AttachmentFiles resolver does NOT apply to news
// pages, which carry their thumbnail in a single BannerImageUrl field. These
// helpers follow the same discipline as extractors.ts — permissive on input,
// conservative on output, warn only when a value is PRESENT but in an
// unexpected shape (never for a legitimately empty value), so the console stays
// clean on the happy path.
//
// Runtime field shapes on Site Pages differ between m365-CLI inspection and the
// SPHttpClient response (an established lesson on this project), and the
// news-link marker in particular is not guaranteed across tenants/versions.
// Everything here is therefore defensive: an unrecognized shape degrades to the
// safe interpretation (treat the page as a News post) rather than throwing.

import { NEWS_POST_ITEM_TYPE, NEWS_LINK_ITEM_TYPE } from '../config/constants';

// --- BannerImageUrl ----------------------------------------------------------

// The news-page thumbnail. BannerImageUrl is a URL/picture column: SharePoint
// usually returns { Url, Description }, but a bare string has been observed.
// Returns undefined for an absent value (the card renders the red fallback);
// warns only for a non-empty object that lacks a usable Url.
export function extractBannerImageUrl(raw: unknown): string | undefined {
  if (raw === null || raw === undefined || raw === '') {
    return undefined;
  }
  if (typeof raw === 'string') {
    return raw;
  }
  if (typeof raw === 'object') {
    const url = (raw as { Url?: string }).Url;
    if (typeof url === 'string' && url) {
      return url;
    }
  }
  console.warn('[PhillipsNews] Unexpected BannerImageUrl shape', raw);
  return undefined;
}

// --- News-link redirect URL --------------------------------------------------

// The external target of a News link lives in a redirect/original-source URL
// field (a Text or URL column depending on shape). Quiet by design: a News
// POST legitimately has no redirect URL, so an absent value is normal and must
// not warn. Returns '' when there is no usable URL.
export function coerceUrlString(raw: unknown): string {
  if (raw === null || raw === undefined || raw === '') {
    return '';
  }
  if (typeof raw === 'string') {
    return raw;
  }
  if (typeof raw === 'object') {
    const url = (raw as { Url?: string }).Url;
    if (typeof url === 'string' && url) {
      return url;
    }
  }
  return '';
}

// --- Page flags --------------------------------------------------------------

// OData__SPSitePageFlags can surface as a plain array (minimal metadata), a
// { results: [...] } wrapper (verbose), or a comma-delimited string. Returns
// the normalized token list; [] for an absent value (no warning).
export function parseFlagTokens(raw: unknown): string[] {
  if (!raw) {
    return [];
  }
  if (Array.isArray(raw)) {
    return raw.filter(isNonEmptyString);
  }
  if (typeof raw === 'object') {
    const results = (raw as { results?: unknown }).results;
    if (Array.isArray(results)) {
      return results.filter(isNonEmptyString);
    }
  }
  if (typeof raw === 'string') {
    return raw
      .split(',')
      .map((token) => token.trim())
      .filter((token) => token.length > 0);
  }
  return [];
}

// --- URL helpers -------------------------------------------------------------

// Extracts the scheme+host origin (e.g. "https://contoso.sharepoint.com") from
// a site URL string. Returns '' if the input has no recognizable origin.
export function extractOrigin(siteUrl: string): string {
  const match = /^(https?:\/\/[^/]+)/i.exec(siteUrl || '');
  return match ? match[1] : '';
}

// Resolves a page's FileRef (a server-relative URL like
// "/sites/x/SitePages/foo.aspx") to an absolute URL using the site origin. A
// FileRef that is already absolute is returned unchanged.
export function toAbsoluteUrl(fileRef: string, origin: string): string {
  if (!fileRef) {
    return '';
  }
  if (/^https?:\/\//i.test(fileRef)) {
    return fileRef;
  }
  return `${origin}${fileRef}`;
}

// --- Post vs link derivation -------------------------------------------------

export interface IPageLinkInput {
  // OData__SPSitePageFlags (may name "News link" among its tokens).
  flags: unknown;
  // OData__OriginalSourceUrl (the external target on a News link / Repost Page).
  redirectUrlRaw: unknown;
  // FileRef of the page (server-relative).
  fileRef: string;
  // Site origin for absolutizing the FileRef of a News post.
  origin: string;
}

export interface IPageLink {
  itemType: string; // NEWS_POST_ITEM_TYPE | NEWS_LINK_ITEM_TYPE
  linkUrl: string;
}

// Derives the item type and click-through URL for a news page:
//   - News link → the external redirect URL (the whole point of a News link).
//   - News post → the page itself (FileRef made absolute).
//
// Two independent signals are consulted because neither is guaranteed across
// tenants: a resolvable redirect URL, and a "News link" token in the page
// flags. A present redirect URL is authoritative (a News post never has one).
// If the flags claim "News link" but no redirect URL resolves, the shape is
// unrecognized — we warn and fall back to treating the page as a post, per the
// defensive contract, rather than emitting a dead/empty link.
export function derivePostOrLink(input: IPageLinkInput): IPageLink {
  const redirectUrl = coerceUrlString(input.redirectUrlRaw);
  if (redirectUrl) {
    return { itemType: NEWS_LINK_ITEM_TYPE, linkUrl: redirectUrl };
  }

  const flaggedAsLink = parseFlagTokens(input.flags).some((token) =>
    /news\s*link/i.test(token)
  );
  if (flaggedAsLink) {
    console.warn(
      '[PhillipsNews] Page flagged as News link but no redirect URL resolved; treating as post',
      input.flags
    );
  }

  return {
    itemType: NEWS_POST_ITEM_TYPE,
    linkUrl: toAbsoluteUrl(input.fileRef, input.origin)
  };
}

// --- Internal ----------------------------------------------------------------

function isNonEmptyString(value: unknown): value is string {
  return typeof value === 'string' && value.length > 0;
}
