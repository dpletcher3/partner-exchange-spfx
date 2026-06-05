// Fixtures for the news-pipeline extractor unit tests. These mirror Site Pages
// REST payloads (PromotedState=2 news pages) plus crafted edge cases. SP REST
// legitimately returns `null` for empty typed fields, so the null literals here
// are intentional (value positions, not type positions — no no-new-null exemption).

const SITE = 'https://phillipscorp.sharepoint.com/sites/PartnerExchange-Sandbox';

// --- extractBannerImageUrl ----------------------------------------------------

export const bannerAsObject = {
  Url: `${SITE}/_layouts/15/getpreview.ashx?guidFileId=abc`,
  Description: 'banner'
};
export const bannerAsObjectExpected = `${SITE}/_layouts/15/getpreview.ashx?guidFileId=abc`;
export const bannerAsString = `${SITE}/SiteAssets/news/banner.jpg`;
export const bannerNull = null;
export const bannerEmpty = '';
export const bannerObjectWithoutUrl = { Description: 'no url here' };

// --- coerceUrlString ----------------------------------------------------------

export const redirectAsString = 'https://www.modernmachineshop.com/article/123';
export const redirectAsObject = { Url: 'https://example.com/external', Description: 'ext' };
export const redirectNull = null;
export const redirectEmpty = '';

// --- parseFlagTokens ----------------------------------------------------------

export const flagsArrayWithNewsLink = ['Promoted', 'News link'];
export const flagsWrappedWithNewsLink = { results: ['News link'] };
export const flagsCommaString = 'Promoted, News link';
export const flagsArrayPostOnly = ['Promoted'];
export const flagsNull = null;

// --- derivePostOrLink ---------------------------------------------------------

export const ORIGIN = 'https://phillipscorp.sharepoint.com';

export const pagePostBasic = {
  flags: ['Promoted'],
  redirectUrlRaw: null as unknown,
  fileRef: '/sites/PartnerExchange-Sandbox/SitePages/amc.aspx',
  origin: ORIGIN,
  expectedLink: 'https://phillipscorp.sharepoint.com/sites/PartnerExchange-Sandbox/SitePages/amc.aspx'
};

export const pageLinkWithRedirect = {
  flags: ['Promoted', 'News link'],
  redirectUrlRaw: 'https://www.modernmachineshop.com/article/123',
  fileRef: '/sites/PartnerExchange-Sandbox/SitePages/repost-mms.aspx',
  origin: ORIGIN,
  expectedLink: 'https://www.modernmachineshop.com/article/123'
};

// Redirect URL present even though flags don't name "News link" — a present
// redirect is authoritative (a post never has one).
export const pageLinkRedirectNoFlag = {
  flags: null as unknown,
  redirectUrlRaw: 'https://example.com/news',
  fileRef: '/sites/x/SitePages/y.aspx',
  origin: ORIGIN,
  expectedLink: 'https://example.com/news'
};

// Flagged as a News link but no redirect URL resolves — unrecognized shape;
// derivation warns and falls back to treating it as a post.
export const pageFlaggedLinkNoUrl = {
  flags: ['News link'],
  redirectUrlRaw: null as unknown,
  fileRef: '/sites/x/SitePages/z.aspx',
  origin: ORIGIN,
  expectedLink: 'https://phillipscorp.sharepoint.com/sites/x/SitePages/z.aspx'
};

// FileRef already absolute — returned unchanged.
export const pagePostAbsoluteFileRef = {
  flags: null as unknown,
  redirectUrlRaw: null as unknown,
  fileRef: 'https://phillipscorp.sharepoint.com/sites/x/SitePages/abs.aspx',
  origin: ORIGIN,
  expectedLink: 'https://phillipscorp.sharepoint.com/sites/x/SitePages/abs.aspx'
};
