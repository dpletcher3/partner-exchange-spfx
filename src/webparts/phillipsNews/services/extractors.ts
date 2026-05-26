// Pure shape-extraction helpers for SharePoint REST responses.
//
// SharePoint serializes typed columns (URL, Image, MultiChoice) in shapes that
// vary between metadata levels and between m365-CLI inspection and the runtime
// SPHttpClient response. These helpers are intentionally permissive on input
// (accept `unknown`, handle every shape we've observed plus a few we haven't)
// and conservative on output (return an empty/none value rather than something
// that would render broken). They log a warning ONLY when the input is present
// but in an unexpected shape — never for a legitimately empty value — so the
// console stays clean on the happy path and a warning is a real diagnostic.
//
// Kept as standalone pure functions (no SPHttpClient dependency) so they are
// unit-testable in isolation.

export interface IRawUrlField {
  Url?: string;
  Description?: string;
}

export interface IRawAttachmentFile {
  FileName?: string;
  ServerRelativeUrl?: string;
}

// URL column. SharePoint usually returns { Url, Description }, but a bare string
// has been observed depending on column/metadata. Returns '' when there is no
// usable URL; warns only for a non-empty object that lacks a Url.
export function extractUrl(raw: unknown): string {
  if (raw === null || raw === undefined || raw === '') {
    return '';
  }
  if (typeof raw === 'string') {
    return raw;
  }
  if (typeof raw === 'object') {
    const url = (raw as IRawUrlField).Url;
    if (typeof url === 'string' && url) {
      return url;
    }
  }
  console.warn('[PhillipsNews] Unexpected LinkUrl shape', raw);
  return '';
}

// Choice / MultiChoice values and a field's Choices collection. Minimal metadata
// returns a plain array; verbose returns { results: [...] }; a single string is
// handled defensively. Returns [] for an empty/absent value (no warning).
export function extractChoices(raw: unknown): string[] {
  if (!raw) {
    return [];
  }
  if (Array.isArray(raw)) {
    return raw.filter(isNonEmptyString);
  }
  if (typeof raw === 'string') {
    return [raw];
  }
  if (typeof raw === 'object') {
    const results = (raw as { results?: unknown }).results;
    if (Array.isArray(results)) {
      return results.filter(isNonEmptyString);
    }
  }
  return [];
}

// Image column. The ThumbnailImage field value carries no URL — only a fileName
// that matches an entry in the item's AttachmentFiles, where the real
// ServerRelativeUrl lives. Returns the matching URL, or undefined when no usable
// URL can be derived (caller renders the red fallback).
//
// Returns undefined (not the spec's literal `null`) to comply with the rig's
// @rushstack/no-new-null lint rule; semantics are identical. See the
// list-service-turn spec deviation note.
export function extractThumbnailUrl(
  rawThumbnail: unknown,
  attachmentFiles: unknown
): string | undefined {
  // Legitimate empty: the item has no thumbnail set. Not a shape problem.
  if (rawThumbnail === null || rawThumbnail === undefined || rawThumbnail === '') {
    return undefined;
  }

  if (typeof rawThumbnail !== 'string') {
    console.warn('[PhillipsNews] Failed to parse ThumbnailImage', rawThumbnail);
    return undefined;
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(rawThumbnail);
  } catch {
    console.warn('[PhillipsNews] Failed to parse ThumbnailImage', rawThumbnail);
    return undefined;
  }

  const fileName = extractFileName(parsed);
  if (!fileName) {
    console.warn('[PhillipsNews] Unexpected ThumbnailImage shape', parsed);
    return undefined;
  }

  const serverRelativeUrl = findAttachmentUrl(attachmentFiles, fileName);
  if (!serverRelativeUrl) {
    console.warn(
      '[PhillipsNews] ThumbnailImage filename not found in AttachmentFiles',
      fileName
    );
    return undefined;
  }

  return serverRelativeUrl;
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

function isNonEmptyString(value: unknown): value is string {
  return typeof value === 'string' && value.length > 0;
}

function extractFileName(parsed: unknown): string | undefined {
  if (parsed && typeof parsed === 'object') {
    const fileName = (parsed as { fileName?: unknown }).fileName;
    if (typeof fileName === 'string' && fileName) {
      return fileName;
    }
  }
  return undefined;
}

function findAttachmentUrl(attachmentFiles: unknown, fileName: string): string | undefined {
  for (const attachment of toAttachmentArray(attachmentFiles)) {
    if (attachment && typeof attachment === 'object') {
      const entry = attachment as IRawAttachmentFile;
      if (entry.FileName === fileName && typeof entry.ServerRelativeUrl === 'string') {
        return entry.ServerRelativeUrl;
      }
    }
  }
  return undefined;
}

// AttachmentFiles can arrive as a plain array (minimal metadata) or wrapped as
// { results: [...] } (verbose). Anything else yields an empty list.
function toAttachmentArray(attachmentFiles: unknown): unknown[] {
  if (Array.isArray(attachmentFiles)) {
    return attachmentFiles;
  }
  if (attachmentFiles && typeof attachmentFiles === 'object') {
    const results = (attachmentFiles as { results?: unknown }).results;
    if (Array.isArray(results)) {
      return results;
    }
  }
  return [];
}
