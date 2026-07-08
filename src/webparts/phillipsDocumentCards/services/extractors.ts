// Pure shape-extraction helpers for SharePoint REST responses (no SPHttpClient
// dependency, so they unit-test in isolation). Adapted verbatim from the Media
// Gallery's extractors — the Image-column shapes are identical (the I16 library
// uses the same modern Image/Thumbnail column type). Permissive on input,
// conservative on output (return '' / undefined rather than something broken).
// They warn ONLY when input is present but malformed — never for a legitimately
// empty value.
//
// Returns `undefined` (not `null`) per the rig's @rushstack/no-new-null rule.

const LOG = '[DocumentCards]';

export interface IRawUrlField {
  Url?: string;
  Description?: string;
}

export interface IRawAttachmentFile {
  FileName?: string;
  ServerRelativeUrl?: string;
}

// URL/Hyperlink column. SharePoint returns { Url, Description }; a bare string
// has been observed depending on metadata. Returns '' when no usable URL.
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
  console.warn(`${LOG} Unexpected URL field shape`, raw);
  return '';
}

// The click target for a document card (D061). Prefers the item-level
// ServerRedirectedEmbedUrl — SharePoint's canonical browser-open URL, routed
// through the Office/PDF web viewer so the file OPENS rather than downloads.
// That value is populated for previewable types and EMPTY for others (.zip,
// .txt, …), so we fall back to the raw FileRef server-relative path, which
// still resolves to a working link (those types legitimately download).
export function resolveDocUrl(serverRedirectedEmbedUrl: unknown, fileRef: unknown): string {
  return asString(serverRedirectedEmbedUrl) || asString(fileRef);
}

// Image / Thumbnail column. The value is a stringified JSON blob. Two shapes
// occur in the wild:
//   1. Modern Image column with an inline URL: { serverRelativeUrl | serverUrl, … }
//   2. "Reserved attachment" form: { fileName, originalImageName } — no URL in
//      the blob; the real URL lives in the item's AttachmentFiles entry whose
//      FileName matches `fileName`.
// Tries the inline URL first, then the AttachmentFiles lookup. Returns undefined
// when no usable URL can be derived (caller shows a placeholder).
export function extractImageColumnUrl(
  rawImage: unknown,
  attachmentFiles: unknown
): string | undefined {
  // Legitimate empty: the item has no image set. Not a shape problem.
  if (rawImage === null || rawImage === undefined || rawImage === '') {
    return undefined;
  }

  if (typeof rawImage !== 'string') {
    console.warn(`${LOG} Failed to parse Image column (not a string)`, rawImage);
    return undefined;
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(rawImage);
  } catch {
    console.warn(`${LOG} Failed to parse Image column JSON`, rawImage);
    return undefined;
  }

  // Shape 1: an inline URL on the blob.
  const inlineUrl = extractInlineUrl(parsed);
  if (inlineUrl) {
    return inlineUrl;
  }

  // Shape 2: fileName → AttachmentFiles lookup.
  const fileName = extractFileName(parsed);
  if (!fileName) {
    console.warn(`${LOG} Unexpected Image column shape (no url or fileName)`, parsed);
    return undefined;
  }

  const url = findAttachmentUrl(attachmentFiles, fileName);
  if (!url) {
    console.warn(`${LOG} Image fileName not found in AttachmentFiles`, fileName);
    return undefined;
  }
  return url;
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

function extractInlineUrl(parsed: unknown): string | undefined {
  if (parsed && typeof parsed === 'object') {
    const obj = parsed as { serverRelativeUrl?: unknown; serverUrl?: unknown };
    if (typeof obj.serverRelativeUrl === 'string' && obj.serverRelativeUrl) {
      return obj.serverRelativeUrl;
    }
    if (typeof obj.serverUrl === 'string' && obj.serverUrl) {
      return obj.serverUrl;
    }
  }
  return undefined;
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

// AttachmentFiles arrives as a plain array (minimal metadata) or { results: [...] }.
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

// Defensive string coercion for text/choice/built-in columns. Empty for
// null/undefined; passes strings through; stringifies anything else.
export function asString(value: unknown): string {
  if (value === null || value === undefined) {
    return '';
  }
  if (typeof value === 'string') {
    return value;
  }
  return String(value);
}
