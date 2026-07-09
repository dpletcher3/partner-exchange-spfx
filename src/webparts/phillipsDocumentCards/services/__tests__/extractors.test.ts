import { resolveDocUrl, resolveCardTarget } from '../extractors';

// D061: the card click target prefers the item-level ServerRedirectedEmbedUrl
// (SharePoint's browser-open URL, routed through the Office/PDF web viewer) and
// falls back to the raw FileRef server-relative path when that value is empty
// (non-previewable types like .zip / .txt have no embed URL and stay download-only).
describe('resolveDocUrl', () => {
  const EMBED = 'https://contoso.sharepoint.com/sites/it/_layouts/15/Doc.aspx?sourcedoc={abc}&action=interactivepreview';
  const FILEREF = '/sites/it/Phillips Documents/Guide.docx';

  it('returns ServerRedirectedEmbedUrl when populated (ignores FileRef)', () => {
    expect(resolveDocUrl(EMBED, FILEREF)).toBe(EMBED);
  });

  it('falls back to FileRef when ServerRedirectedEmbedUrl is an empty string', () => {
    expect(resolveDocUrl('', FILEREF)).toBe(FILEREF);
  });

  it('falls back to FileRef when ServerRedirectedEmbedUrl is null or undefined', () => {
    expect(resolveDocUrl(null, FILEREF)).toBe(FILEREF);
    expect(resolveDocUrl(undefined, FILEREF)).toBe(FILEREF);
  });

  it('returns "" when both values are empty', () => {
    expect(resolveDocUrl('', '')).toBe('');
    expect(resolveDocUrl(undefined, undefined)).toBe('');
  });
});

// D062: an external CardLink (Hyperlink column) wins — the card links out and opens
// in a new tab; otherwise the card is an I25 document link (same-tab viewer via
// ServerRedirectedEmbedUrl, FileRef fallback).
describe('resolveCardTarget', () => {
  const EMBED = 'https://contoso.sharepoint.com/sites/it/_layouts/15/Doc.aspx?sourcedoc={abc}&action=interactivepreview';
  const FILEREF = '/sites/it/Phillips Documents/Guide.docx';
  const EXTERNAL = 'https://books.phillipscorp.com/handbook';

  it('external when CardLink is a { Url } object (ignores embed/FileRef)', () => {
    expect(resolveCardTarget({ Url: EXTERNAL, Description: EXTERNAL }, EMBED, FILEREF)).toEqual({
      href: EXTERNAL,
      external: true
    });
  });

  it('external when CardLink is a bare string', () => {
    expect(resolveCardTarget(EXTERNAL, EMBED, FILEREF)).toEqual({ href: EXTERNAL, external: true });
  });

  it('document (same-tab) when CardLink is empty — prefers the embed URL', () => {
    expect(resolveCardTarget('', EMBED, FILEREF)).toEqual({ href: EMBED, external: false });
    expect(resolveCardTarget(null, EMBED, FILEREF)).toEqual({ href: EMBED, external: false });
    expect(resolveCardTarget(undefined, EMBED, FILEREF)).toEqual({ href: EMBED, external: false });
  });

  it('document falls back to FileRef when CardLink and embed are both empty', () => {
    expect(resolveCardTarget('', '', FILEREF)).toEqual({ href: FILEREF, external: false });
  });

  it('href "" and external false when everything is empty', () => {
    expect(resolveCardTarget(undefined, undefined, undefined)).toEqual({ href: '', external: false });
  });
});
