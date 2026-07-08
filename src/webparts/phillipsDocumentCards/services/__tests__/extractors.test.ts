import { resolveDocUrl } from '../extractors';

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
