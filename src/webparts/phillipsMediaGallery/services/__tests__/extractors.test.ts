import { extractUrl, extractImageColumnUrl } from '../extractors';

describe('extractUrl', () => {
  let warnSpy: jest.SpyInstance;
  beforeEach(() => {
    warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => undefined);
  });
  afterEach(() => warnSpy.mockRestore());

  it('{ Url, Description } object returns Url', () => {
    expect(extractUrl({ Url: 'https://vimeo.com/123', Description: 'https://vimeo.com/123' })).toBe(
      'https://vimeo.com/123'
    );
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('bare string returns the string', () => {
    expect(extractUrl('https://vimeo.com/123')).toBe('https://vimeo.com/123');
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('null / empty return "" with no warning', () => {
    expect(extractUrl(null)).toBe('');
    expect(extractUrl(undefined)).toBe('');
    expect(extractUrl('')).toBe('');
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('object without Url returns "" and warns', () => {
    expect(extractUrl({ Description: 'x' })).toBe('');
    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining('[MediaGallery] Unexpected URL field shape'),
      { Description: 'x' }
    );
  });
});

describe('extractImageColumnUrl', () => {
  let warnSpy: jest.SpyInstance;
  beforeEach(() => {
    warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => undefined);
  });
  afterEach(() => warnSpy.mockRestore());

  // The 15 Practices shape: { fileName, originalImageName }, resolved via the
  // item's AttachmentFiles (FileName → ServerRelativeUrl).
  const RESERVED = JSON.stringify({
    fileName: 'Reserved_ImageAttachment_[6]_[Image0][32]_[abc][1]_[8].png',
    originalImageName: 'ChatGPT Image'
  });
  const attachments = [
    { FileName: 'other.png', ServerRelativeUrl: '/sites/x/Attachments/1/other.png' },
    {
      FileName: 'Reserved_ImageAttachment_[6]_[Image0][32]_[abc][1]_[8].png',
      ServerRelativeUrl: '/sites/x/Lists/15 Practices/Attachments/1/Reserved_ImageAttachment_[6]_[Image0][32]_[abc][1]_[8].png'
    }
  ];

  it('resolves fileName via AttachmentFiles (reserved-attachment shape)', () => {
    expect(extractImageColumnUrl(RESERVED, attachments)).toBe(
      '/sites/x/Lists/15 Practices/Attachments/1/Reserved_ImageAttachment_[6]_[Image0][32]_[abc][1]_[8].png'
    );
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('AttachmentFiles wrapped in { results: [...] } also resolves', () => {
    expect(extractImageColumnUrl(RESERVED, { results: attachments })).toContain(
      '/Reserved_ImageAttachment_'
    );
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('inline serverRelativeUrl shape returns it directly (no attachments needed)', () => {
    const inline = JSON.stringify({ type: 'thumbnail', serverRelativeUrl: '/sites/x/SiteAssets/pic.jpg' });
    expect(extractImageColumnUrl(inline, [])).toBe('/sites/x/SiteAssets/pic.jpg');
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('null / empty return undefined with no warning', () => {
    expect(extractImageColumnUrl(null, attachments)).toBeUndefined();
    expect(extractImageColumnUrl('', attachments)).toBeUndefined();
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('non-string value warns and returns undefined', () => {
    expect(extractImageColumnUrl({ fileName: 'x' }, attachments)).toBeUndefined();
    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining('[MediaGallery] Failed to parse Image column (not a string)'),
      { fileName: 'x' }
    );
  });

  it('malformed JSON warns and returns undefined', () => {
    expect(extractImageColumnUrl('{not valid json', attachments)).toBeUndefined();
    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining('[MediaGallery] Failed to parse Image column JSON'),
      '{not valid json'
    );
  });

  it('parsed but no url or fileName warns and returns undefined', () => {
    expect(extractImageColumnUrl(JSON.stringify({ originalImageName: 'x' }), attachments)).toBeUndefined();
    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining('[MediaGallery] Unexpected Image column shape'),
      { originalImageName: 'x' }
    );
  });

  it('fileName not found in AttachmentFiles warns and returns undefined', () => {
    const missing = JSON.stringify({ fileName: 'nope.png' });
    expect(extractImageColumnUrl(missing, attachments)).toBeUndefined();
    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining('[MediaGallery] Image fileName not found in AttachmentFiles'),
      'nope.png'
    );
  });
});
