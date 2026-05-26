import { extractUrl, extractChoices, extractThumbnailUrl } from '../extractors';
import * as fx from './fixtures/raw-list-items';

// Each extractor is exercised for: happy path, the alternative shape we've
// observed, null/undefined input (empty result, no warning), malformed input
// (empty result, warning), and empty content (empty result, no warning).
// Warnings are verified by spying on console.warn and asserting both the
// message prefix and that the offending shape is included in the payload.

describe('extractUrl', () => {
  let warnSpy: jest.SpyInstance;
  beforeEach(() => {
    warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => undefined);
  });
  afterEach(() => {
    warnSpy.mockRestore();
  });

  it('happy path: { Url, Description } object returns Url', () => {
    expect(extractUrl(fx.urlAsObject)).toBe('https://phillipscorp.com/article');
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('alternative shape: bare string returns the string', () => {
    expect(extractUrl(fx.urlAsString)).toBe('https://phillipscorp.com/article');
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('null input returns empty string, no warning', () => {
    expect(extractUrl(fx.urlNull)).toBe('');
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('empty string input returns empty string, no warning', () => {
    expect(extractUrl(fx.urlEmpty)).toBe('');
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('object without Url returns empty string and warns with the shape', () => {
    expect(extractUrl(fx.urlObjectWithoutUrl)).toBe('');
    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining('[PhillipsNews] Unexpected LinkUrl shape'),
      fx.urlObjectWithoutUrl
    );
  });
});

describe('extractChoices', () => {
  let warnSpy: jest.SpyInstance;
  beforeEach(() => {
    warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => undefined);
  });
  afterEach(() => {
    warnSpy.mockRestore();
  });

  it('happy path: plain array returns as-is', () => {
    expect(extractChoices(fx.categoriesAsArray)).toEqual([
      'Phillips Loop',
      'Phillips In The News'
    ]);
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('alternative shape: { results } wrapper is unwrapped', () => {
    expect(extractChoices(fx.categoriesAsWrapped)).toEqual([
      'Phillips Loop',
      'Phillips In The News'
    ]);
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('single string is wrapped into a one-element array', () => {
    expect(extractChoices(fx.categoriesAsString)).toEqual(['Phillips Loop']);
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('null input returns empty array, no warning', () => {
    expect(extractChoices(fx.categoriesNull)).toEqual([]);
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('empty array input returns empty array, no warning', () => {
    expect(extractChoices([])).toEqual([]);
    expect(warnSpy).not.toHaveBeenCalled();
  });
});

describe('extractThumbnailUrl', () => {
  let warnSpy: jest.SpyInstance;
  beforeEach(() => {
    warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => undefined);
  });
  afterEach(() => {
    warnSpy.mockRestore();
  });

  it('happy path: filename matches an attachment, returns its ServerRelativeUrl', () => {
    const result = extractThumbnailUrl(
      fx.thumbnailValidWithMatch.rawThumbnail,
      fx.thumbnailValidWithMatch.attachmentFiles
    );
    expect(result).toBe(fx.thumbnailValidWithMatch.expectedUrl);
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('null field returns undefined, no warning (legitimate empty)', () => {
    expect(
      extractThumbnailUrl(fx.thumbnailNullField.rawThumbnail, fx.thumbnailNullField.attachmentFiles)
    ).toBeUndefined();
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('empty string field returns undefined, no warning (legitimate empty)', () => {
    expect(
      extractThumbnailUrl(
        fx.thumbnailEmptyStringField.rawThumbnail,
        fx.thumbnailEmptyStringField.attachmentFiles
      )
    ).toBeUndefined();
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('filename not in attachments returns undefined and warns with the filename', () => {
    const result = extractThumbnailUrl(
      fx.thumbnailFilenameNotInAttachments.rawThumbnail,
      fx.thumbnailFilenameNotInAttachments.attachmentFiles
    );
    expect(result).toBeUndefined();
    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining('[PhillipsNews] ThumbnailImage filename not found in AttachmentFiles'),
      fx.thumbnailFilenameNotInAttachments.expectedFileName
    );
  });

  it('malformed JSON returns undefined and warns with the raw value', () => {
    const result = extractThumbnailUrl(
      fx.thumbnailMalformedJson.rawThumbnail,
      fx.thumbnailMalformedJson.attachmentFiles
    );
    expect(result).toBeUndefined();
    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining('[PhillipsNews] Failed to parse ThumbnailImage'),
      fx.thumbnailMalformedJson.rawThumbnail
    );
  });

  it('parsed object missing fileName returns undefined and warns with the parsed shape', () => {
    const result = extractThumbnailUrl(
      fx.thumbnailMissingFileNameKey.rawThumbnail,
      fx.thumbnailMissingFileNameKey.attachmentFiles
    );
    expect(result).toBeUndefined();
    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining('[PhillipsNews] Unexpected ThumbnailImage shape'),
      fx.thumbnailMissingFileNameKey.parsed
    );
  });

  it('undefined attachmentFiles returns undefined and warns filename-not-found', () => {
    const result = extractThumbnailUrl(
      fx.thumbnailAttachmentFilesUndefined.rawThumbnail,
      fx.thumbnailAttachmentFilesUndefined.attachmentFiles
    );
    expect(result).toBeUndefined();
    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining('[PhillipsNews] ThumbnailImage filename not found in AttachmentFiles'),
      fx.thumbnailAttachmentFilesUndefined.expectedFileName
    );
  });
});
