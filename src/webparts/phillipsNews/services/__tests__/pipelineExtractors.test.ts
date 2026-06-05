import {
  extractBannerImageUrl,
  coerceUrlString,
  parseFlagTokens,
  extractOrigin,
  toAbsoluteUrl,
  derivePostOrLink
} from '../pipelineExtractors';
import { NEWS_POST_ITEM_TYPE, NEWS_LINK_ITEM_TYPE } from '../../config/constants';
import * as fx from './fixtures/raw-pages';

// Same discipline as extractors.test.ts: each helper is exercised for happy
// path, alternative shapes, empty input (no warning), and unexpected shape
// (conservative result, warning where the contract promises one).

describe('extractBannerImageUrl', () => {
  let warnSpy: jest.SpyInstance;
  beforeEach(() => {
    warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => undefined);
  });
  afterEach(() => {
    warnSpy.mockRestore();
  });

  it('happy path: { Url } object returns Url', () => {
    expect(extractBannerImageUrl(fx.bannerAsObject)).toBe(fx.bannerAsObjectExpected);
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('alternative shape: bare string returns the string', () => {
    expect(extractBannerImageUrl(fx.bannerAsString)).toBe(fx.bannerAsString);
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('null input returns undefined, no warning', () => {
    expect(extractBannerImageUrl(fx.bannerNull)).toBeUndefined();
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('empty string input returns undefined, no warning', () => {
    expect(extractBannerImageUrl(fx.bannerEmpty)).toBeUndefined();
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('object without Url returns undefined and warns with the shape', () => {
    expect(extractBannerImageUrl(fx.bannerObjectWithoutUrl)).toBeUndefined();
    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining('[PhillipsNews] Unexpected BannerImageUrl shape'),
      fx.bannerObjectWithoutUrl
    );
  });
});

describe('coerceUrlString', () => {
  let warnSpy: jest.SpyInstance;
  beforeEach(() => {
    warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => undefined);
  });
  afterEach(() => {
    warnSpy.mockRestore();
  });

  it('string returns as-is', () => {
    expect(coerceUrlString(fx.redirectAsString)).toBe(fx.redirectAsString);
  });

  it('{ Url } object returns Url', () => {
    expect(coerceUrlString(fx.redirectAsObject)).toBe('https://example.com/external');
  });

  it('null returns empty string, never warns (absent redirect is normal)', () => {
    expect(coerceUrlString(fx.redirectNull)).toBe('');
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('empty string returns empty string, never warns', () => {
    expect(coerceUrlString(fx.redirectEmpty)).toBe('');
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('object without Url returns empty string, never warns', () => {
    expect(coerceUrlString({ Description: 'x' })).toBe('');
    expect(warnSpy).not.toHaveBeenCalled();
  });
});

describe('parseFlagTokens', () => {
  it('plain array is filtered to non-empty strings', () => {
    expect(parseFlagTokens(fx.flagsArrayWithNewsLink)).toEqual(['Promoted', 'News link']);
  });

  it('{ results } wrapper is unwrapped', () => {
    expect(parseFlagTokens(fx.flagsWrappedWithNewsLink)).toEqual(['News link']);
  });

  it('comma-delimited string is split and trimmed', () => {
    expect(parseFlagTokens(fx.flagsCommaString)).toEqual(['Promoted', 'News link']);
  });

  it('null returns empty array', () => {
    expect(parseFlagTokens(fx.flagsNull)).toEqual([]);
  });
});

describe('extractOrigin', () => {
  it('extracts scheme+host from a site URL', () => {
    expect(extractOrigin('https://phillipscorp.sharepoint.com/sites/PartnerExchange-Sandbox')).toBe(
      'https://phillipscorp.sharepoint.com'
    );
  });

  it('returns empty string when there is no recognizable origin', () => {
    expect(extractOrigin('not-a-url')).toBe('');
  });
});

describe('toAbsoluteUrl', () => {
  it('prefixes a server-relative URL with the origin', () => {
    expect(toAbsoluteUrl('/sites/x/SitePages/a.aspx', 'https://phillipscorp.sharepoint.com')).toBe(
      'https://phillipscorp.sharepoint.com/sites/x/SitePages/a.aspx'
    );
  });

  it('returns an already-absolute URL unchanged', () => {
    expect(toAbsoluteUrl('https://host/x.aspx', 'https://phillipscorp.sharepoint.com')).toBe(
      'https://host/x.aspx'
    );
  });

  it('returns empty string for an empty FileRef', () => {
    expect(toAbsoluteUrl('', 'https://phillipscorp.sharepoint.com')).toBe('');
  });
});

describe('derivePostOrLink', () => {
  let warnSpy: jest.SpyInstance;
  beforeEach(() => {
    warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => undefined);
  });
  afterEach(() => {
    warnSpy.mockRestore();
  });

  it('post: no redirect URL → absolute FileRef, no warning', () => {
    const result = derivePostOrLink(fx.pagePostBasic);
    expect(result.itemType).toBe(NEWS_POST_ITEM_TYPE);
    expect(result.linkUrl).toBe(fx.pagePostBasic.expectedLink);
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('link: redirect URL + News link flag → external URL', () => {
    const result = derivePostOrLink(fx.pageLinkWithRedirect);
    expect(result.itemType).toBe(NEWS_LINK_ITEM_TYPE);
    expect(result.linkUrl).toBe(fx.pageLinkWithRedirect.expectedLink);
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('link: present redirect URL is authoritative even without the flag', () => {
    const result = derivePostOrLink(fx.pageLinkRedirectNoFlag);
    expect(result.itemType).toBe(NEWS_LINK_ITEM_TYPE);
    expect(result.linkUrl).toBe(fx.pageLinkRedirectNoFlag.expectedLink);
  });

  it('flagged as link but no redirect URL → warns and falls back to post', () => {
    const result = derivePostOrLink(fx.pageFlaggedLinkNoUrl);
    expect(result.itemType).toBe(NEWS_POST_ITEM_TYPE);
    expect(result.linkUrl).toBe(fx.pageFlaggedLinkNoUrl.expectedLink);
    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining('[PhillipsNews] Page flagged as News link but no redirect URL resolved'),
      fx.pageFlaggedLinkNoUrl.flags
    );
  });

  it('post with an already-absolute FileRef is returned unchanged', () => {
    const result = derivePostOrLink(fx.pagePostAbsoluteFileRef);
    expect(result.itemType).toBe(NEWS_POST_ITEM_TYPE);
    expect(result.linkUrl).toBe(fx.pagePostAbsoluteFileRef.expectedLink);
  });
});
