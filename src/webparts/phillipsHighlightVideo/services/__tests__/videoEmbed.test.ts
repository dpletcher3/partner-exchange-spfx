import { resolveVideoEmbed } from '../videoEmbed';

// Covers the provider dispatch this web part adds. The underlying host/id parsing
// for the non-shorts forms is the Media Card Gallery's detectVideo, which has its
// own suite (phillipsMediaGallery/services/__tests__/videoThumbnail.test.ts) —
// these tests assert the DISPATCH and the embed-src construction, plus /shorts/,
// which detectVideo does not recognize.

describe('resolveVideoEmbed — YouTube', () => {
  it('watch?v= form', () => {
    expect(resolveVideoEmbed('https://www.youtube.com/watch?v=dQw4w9WgXcQ')).toEqual({
      provider: 'youtube',
      id: 'dQw4w9WgXcQ',
      embedSrc: 'https://www.youtube.com/embed/dQw4w9WgXcQ',
      providerLabel: 'YouTube'
    });
  });

  it('watch?v= with extra query params', () => {
    const r = resolveVideoEmbed('https://www.youtube.com/watch?v=dQw4w9WgXcQ&t=42s&list=PL123');
    expect(r).toBeDefined();
    expect(r?.provider).toBe('youtube');
    expect(r?.id).toBe('dQw4w9WgXcQ');
    expect(r?.embedSrc).toBe('https://www.youtube.com/embed/dQw4w9WgXcQ');
  });

  it('youtu.be short form', () => {
    expect(resolveVideoEmbed('https://youtu.be/dQw4w9WgXcQ')?.embedSrc).toBe(
      'https://www.youtube.com/embed/dQw4w9WgXcQ'
    );
  });

  it('youtu.be with query params', () => {
    expect(resolveVideoEmbed('https://youtu.be/dQw4w9WgXcQ?t=30')?.id).toBe('dQw4w9WgXcQ');
  });

  it('/embed/ form', () => {
    expect(resolveVideoEmbed('https://www.youtube.com/embed/dQw4w9WgXcQ?rel=0')?.id).toBe(
      'dQw4w9WgXcQ'
    );
  });

  it('/shorts/ form (not recognized by detectVideo — added here)', () => {
    expect(resolveVideoEmbed('https://www.youtube.com/shorts/abc123XYZ')).toEqual({
      provider: 'youtube',
      id: 'abc123XYZ',
      embedSrc: 'https://www.youtube.com/embed/abc123XYZ',
      providerLabel: 'YouTube'
    });
  });

  it('/shorts/ with query params', () => {
    expect(resolveVideoEmbed('https://www.youtube.com/shorts/abc123XYZ?feature=share')?.id).toBe(
      'abc123XYZ'
    );
  });

  it('m. subdomain still dispatches', () => {
    expect(resolveVideoEmbed('https://m.youtube.com/watch?v=abc123')?.provider).toBe('youtube');
  });
});

describe('resolveVideoEmbed — Vimeo (unchanged construction)', () => {
  it('plain vimeo.com/{id}', () => {
    expect(resolveVideoEmbed('https://vimeo.com/1165473810')).toEqual({
      provider: 'vimeo',
      id: '1165473810',
      embedSrc: 'https://player.vimeo.com/video/1165473810',
      providerLabel: 'Vimeo'
    });
  });

  it('vimeo.com/{id} with query string (15 Practices shape)', () => {
    expect(resolveVideoEmbed('https://vimeo.com/1165473810?fl=ip&fe=ec')?.embedSrc).toBe(
      'https://player.vimeo.com/video/1165473810'
    );
  });

  it('channel/showcase URL surfaces the numeric id', () => {
    expect(resolveVideoEmbed('https://vimeo.com/channels/staffpicks/987654')?.id).toBe('987654');
  });
});

describe('resolveVideoEmbed — unrecognized falls through to undefined', () => {
  it('empty string', () => {
    expect(resolveVideoEmbed('')).toBeUndefined();
  });

  it('malformed / non-absolute', () => {
    expect(resolveVideoEmbed('not a url at all')).toBeUndefined();
  });

  it('site-relative path', () => {
    expect(resolveVideoEmbed('/Lists/Videos/1')).toBeUndefined();
  });

  it('unrelated host', () => {
    expect(resolveVideoEmbed('https://phillipscorp.com/video.mp4')).toBeUndefined();
  });

  it('youtube host with no parseable id', () => {
    expect(resolveVideoEmbed('https://www.youtube.com/results?search_query=cats')).toBeUndefined();
  });

  it('vimeo host with no numeric segment', () => {
    expect(resolveVideoEmbed('https://vimeo.com/user/foo')).toBeUndefined();
  });

  it('shorts path on a non-youtube host is not claimed', () => {
    expect(resolveVideoEmbed('https://example.com/shorts/abc123')).toBeUndefined();
  });
});
