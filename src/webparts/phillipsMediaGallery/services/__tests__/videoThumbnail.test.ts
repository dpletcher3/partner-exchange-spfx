import { detectVideo, getYouTubeThumbnails } from '../videoThumbnail';

describe('detectVideo — YouTube', () => {
  it('watch?v= form', () => {
    expect(detectVideo('https://www.youtube.com/watch?v=dQw4w9WgXcQ')).toEqual({
      kind: 'youtube',
      id: 'dQw4w9WgXcQ'
    });
  });

  it('youtu.be short form', () => {
    expect(detectVideo('https://youtu.be/dQw4w9WgXcQ')).toEqual({
      kind: 'youtube',
      id: 'dQw4w9WgXcQ'
    });
  });

  it('/embed/ form', () => {
    expect(detectVideo('https://www.youtube.com/embed/dQw4w9WgXcQ?rel=0')).toEqual({
      kind: 'youtube',
      id: 'dQw4w9WgXcQ'
    });
  });

  it('m. subdomain still detected', () => {
    expect(detectVideo('https://m.youtube.com/watch?v=abc123')).toEqual({
      kind: 'youtube',
      id: 'abc123'
    });
  });

  it('youtube host without a parseable id → other', () => {
    expect(detectVideo('https://www.youtube.com/results?search_query=cats')).toEqual({
      kind: 'other'
    });
  });
});

describe('detectVideo — Vimeo', () => {
  it('plain vimeo.com/{id}', () => {
    expect(detectVideo('https://vimeo.com/1165473810')).toEqual({ kind: 'vimeo', id: '1165473810' });
  });

  it('vimeo.com/{id} with query string (15 Practices shape)', () => {
    expect(detectVideo('https://vimeo.com/1165473810?fl=ip&fe=ec')).toEqual({
      kind: 'vimeo',
      id: '1165473810'
    });
  });

  it('channel/showcase URL surfaces the numeric id', () => {
    expect(detectVideo('https://vimeo.com/channels/staffpicks/987654')).toEqual({
      kind: 'vimeo',
      id: '987654'
    });
  });

  it('vimeo host without a numeric segment → other', () => {
    expect(detectVideo('https://vimeo.com/user/foo')).toEqual({ kind: 'other' });
  });
});

describe('detectVideo — edge / malformed input', () => {
  it('empty string → none', () => {
    expect(detectVideo('')).toEqual({ kind: 'none' });
  });

  it('non-URL string → other (never throws)', () => {
    expect(detectVideo('not a url at all')).toEqual({ kind: 'other' });
  });

  it('relative path (not absolute) → other', () => {
    expect(detectVideo('/Lists/Videos/1')).toEqual({ kind: 'other' });
  });

  it('non-video absolute URL → other', () => {
    expect(detectVideo('https://phillipscorp.com/video.mp4')).toEqual({ kind: 'other' });
  });
});

describe('getYouTubeThumbnails', () => {
  it('returns maxres + hq urls for the id', () => {
    expect(getYouTubeThumbnails('dQw4w9WgXcQ')).toEqual({
      maxRes: 'https://img.youtube.com/vi/dQw4w9WgXcQ/maxresdefault.jpg',
      hq: 'https://img.youtube.com/vi/dQw4w9WgXcQ/hqdefault.jpg'
    });
  });
});
