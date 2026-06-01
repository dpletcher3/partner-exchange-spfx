// Pure video detection + thumbnail helpers — NO SPFx imports, so this module is
// unit-testable in isolation (importing @microsoft/sp-http here would pull in
// @msinternal/ecs-flight, which jest can't resolve). The HttpClient-dependent
// Vimeo oEmbed fetch lives in ./vimeoOEmbed.
//
// detectVideo() is malformed-safe (a non-URL string yields 'other', never
// throws). YouTube thumbnails are deterministic from the id.

export type VideoKind = 'youtube' | 'vimeo' | 'other' | 'none';

export interface IDetectedVideo {
  kind: VideoKind;
  // Present for youtube/vimeo only.
  id?: string;
}

export interface IYouTubeThumbnails {
  // maxresdefault isn't present for every upload — the card's <img onError>
  // falls back to hqdefault, which always exists for public videos.
  maxRes: string;
  hq: string;
}

export function detectVideo(url: string): IDetectedVideo {
  if (!url) {
    return { kind: 'none' };
  }
  let parsed: URL;
  try {
    parsed = new URL(url);
  } catch {
    // Malformed / non-absolute URL — not parseable, so no thumbnail.
    return { kind: 'other' };
  }
  const host = parsed.hostname.toLowerCase();

  if (host.indexOf('youtube.com') !== -1) {
    const v = parsed.searchParams.get('v');
    if (v) {
      return { kind: 'youtube', id: v };
    }
    const embedMatch = parsed.pathname.match(/^\/embed\/([^/?#]+)/);
    if (embedMatch) {
      return { kind: 'youtube', id: embedMatch[1] };
    }
    return { kind: 'other' };
  }

  if (host.indexOf('youtu.be') !== -1) {
    const seg = parsed.pathname.replace(/^\/+/, '').split('/')[0];
    return seg ? { kind: 'youtube', id: seg } : { kind: 'other' };
  }

  if (host.indexOf('vimeo.com') !== -1) {
    // Public URL is vimeo.com/{id}; channel/showcase URLs surface the numeric id
    // as a path segment. Query strings (?fl=…&fe=…) live outside the pathname.
    for (const seg of parsed.pathname.split('/')) {
      if (/^\d+$/.test(seg)) {
        return { kind: 'vimeo', id: seg };
      }
    }
    return { kind: 'other' };
  }

  return { kind: 'other' };
}

export function getYouTubeThumbnails(videoId: string): IYouTubeThumbnails {
  const id = encodeURIComponent(videoId);
  return {
    maxRes: `https://img.youtube.com/vi/${id}/maxresdefault.jpg`,
    hq: `https://img.youtube.com/vi/${id}/hqdefault.jpg`
  };
}
