// Video-source parsing for Phillips Media Card. Maps raw editor-supplied
// URLs to either an embed URL (YouTube/Vimeo) or the URL itself (SharePoint
// HTML5 video). Returns undefined when YouTube/Vimeo URLs don't match the
// expected shape — the caller falls back to opening the URL in a new tab
// rather than rendering a broken iframe.

export type VideoSourceType = 'sharepoint' | 'youtube' | 'vimeo';

export interface IParsedVideo {
  type: VideoSourceType;
  // For YouTube/Vimeo: a URL embeddable via <iframe>. For SharePoint: the
  // original URL (consumed by HTML5 <video src>).
  embedUrl: string;
  // The raw video ID for YouTube and Vimeo (undefined for SharePoint).
  // Surfaced so consumers like the thumbnail resolver don't have to re-parse
  // the original URL.
  videoId?: string;
}

export function parseVideoUrl(url: string, type: VideoSourceType): IParsedVideo | undefined {
  if (!url) {
    return undefined;
  }
  switch (type) {
    case 'sharepoint':
      // No parsing — the HTML5 <video> element consumes the URL directly.
      // Includes SharePoint Site Assets, Stream classic, and any other MP4.
      return { type: 'sharepoint', embedUrl: url };
    case 'youtube': {
      const id = extractYouTubeId(url);
      if (!id) {
        return undefined;
      }
      return {
        type: 'youtube',
        embedUrl: `https://www.youtube.com/embed/${encodeURIComponent(id)}?autoplay=1`,
        videoId: id
      };
    }
    case 'vimeo': {
      const id = extractVimeoId(url);
      if (!id) {
        return undefined;
      }
      return {
        type: 'vimeo',
        embedUrl: `https://player.vimeo.com/video/${encodeURIComponent(id)}?autoplay=1`,
        videoId: id
      };
    }
    default:
      return undefined;
  }
}

// Extract video ID from any of the three accepted YouTube URL shapes:
//   - youtube.com/watch?v=ID
//   - youtu.be/ID
//   - youtube.com/embed/ID
// Hostname check tolerates www. / m. subdomains via indexOf rather than
// equality so a stray prefix doesn't reject an otherwise valid URL.
function extractYouTubeId(url: string): string | undefined {
  let parsed: URL;
  try {
    parsed = new URL(url);
  } catch {
    return undefined;
  }
  const host = parsed.hostname.toLowerCase();

  if (host.indexOf('youtube.com') !== -1) {
    const v = parsed.searchParams.get('v');
    if (v) {
      return v;
    }
    const embedMatch = parsed.pathname.match(/^\/embed\/([^/?#]+)/);
    if (embedMatch) {
      return embedMatch[1];
    }
  }

  if (host.indexOf('youtu.be') !== -1) {
    const path = parsed.pathname.replace(/^\/+/, '').split('/')[0];
    if (path) {
      return path;
    }
  }

  return undefined;
}

// Vimeo public-URL shape is just vimeo.com/{numeric-id}. Channel and showcase
// URLs (vimeo.com/channels/foo/ID, vimeo.com/showcase/.../ID) also surface the
// numeric ID as the trailing path segment — we look for the first all-digits
// path segment so those shapes work too.
function extractVimeoId(url: string): string | undefined {
  let parsed: URL;
  try {
    parsed = new URL(url);
  } catch {
    return undefined;
  }
  const host = parsed.hostname.toLowerCase();
  if (host.indexOf('vimeo.com') === -1) {
    return undefined;
  }
  for (const segment of parsed.pathname.split('/')) {
    if (/^\d+$/.test(segment)) {
      return segment;
    }
  }
  return undefined;
}
