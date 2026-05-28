// Thumbnail resolution helpers. YouTube thumbnails are deterministic from the
// video ID, so the YouTube helper is synchronous. Vimeo thumbnails require an
// oEmbed call at edit time — the result is persisted into a saved property by
// the WebPart so the runtime render doesn't trigger a network round-trip.

export function getYouTubeThumbnail(videoId: string): string {
  // hqdefault.jpg is always available for public videos at 480x360 — the
  // right footprint for our 16:9 card image at typical desktop widths.
  // Alternative sizes (mqdefault, sddefault, maxresdefault) sometimes 404 for
  // older or low-resolution uploads, so we stick with hqdefault.
  return `https://img.youtube.com/vi/${encodeURIComponent(videoId)}/hqdefault.jpg`;
}

interface IVimeoOEmbedResponse {
  thumbnail_url?: unknown;
}

// Fetches Vimeo's oEmbed metadata for a video URL and returns the
// thumbnail URL on success. Returns null for every failure mode (network
// failure, non-2xx response, missing field, JSON parse error) so the caller
// has a single boolean check rather than a try/catch ladder. Vimeo's oEmbed
// endpoint serves CORS headers, so this works directly from the browser.
export async function getVimeoThumbnail(videoUrl: string): Promise<string | null> {
  if (!videoUrl) {
    return null;
  }
  try {
    const oembedUrl = `https://vimeo.com/api/oembed.json?url=${encodeURIComponent(videoUrl)}`;
    const response = await fetch(oembedUrl);
    if (!response.ok) {
      return null;
    }
    const data = (await response.json()) as IVimeoOEmbedResponse;
    if (typeof data.thumbnail_url === 'string' && data.thumbnail_url) {
      return data.thumbnail_url;
    }
    return null;
  } catch {
    // Network failure, CORS rejection (unlikely from Vimeo but defensive),
    // JSON parse error — all collapse to "we don't have a thumbnail."
    return null;
  }
}
