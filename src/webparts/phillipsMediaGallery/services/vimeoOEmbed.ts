import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';

// Vimeo oEmbed thumbnail fetch via SPFx HttpClient, cached per URL. Kept apart
// from videoThumbnail.ts so the pure helpers there stay unit-testable (importing
// @microsoft/sp-http pulls in @msinternal/ecs-flight, which jest can't resolve).

const LOG = '[MediaGallery]';

interface IVimeoOEmbedResponse {
  thumbnail_url?: unknown;
}

// Per-URL cache (module-level so it survives re-renders and is shared across
// cards on the page). A resolved miss is cached too, so a CSP-blocked tenant
// doesn't re-hit Vimeo on every render.
const vimeoCache = new Map<string, string | undefined>();

// Returns the thumbnail URL on success, or undefined for every failure mode
// (CSP block, network, non-2xx, missing field, parse error) — caller falls back
// to the placeholder. Failures warn (loud, unlike the @pnp silent catch).
export async function getVimeoThumbnail(
  videoUrl: string,
  httpClient: HttpClient
): Promise<string | undefined> {
  if (!videoUrl) {
    return undefined;
  }
  if (vimeoCache.has(videoUrl)) {
    return vimeoCache.get(videoUrl);
  }

  try {
    const oembedUrl = `https://vimeo.com/api/oembed.json?url=${encodeURIComponent(videoUrl)}`;
    const response: HttpClientResponse = await httpClient.get(oembedUrl, HttpClient.configurations.v1);
    if (!response.ok) {
      console.warn(`${LOG} Vimeo oEmbed non-OK (${response.status}) for ${videoUrl}`);
      vimeoCache.set(videoUrl, undefined);
      return undefined;
    }
    const data = (await response.json()) as IVimeoOEmbedResponse;
    if (typeof data.thumbnail_url === 'string' && data.thumbnail_url) {
      vimeoCache.set(videoUrl, data.thumbnail_url);
      return data.thumbnail_url;
    }
    console.warn(`${LOG} Vimeo oEmbed missing thumbnail_url for ${videoUrl}`, data);
    vimeoCache.set(videoUrl, undefined);
    return undefined;
  } catch (err) {
    // Most likely a tenant CSP block on vimeo.com (connect-src). Loud so the
    // operator knows to allow it or map a main-image override.
    console.warn(`${LOG} Vimeo oEmbed fetch FAILED (CSP?) for ${videoUrl}`, err);
    vimeoCache.set(videoUrl, undefined);
    return undefined;
  }
}
