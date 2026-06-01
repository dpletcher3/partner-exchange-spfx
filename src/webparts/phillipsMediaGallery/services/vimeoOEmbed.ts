import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';

// Vimeo oEmbed thumbnail resolution via SPFx HttpClient, cached in sessionStorage
// so a given video is fetched at most once per session (instead of on every page
// navigation). Kept apart from videoThumbnail.ts so the pure helpers there stay
// unit-testable (importing @microsoft/sp-http pulls in @msinternal/ecs-flight,
// which jest can't resolve).

const LOG = '[MediaGallery]';

// Namespaced + versioned key per video. Bump the `v1` segment to bust every
// cached thumbnail (e.g. if the resolution strategy changes).
const KEY_PREFIX = 'phil-mg:vimeo-thumb:v1:';

// Negative sentinel: a video whose oEmbed failed this session. Stored so a
// known-failing video (e.g. CSP-blocked) isn't re-fetched on every render;
// sessionStorage clears at session end, so it retries next session.
const NEG = '__neg__';

interface IVimeoOEmbedResponse {
  thumbnail_url?: unknown;
}

type CacheResult = { state: 'hit'; url: string } | { state: 'negative' } | { state: 'miss' };

// In-memory fallback used ONLY when sessionStorage is unavailable (private mode
// / storage disabled). Keyed by videoId; value undefined = negative.
const memFallback = new Map<string, string | undefined>();

function cacheGet(videoId: string): CacheResult {
  const key = KEY_PREFIX + videoId;
  try {
    const v = window.sessionStorage.getItem(key);
    if (v === null) {
      return { state: 'miss' };
    }
    return v === NEG ? { state: 'negative' } : { state: 'hit', url: v };
  } catch {
    if (memFallback.has(videoId)) {
      const m = memFallback.get(videoId);
      return m === undefined ? { state: 'negative' } : { state: 'hit', url: m };
    }
    return { state: 'miss' };
  }
}

function cacheSet(videoId: string, url: string | undefined): void {
  const key = KEY_PREFIX + videoId;
  try {
    window.sessionStorage.setItem(key, url === undefined ? NEG : url);
  } catch {
    memFallback.set(videoId, url);
  }
}

// Returns the thumbnail URL, or undefined on any failure (caller → placeholder).
// Reads sessionStorage first; only fetches on a cache miss. Logs the per-card
// cache outcome (cache-hit / network-fetch / negative-hit) alongside the loud
// failure warnings.
export async function getVimeoThumbnail(
  videoUrl: string,
  videoId: string,
  httpClient: HttpClient
): Promise<string | undefined> {
  if (!videoUrl || !videoId) {
    return undefined;
  }

  const cached = cacheGet(videoId);
  if (cached.state === 'hit') {
    console.log(`${LOG} vimeo ${videoId}: cache-hit`);
    return cached.url;
  }
  if (cached.state === 'negative') {
    console.log(`${LOG} vimeo ${videoId}: negative-hit (known-failing this session)`);
    return undefined;
  }

  // Miss → network fetch, then cache the outcome (URL or negative sentinel).
  const url = await fetchOEmbed(videoUrl, httpClient);
  cacheSet(videoId, url);
  console.log(
    `${LOG} vimeo ${videoId}: network-fetch (${url ? 'resolved → cached' : 'failed → negative cached'})`
  );
  return url;
}

async function fetchOEmbed(videoUrl: string, httpClient: HttpClient): Promise<string | undefined> {
  try {
    const oembedUrl = `https://vimeo.com/api/oembed.json?url=${encodeURIComponent(videoUrl)}`;
    const response: HttpClientResponse = await httpClient.get(oembedUrl, HttpClient.configurations.v1);
    if (!response.ok) {
      console.warn(`${LOG} Vimeo oEmbed non-OK (${response.status}) for ${videoUrl}`);
      return undefined;
    }
    const data = (await response.json()) as IVimeoOEmbedResponse;
    if (typeof data.thumbnail_url === 'string' && data.thumbnail_url) {
      return data.thumbnail_url;
    }
    console.warn(`${LOG} Vimeo oEmbed missing thumbnail_url for ${videoUrl}`, data);
    return undefined;
  } catch (err) {
    // Most likely a tenant CSP block on vimeo.com (connect-src). Loud so the
    // operator knows to allow it or map a main-image override.
    console.warn(`${LOG} Vimeo oEmbed fetch FAILED (CSP?) for ${videoUrl}`, err);
    return undefined;
  }
}
