// Provider dispatch for the Highlight Video player: turns a raw list-supplied URL
// into the iframe src to embed, for Vimeo OR YouTube.
//
// Pure — NO SPFx imports — so this module is unit-testable in isolation (the same
// constraint videoThumbnail.ts documents: importing @microsoft/sp-http pulls in
// @msinternal/ecs-flight, which jest can't resolve).
//
// Deliberately a THIN WRAPPER over the Media Card Gallery's detectVideo rather
// than a second parser (D035 / spec §3: reuse, don't re-implement). detectVideo
// already recognizes vimeo.com, youtube.com/watch?v=, youtu.be/{id} and
// youtube.com/embed/{id}. Only two things are added here:
//   1. /shorts/{id}, which detectVideo classifies as 'other';
//   2. the embed-src construction + provider label, which are player concerns
//      and belong to this web part, not to the gallery's thumbnail helper.
//
// Malformed-safe by inheritance: detectVideo never throws, and the shorts probe
// below is wrapped in its own try/catch.

import { detectVideo } from '../../phillipsMediaGallery/services/videoThumbnail';

export type EmbedProvider = 'vimeo' | 'youtube';

export interface IVideoEmbed {
  provider: EmbedProvider;
  id: string;
  // The iframe src.
  embedSrc: string;
  // Human label for the "Open in …" fallback link.
  providerLabel: string;
}

// youtube.com/shorts/{id} — detectVideo returns 'other' for this shape, so it is
// parsed here. Any other host, or a shorts path with no id segment, yields ''.
function parseYouTubeShortsId(url: string): string {
  if (!url) {
    return '';
  }
  let parsed: URL;
  try {
    parsed = new URL(url);
  } catch {
    return '';
  }
  if (parsed.hostname.toLowerCase().indexOf('youtube.com') === -1) {
    return '';
  }
  const match = parsed.pathname.match(/^\/shorts\/([^/?#]+)/);
  return match ? match[1] : '';
}

// Returns the embed descriptor for a recognized Vimeo/YouTube URL, or undefined
// for anything else (empty, malformed, or a non-video host). Callers treat
// undefined exactly as they previously treated an unparseable Vimeo URL.
export function resolveVideoEmbed(url: string): IVideoEmbed | undefined {
  const detected = detectVideo(url);

  if (detected.kind === 'vimeo' && detected.id) {
    return {
      provider: 'vimeo',
      id: detected.id,
      // Unchanged from the pre-YouTube implementation.
      embedSrc: `https://player.vimeo.com/video/${detected.id}`,
      providerLabel: 'Vimeo'
    };
  }

  if (detected.kind === 'youtube' && detected.id) {
    return youTubeEmbed(detected.id);
  }

  // detectVideo doesn't know /shorts/ — check before giving up.
  const shortsId = parseYouTubeShortsId(url);
  if (shortsId) {
    return youTubeEmbed(shortsId);
  }

  return undefined;
}

function youTubeEmbed(id: string): IVideoEmbed {
  return {
    provider: 'youtube',
    id,
    embedSrc: `https://www.youtube.com/embed/${id}`,
    providerLabel: 'YouTube'
  };
}
