# Phillips Media Card — Web Part Spec

## Purpose

Generic SPFx web part rendering a horizontal two-column card with header
+ description text on the left and a large image on the right. When
configured with a video URL, the image becomes a click-to-play target
opening the video in a modal player. Designed first for the 15 practice
pages on Our Culture but reusable anywhere a "featured content" callout
is needed.

## Property pane

- **Eyebrow** (text, optional, max 30 chars) — small uppercase Phillips-red
  label above the header (e.g., "PRACTICE 01")
- **Header** (text, required, max 60 chars) — main heading
- **Description** (multi-line text, required, max 300 chars)
- **Image** (SharePoint File Picker, conditionally required — see
  "Image field behavior" below)
- **Video URL** (text, optional)
- **Video source type** (dropdown, enabled only when Video URL is set) —
  `SharePoint/Stream/MP4` | `YouTube` | `Vimeo`

### Image field behavior

The Image field's label and required flag adapt to whether a thumbnail
can be auto-derived from the video URL:

| Source state | Label | Required? |
| --- | --- | --- |
| No video URL set, OR source = SharePoint/Stream/MP4 | `Image *` | Yes |
| Source = YouTube | `Image (optional — auto-derived from YouTube)` | No |
| Source = Vimeo, oEmbed succeeded | `Image (optional — auto-derived from Vimeo)` | No |
| Source = Vimeo, oEmbed failed | `Image (required — couldn't fetch Vimeo thumbnail)` plus inline error | Yes |

An uploaded image always wins over an auto-derived one (see
"Image resolution order" below), so editors can override the auto-thumbnail
when needed.

## Runtime behavior

- Two-column horizontal card. Image right (16:9 aspect ratio,
  `object-fit: cover`), text left.
- 16px rounded corners on outer container.
- Phillips brand styling — fonts, colors, spacing consistent with
  Phillips List View.
- Eyebrow when set: 12px uppercase Phillips red, letter-spacing 1px,
  small bottom margin before header.
- Header: Phillips display font, 32px desktop scaling to 24px below the
  600px breakpoint, weight 700, line-height ~1.2.
- Description: body text, neutral dark gray, line-height 1.5.
- Image without video URL: static, no overlay, no click behavior.
- Image with video URL: semi-transparent dark play-button overlay
  (60px circle, white triangle SVG centered). Hover slightly lightens.
  Click opens modal.

### Image resolution order

At render time, the card picks an image from the first match below:

1. Uploaded image (`imagePicker` property set) — always wins.
2. `vimeoThumbnailUrl` property if source type = Vimeo and the stored
   value is non-empty.
3. YouTube derived URL (`https://img.youtube.com/vi/{id}/hqdefault.jpg`)
   computed from the parsed YouTube video ID when source type = YouTube.
4. Configuration-placeholder state (the unconfigured message) when none
   of the above is available.

The unconfigured branch also triggers if header or description is empty.

### Vimeo oEmbed fetch

Vimeo thumbnails aren't deterministic from the URL, so we resolve them
at edit time via Vimeo's oEmbed endpoint:

- **Endpoint:** `https://vimeo.com/api/oembed.json?url={encodedVimeoUrl}`
- **Trigger:** property pane field change on Video URL or Video source
  type, when the resulting (URL, type) pair is (non-empty, Vimeo).
- **Response:** JSON `{ thumbnail_url: "...", ... }`. We persist only
  `thumbnail_url` into the `vimeoThumbnailUrl` web part property.
- **Stale-fetch guard:** an incrementing token on the WebPart class is
  captured per call; if the editor changes the URL while a fetch is in
  flight, the stale response is discarded so a newer URL's thumbnail
  isn't overwritten by an older URL's late response.
- **Failure modes** (all clear `vimeoThumbnailUrl` and surface the
  inline error described in "Image field behavior"):
  - Network / CORS / timeout / fetch reject
  - Non-2xx response (404 for private or deleted videos)
  - Response missing `thumbnail_url`
- The fetch is **not** retried automatically — the editor's next URL
  edit triggers a fresh attempt.

### YouTube auto-thumbnail

YouTube thumbnails are derived synchronously from the video ID
(`https://img.youtube.com/vi/{id}/hqdefault.jpg`). `hqdefault` is
always served for public videos at 480×360, which fits our 16:9 card
slot cleanly. No network call at edit time; no stored property.

## Modal player

- Centered overlay with semi-transparent black backdrop.
- Close affordances: X in top-right, ESC key, backdrop click.
- Player auto-pauses on close (video element pause() / iframe src cleared).
- Max width 1000px, max height 80vh.
- Backdrop has `aria-modal="true"` and focus trap on the close button.

### Video source handling

- **SharePoint/Stream/MP4:** HTML5 `<video controls autoplay>` element
  with URL as src.
- **YouTube:** parse video ID from URLs of shape `youtube.com/watch?v=ID`,
  `youtu.be/ID`, or `youtube.com/embed/ID`. Render iframe with
  `https://www.youtube.com/embed/{id}?autoplay=1`.
- **Vimeo:** parse video ID from URLs of shape `vimeo.com/ID`. Render
  iframe with `https://player.vimeo.com/video/{id}?autoplay=1`.
- If URL parsing fails for YouTube/Vimeo: fall back to opening the URL
  in a new tab rather than the modal.

## Responsive

- Below 600px viewport width: stack columns vertically — image on top,
  text below. Card stays rounded; horizontal paddings adapt to ~16px.
- Above 600px: two-column horizontal layout, left column ~42% width,
  right column ~58% width.

## Property pane validation

- Header and description required.
- Image required per the "Image field behavior" matrix above.
- Video URL: when set, video source type required. Basic URL format check.
- Eyebrow optional; when empty, no extra spacing reserved above header.
- maxLength on text fields enforced at property pane level.

## Explicit non-features

- No CTA button (image click is the only action)
- No layout direction toggle (image always right at desktop widths)
- No video autoplay outside the modal
- No video looping or playlist support
- No video source types beyond the three listed
- No accessibility audit beyond standard HTML semantics

## Solution version

Bump from 1.0.2.0 to 1.0.3.0.

## Definition of done

- [ ] Web part builds without errors
- [ ] Web part deployed to tenant App Catalog at v1.0.3.0
- [ ] Appears in "Phillips Custom Components" group in page toolbox
- [ ] Property pane renders correctly; required-field validation works
- [ ] Static image card (no video URL) renders correctly
- [ ] Play overlay appears when video URL is set
- [ ] Modal opens, plays video, closes cleanly for all three source types
- [ ] Responsive stack at narrow widths
- [ ] No console errors
- [ ] Human visual verification complete
