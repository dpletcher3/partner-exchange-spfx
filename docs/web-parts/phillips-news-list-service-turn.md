# Phillips News — list-service turn spec

Hardening pass on `NewsRepositoryService.ts` and the shape-extraction helpers it depends on. This is the second turn in the build sequence (`scaffold → list service → cards → polish → +Add → deploy`). Scaffold turn closed in commits `2deb7a1`, `739ee57`, `206adc1`, `cdc2d8b`.

The turn has two goals:

1. **Close scaffold deviation #2** — resolve `Reserved_ImageAttachment` to a usable image URL so real thumbnails render in cards. Falls back gracefully to the existing red placeholder when extraction yields nothing.
2. **Lock in the defensive extractors with unit tests** — `extractUrl`, `extractChoices`, and the new image extractor become testable pure functions with fixtures covering every shape variant we've seen and a couple we haven't.

Explicitly **out of scope** for this turn: retry/backoff, in-memory caching, telemetry/logging beyond the existing `console.warn`. Each of those is reasonable future work but none has a confirmed need yet.

## Runtime shape — confirmed

Diagnostic run against the sandbox list on 2026-05-26. The runtime shape is:

**`ThumbnailImage` field value** (a JSON string of an object):

```json
"{\"fileName\":\"Reserved_ImageAttachment_[14]_[ThumbnailImage][32]_[GUID][1]_[1].png\",\"originalImageName\":\"<user-friendly name>\"}"
```

There is **no URL in the `ThumbnailImage` value itself**. The actual image file is stored as a list-item attachment.

**`AttachmentFiles`** (available via `$expand`, one entry per attachment on the item):

```json
[
  {
    "FileName": "Reserved_ImageAttachment_[14]_[ThumbnailImage][32]_[GUID][1]_[1].png",
    "ServerRelativeUrl": "/sites/PartnerExchange-Sandbox/Lists/News Repository/Attachments/{itemId}/Reserved_ImageAttachment_[...].png"
  }
]
```

The `FileName` in an `AttachmentFiles` entry matches the `fileName` in the `ThumbnailImage` JSON string exactly. That match is the link between the two.

### Approach

Single REST call. Add `AttachmentFiles` to the `$expand`, return both the field and the attachment array in one response. Extractor matches by filename and returns the matching `ServerRelativeUrl`.

No per-item N+1 calls. No URL construction from convention (the API tells us the path directly).

### Other shape notes from the same diagnostic

Confirmed behavior under `Accept: application/json;odata=nometadata` from the live browser context:

- `LinkUrl` returns as `{ Description, Url }` — the structured object, not an empty string. The earlier "empty string" symptom Claude Code hit during scaffold likely came from a different metadata level or `$select` interaction. The `extractUrl` defensive helper handles both shapes correctly and stays.
- `Category` (MultiChoice) returns as a plain array (`["Phillips In The News"]`), not wrapped in `{ results: ... }`. The `extractChoices` helper handles both unwrapped and wrapped shapes and stays.

No behavioral changes needed for those extractors. They are confirmed working; the unit tests just lock the behavior in.

## Scope: image URL extraction

### `extractThumbnailUrl(rawThumbnail: unknown, attachmentFiles: unknown): string | null`

Pure-function extractor sitting alongside `extractUrl` and `extractChoices`. Contract:

- **Input:**
  - `rawThumbnail` — the `ThumbnailImage` field value as returned by SP REST (a JSON string, or null when no thumbnail set)
  - `attachmentFiles` — the `AttachmentFiles` array from the item (expanded in the REST query)
- **Output:** a server-relative URL string usable as `<img src>`, or `null` when no usable URL can be derived.
- **Behavior:**
  - If `rawThumbnail` is null/undefined/empty string, return `null` (item has no thumbnail set — render the red fallback). **No warning** — this is a legitimate empty state, not a shape problem.
  - Parse `rawThumbnail` as JSON. If parse fails, log `[PhillipsNews] Failed to parse ThumbnailImage` with the raw value and return `null`.
  - Extract `fileName` from the parsed object. If absent, log `[PhillipsNews] Unexpected ThumbnailImage shape` with the parsed object and return `null`.
  - Walk `attachmentFiles`. Find the entry whose `FileName` equals the extracted `fileName`. Return its `ServerRelativeUrl`.
  - If no match, log `[PhillipsNews] ThumbnailImage filename not found in AttachmentFiles` with the filename and return `null`.

The function is intentionally permissive on the way in and conservative on the way out — better to fall back to the red placeholder than to render a broken `<img>`.

### Service query update

The `getNewsItems` REST URL needs two changes:

1. Add `AttachmentFiles` to the `$expand`
2. Add `AttachmentFiles/FileName,AttachmentFiles/ServerRelativeUrl` to the `$select` so the response stays lean (otherwise SharePoint serializes the full attachment metadata for each one)

New URL shape:

```
{sourceSiteUrl}/_api/web/lists/getbytitle('{listTitle}')/items
  ?$select=Id,Title,Category,ItemType,LinkUrl,ThumbnailImage,ShortDescription,PublishedDate,AttachmentFiles/FileName,AttachmentFiles/ServerRelativeUrl
  &$expand=AttachmentFiles
  &$filter=...        (built from itemTypeFilter; categoryFilter still client-side)
  &$orderby=PublishedDate desc
  &$top={maxItems}
```

Pass both the raw `ThumbnailImage` value and `AttachmentFiles` to `extractThumbnailUrl` in the item mapper:

```typescript
thumbnailImageUrl: extractThumbnailUrl(raw.ThumbnailImage, raw.AttachmentFiles),
```

The model field `thumbnailImageUrl: string | null` already exists from the scaffold. The card component renders the red fallback when `thumbnailImageUrl` is null, otherwise renders an `<img>`.

### NewsCard rendering update

After this turn, the card should:

- If `thumbnailImageUrl` is set, render `<img src={thumbnailImageUrl} alt={item.title} />` constrained to the 4:3 aspect with `object-fit: cover` so any image dimension fits cleanly.
- If `thumbnailImageUrl` is null, render the existing red fallback unchanged.
- If `<img>` fails to load at runtime (broken URL, deleted attachment), gracefully fall back to the red placeholder via the `onError` handler.

Alt text uses the item title. The `ThumbnailImage` value has an `originalImageName` field but it's a filename, not descriptive alt text — using the title is more useful semantically.

The `ServerRelativeUrl` from SharePoint is properly URL-encoded by the browser when consumed in an `<img src>`. No manual encoding needed in the extractor — the brackets and spaces in `Reserved_ImageAttachment_[...]` filenames work as-is.

## Scope: unit tests on the extractors

New file: `src/webparts/phillipsNews/services/__tests__/extractors.test.ts` (or wherever the SPFx repo conventionally puts tests — match existing patterns).

### Fixtures

A shared `__tests__/fixtures/raw-list-items.ts` module exports realistic raw responses captured from the live REST endpoint, plus crafted edge cases.

**For `extractUrl`:**

- `urlAsObject` — `{ Url: 'https://phillipscorp.com/article', Description: 'Article' }`
- `urlAsString` — `'https://phillipscorp.com/article'`
- `urlEmpty` — `''`
- `urlNull` — `null`
- `urlObjectWithoutUrl` — `{ Description: 'no url here' }` (warning case)

**For `extractChoices`:**

- `categoriesAsArray` — `['Phillips Loop', 'Phillips In The News']`
- `categoriesAsWrapped` — `{ results: ['Phillips Loop', 'Phillips In The News'] }`
- `categoriesAsString` — `'Phillips Loop'` (single-choice column accidentally)
- `categoriesNull` — `null`

**For `extractThumbnailUrl`:**

- `thumbnailValidWithMatch` — `rawThumbnail` is a valid JSON string with a `fileName` that matches one entry in `attachmentFiles` → returns that entry's `ServerRelativeUrl`
- `thumbnailNullField` — `rawThumbnail: null`, `attachmentFiles: []` → returns null, no warning
- `thumbnailEmptyStringField` — `rawThumbnail: ''`, `attachmentFiles: []` → returns null, no warning
- `thumbnailFilenameNotInAttachments` — valid `rawThumbnail` but no matching `FileName` in `attachmentFiles` → returns null, filename-not-found warning logged
- `thumbnailMalformedJson` — `rawThumbnail: 'not json'`, `attachmentFiles: []` → returns null, parse-fail warning logged
- `thumbnailMissingFileNameKey` — `rawThumbnail: '{"otherKey":"x"}'`, `attachmentFiles: []` → returns null, unexpected-shape warning logged
- `thumbnailAttachmentFilesUndefined` — valid `rawThumbnail`, `attachmentFiles: undefined` → returns null, filename-not-found warning logged

### Test cases

Each extractor gets:

1. **Happy path** — the most common shape returns the expected value.
2. **Alternative shape** — the variant we discovered (e.g., wrapped `{ results: }` for choices).
3. **Null/undefined input** — returns the documented empty value, no exception, no warning when the input is legitimately empty.
4. **Malformed input** — returns the documented empty value, logs a warning.
5. **Empty content** — input present but with no usable data (empty string, empty array) returns the documented empty value, no warning (it's not an unexpected shape — it's a legitimately empty field).

Warnings are tested by spying on `console.warn` and asserting it was called with the expected message prefix and that the raw shape is included in the log payload.

### Test runner

Use whatever Jest/Mocha/test runner the SPFx repo is already configured for. If none is configured yet, this turn does not introduce one — instead, the tests live in the file but are documented as "to be wired into the runner during a future testing-infrastructure turn." Claude Code should note this state if it applies.

## Scope: NewsRepositoryService no other behavioral changes

Specifically NOT changed in this turn:

- No retry logic
- No backoff
- No in-memory caching
- No telemetry hooks
- No alternative transports (e.g., PnPjs)
- No throttling-aware request pacing

Each of these is a reasonable future addition when there is concrete evidence of need. None has that evidence today.

## Files touched

```
src/webparts/phillipsNews/services/
  NewsRepositoryService.ts          (query update: $select + $expand; mapping update)
  extractors.ts                     (new file if extractors moved out of the service file; otherwise inline)
  __tests__/
    extractors.test.ts              (new)
    fixtures/
      raw-list-items.ts             (new)

src/webparts/phillipsNews/components/
  NewsCard.tsx                      (img element + onError fallback)
  NewsCard.module.scss              (img sizing rules — aspect-ratio + object-fit cover)
```

No changes to `PhillipsNewsWebPart.ts`, no changes to manifest. Solution version goes 1.0.0.18 → 1.0.0.19 at deploy time.

## Definition of done

- `extractThumbnailUrl` exists, accepts `(rawThumbnail, attachmentFiles)`, returns the matching `ServerRelativeUrl` or null
- Service query expanded to include `AttachmentFiles` with a focused `$select`
- Service mapper passes both values to the extractor and populates `thumbnailImageUrl` on the model
- NewsCard renders real `<img>` when the URL is present, red fallback when null, red fallback when the image fails to load via `onError`
- All three extractors have unit tests for happy/alternative/null/malformed/empty cases
- Tests pass locally (or are documented as awaiting test-runner configuration)
- Deployed bundle visually shows real photos on the seeded News Repository items that have thumbnails set
- Console clean on the happy path — no `[PhillipsNews]` warnings unless a shape actually fails extraction
- This doc updated if implementation surfaces a shape detail not captured here

## Verification (browser walkthrough)

After deployment:

1. Hard-refresh the hub home page
2. Confirm cards now show actual photos for items that have `ThumbnailImage` set on the list (all 5 seeded items do)
3. Temporarily clear the ThumbnailImage on one item via the SP UI → confirm that item's card now shows the red fallback while the others stay populated
4. Restore the thumbnail on that item
5. Confirm `onError` fallback by manually editing one item's ThumbnailImage field value via the SP UI to reference a non-existent filename, then reloading — the card should render red, not a broken-image icon. Restore afterward.
6. Console clean on the happy path
7. Click a card with a real image; tab opens to the article as before — image work shouldn't have touched the link path, but confirm it didn't regress

## Polish-turn handoff

After this turn closes, the polish turn picks up with real photos rendering. That changes which polish items are most valuable:

- "Soften the no-thumbnail fallback" becomes more relevant — you can see which items lack thumbnails and what the red blocks look like next to real photos
- "Bump grid minmax floor to ~240px" can be evaluated against real visual density rather than an all-red placeholder grid
- Hover treatment can be designed against real photo content

The polish-turn backlog in `phillips-news.md` doesn't need updating from this turn unless something new is discovered.

## Implementation notes (surfaced during execution, 2026-05-26)

Details that diverged from or refined the spec while building. Captured here per the Definition of done ("This doc updated if implementation surfaces a shape detail not captured here").

- **Model field rename, not reuse.** The spec assumed `thumbnailImageUrl: string | null` "already exists from the scaffold." It didn't — the scaffold modeled the thumbnail as `thumbnail?: INewsThumbnail` (an object). This turn replaced that with `thumbnailImageUrl?: string` and removed the now-unused `INewsThumbnail` interface and `parseThumbnail` helper.
- **`undefined`, not `null`, as the "no URL" sentinel.** `extractThumbnailUrl` returns `string | undefined` (and the model field is `thumbnailImageUrl?: string`) to comply with the rig's `@rushstack/no-new-null` lint rule. Semantics are identical to the spec's `null`; tests assert `toBeUndefined()`. (Note: the rule only flags `null` in *type* positions, not value positions — so the test fixtures still use real `null` literals to model SP REST payloads.)
- **Extractors live in `extractors.ts`** as exported pure functions (`extractUrl`, `extractChoices`, `extractThumbnailUrl`). The scaffold's `toCategoryArray` was unified into `extractChoices` (identical shape-handling: array / `{ results }` / bare string), so the item `Category` value and the field `Choices` collection now go through one tested helper.
- **`AttachmentFiles` handled defensively for both shapes.** The extractor accepts a plain array (minimal metadata) and a `{ results: [...] }` wrapper (verbose), consistent with this turn's central lesson about runtime-vs-CLI shape drift.
- **Test runner was already configured.** heft-jest is wired via the `@microsoft/spfx-web-build-rig` (`testMatch: lib-commonjs/**/*.test.js`), so the unit tests run as part of `npm run build` (the `heft test` step). **17 tests pass**; this turn did NOT need to defer them to a future testing-infrastructure turn.
- **Live-data check.** Running the new `$expand=AttachmentFiles` query against the sandbox list confirmed all 5 seeded items resolve their `ThumbnailImage` fileName to a real `ServerRelativeUrl` (one attachment each, `/sites/PartnerExchange-Sandbox/Lists/News Repository/Attachments/{itemId}/...`).

## Decisions captured

| Decision | Choice | Date |
|---|---|---|
| MVP scope for image URL resolution | Yes — extract + render, defer srcset/resize variants | 2026-05-26 |
| Path for thumbnail URL extraction | `$expand=AttachmentFiles`, match by filename, return `ServerRelativeUrl` | 2026-05-26 |
| Unit tests for defensive extractors | Yes, all three (`extractUrl`, `extractChoices`, `extractThumbnailUrl`) | 2026-05-26 |
| Retry/backoff | Punt — no evidence of need | 2026-05-26 |
| Caching | Punt — no evidence of need | 2026-05-26 |
| Telemetry | Punt — console.warn is sufficient for v1 | 2026-05-26 |
| Runtime shape diagnostic before implementation | Done; results captured in "Runtime shape — confirmed" above | 2026-05-26 |
