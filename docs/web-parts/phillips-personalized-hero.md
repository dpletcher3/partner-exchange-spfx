# I08 — Personalized Hero web part spec

Build SPFx web part that displays a dynamic time-of-day greeting with the current user's first name, on a configurable banner (color or image background). Drops on the hub home page; replaces the stock Banner / Hero placeholder.

This is **I08 from the increments backlog**, executed ahead of original sequence. See `docs/decisions.md` (new D-entry) for the resequence rationale.

## Identity

- **Name (internal):** `PhillipsPersonalizedHero`
- **Display name:** Phillips Personalized Hero
- **Description:** Time-of-day greeting with the current user's first name, on a configurable color or image banner.
- **Group in web part picker:** Phillips Custom Components (same group as Phillips News)
- **Bundled in:** the existing partner-exchange-spfx solution (no new sppkg)
- **Solution version bump:** 1.0.0.19 → 1.0.0.20

## Scope — explicit MVP

What this turn delivers:

- Color or image background, editor-toggleable
- Editor-adjustable banner height
- Hardcoded greeting wording: `Good {timeOfDay}, {firstName}` (or `Good {timeOfDay}` if no name available)
- Time-of-day computed from browser local time on page load only, with the agreed boundaries
- First name parsed from `pageContext.user.displayName`
- Editor-adjustable greeting color, size, weight, and alignment
- Current date and time display at the bottom-left of the banner, refreshing every minute (added as a refinement — see "Refresh behavior" and the date/time decision row)

What this turn does NOT deliver (explicitly deferred):

- Time Zones link
- Editable greeting wording or rich text on the greeting
- Editable greeting wording or rich text on the greeting
- Subtitle text below the greeting
- Font selection control (font is hardcoded; deferred to a future turn if needed)
- Refresh-on-time-boundary-cross (greeting is computed once on load; refresh fixes stale greetings)
- Microsoft Graph integration for richer profile data
- CTA button or links inside the banner
- Multiple background image options or carousel behavior

## Property pane

Single page, two groups.

### Group 1 — Banner

| Property | Type | Default | Notes |
|---|---|---|---|
| `backgroundType` | dropdown | `Color` | Two values: `Color`, `Image`. Toggles which downstream property is shown. |
| `backgroundColor` | color picker (full) | `#C8102E` (Phillips red) | Shown when `backgroundType` is Color. Full picker, not constrained palette — per design call. |
| `backgroundImage` | file picker (`PropertyFieldFilePicker` from `@pnp/spfx-property-controls`) | empty | Shown when `backgroundType` is Image. Three tabs — Browse (site files), Upload, From a link; the config-dependent tabs (stock/web search, organisational assets, OneDrive, recent) are hidden. The picker reports an `IFilePickerResult`; the file's `fileAbsoluteUrl` is what gets stored/read (coerced to a URL string at read time). |
| `bannerHeight` | number input | `450` | Min 200, max 600. Pixel height of the banner. |

### Group 2 — Greeting

| Property | Type | Default | Notes |
|---|---|---|---|
| `greetingColor` | color picker (full) | `#FFFFFF` (white) | Full picker. |
| `greetingSize` | dropdown | `42` (px) | Choices: `28`, `36`, `42`, `48`, `56` px. Or a slider 24–64. Implementation choice — pick whichever the SPFx PropertyPaneSlider / Dropdown matches existing repo conventions for. |
| `greetingWeight` | dropdown | `500` | Choices: `400` (regular), `500` (medium), `700` (bold). |
| `greetingAlignment` | choice group | `Left` | Three values: `Left`, `Center`, `Right`. Hardcoded to Left in v1 if the choice group adds complexity; otherwise expose. |

Note: `greetingAlignment` is borderline scope creep — the design instinct from earlier conversation was left-aligned, hardcoded. Including it in the property pane is fine if it's a small addition; if it complicates the implementation noticeably, hardcode to Left and document in deviations.

## Data sources

### User name

`this.context.pageContext.user.displayName` — already populated, no extra fetch needed.

Parse first name by splitting on whitespace and taking the first token. If `displayName` is empty, null, or not a string, fall back to no name (the greeting becomes `Good {timeOfDay}` with no comma or name).

```typescript
function extractFirstName(displayName: string | undefined): string | null {
  if (!displayName || typeof displayName !== 'string') return null;
  const first = displayName.trim().split(/\s+/)[0];
  return first || null;
}
```

### Time of day

Compute once at component mount using `new Date()`. Browser local time. Boundaries:

| Hour range (24h, local) | Greeting |
|---|---|
| 5:00 – 11:59 | Good morning |
| 12:00 – 17:59 | Good afternoon |
| 18:00 – 22:59 | Good evening |
| 23:00 – 4:59 | Good evening |

Implementation:

```typescript
function computeGreeting(date: Date): string {
  const hour = date.getHours();
  if (hour >= 5 && hour < 12) return 'Good morning';
  if (hour >= 12 && hour < 18) return 'Good afternoon';
  return 'Good evening';
}
```

No timer, no re-evaluation. Page refresh updates the greeting if a boundary has crossed since load.

## Rendering

### Layout

- `<section>` element, height = `bannerHeight` pixels, with 16px rounded corners (`--phil-radius-xl`, matching Phillips News cards) and a small horizontal margin so the rounding stays visible in a full-width section
- Background: `backgroundColor` is **always** applied as the base layer; when `backgroundType === 'Image'` and an image URL is present it is layered on top via `background-image` with `background-size: cover` and `background-position: center center`. A missing or broken image simply reveals the color underneath, so the banner is never empty at runtime
- Column layout (`flex-direction: column; justify-content: space-between`): the **greeting sits at the top-left**, the **date/time line at the bottom-left**. Horizontal alignment of both follows `greetingAlignment` (left by default) via `align-items` + `text-align`; vertical padding keeps them off the top/bottom edges

### Greeting element

- Single `<h1>` (semantically correct as a page-level heading; the hub home doesn't currently have an h1 that I know of, and this is the right place for it)
- Inline styles or CSS variables for color, font size, weight — driven by the property pane values
- Font family: hardcoded in SCSS, matches the BrandedHeader font (whichever that is — see `src/styles/_tokens.scss`)
- Text shadow: subtle dark drop shadow (`text-shadow: 0 1px 3px rgba(0,0,0,0.3)`) — gives white text legibility against any image background without needing a full scrim overlay. Lightweight, always-on, no scrim layer needed.

### Date/time element

- A `<div>` at the bottom-left of the banner showing the current date and time, e.g. `Tuesday, May 26, 2026 · 5:25 PM` — full weekday, full date, then time with no seconds. Formatted via `toLocaleDateString(undefined, { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })` and `toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' })` (the `undefined` locale uses the browser's locale), joined with ` · `. Falls back to `date.toString()` if locale data is unavailable.
- Inherits `greetingColor` (applied inline, so it always matches the greeting). Same font face and `text-shadow` as the greeting; font size fixed at **16px**, independent of `greetingSize` (which is the primary text).
- Not editor-toggleable in v1 — no property-pane control.

### Edit-mode affordance

When the web part has `backgroundType === 'Image'` but no image is selected, render a placeholder state that says "Pick a background image in the property pane" — visible only in edit mode (`this.displayMode === DisplayMode.Edit`). In read mode with no image, the fallback to `backgroundColor` kicks in and the user sees a solid color banner with the greeting.

### Refresh behavior

- **Greeting: load-only.** The time-of-day phrase and first name are computed once when the component mounts (`new Date()` + `pageContext`). A boundary crossed while the tab stays open does not update the greeting; a page refresh does.
- **Date/time: updates every minute.** A `useEffect` schedules the first tick at the next wall-clock minute boundary (`60000 - (Date.now() % 60000)` ms), then a `setInterval(…, 60000)` keeps it current so the displayed minute flips at the same instant as the real minute. Both the boundary timeout and the interval are cleared in the effect's cleanup, so nothing leaks across SharePoint single-page navigations.

## States

Simpler than Phillips News — there's no async data load:

1. **Rendered** — the only "state." Greeting computes synchronously from `pageContext` and `Date()`, banner renders.
2. **Edit-mode placeholder** — Image selected but no image picked. Visible only to editors.

No loading state, no empty state, no error state. There's nothing that can fail asynchronously.

## Accessibility

- The banner is a `<section>` with `role="banner"` and an `aria-labelledby` pointing at the greeting `<h1>` (so screen readers announce "Banner: Good morning, Dan")
- Greeting text has sufficient contrast against the chosen background — but the editor is responsible for picking colors that pass contrast checks since they have a full color picker. We don't enforce contrast in v1 (could be a future enhancement).
- Background image has `aria-hidden="true"` and no alt text — the greeting is the meaningful content; the image is decorative.
- If `displayName` is available but parsing yields no first name (edge case), the greeting falls back to the generic time-of-day phrase without an awkward "Good morning, !" rendering.

## Files

```
src/webparts/phillipsPersonalizedHero/
  PhillipsPersonalizedHeroWebPart.ts          (main web part, property pane)
  PhillipsPersonalizedHeroWebPart.manifest.json
  loc/
    en-us.js
    mystrings.d.ts
  components/
    PhillipsPersonalizedHero.tsx              (top-level component)
    PhillipsPersonalizedHero.module.scss
    Greeting.tsx                              (the h1 element with computed text)
  utils/
    timeOfDay.ts                              (computeGreeting + extractFirstName, pure functions)
    __tests__/
      timeOfDay.test.ts                       (unit tests for the boundaries and name parsing)
```

Plus updates to `config/config.json` (register the new bundle/web part) and `config/package-solution.json` (1.0.0.19 → 1.0.0.20).

## Unit tests

Lightweight, since most of the logic is in two pure functions:

### `computeGreeting`

- 00:00 → "Good evening"
- 04:59 → "Good evening"
- 05:00 → "Good morning"
- 11:59 → "Good morning"
- 12:00 → "Good afternoon"
- 17:59 → "Good afternoon"
- 18:00 → "Good evening"
- 22:59 → "Good evening"
- 23:00 → "Good evening"

Each test instantiates a `Date` at the specified hour (use `new Date(2026, 4, 26, hour, 0, 0)` for deterministic dates, not `Date.now()`).

### `extractFirstName`

- `'Dan Pletcher'` → `'Dan'`
- `'Dan'` → `'Dan'`
- `''` → `null`
- `null` → `null`
- `undefined` → `null`
- `'  Dan  Pletcher  '` (extra whitespace) → `'Dan'`
- `'Dan-Pletcher'` (hyphenated, no space) → `'Dan-Pletcher'` (or `'Dan'` depending on intent — document the choice; recommend leaving hyphenated names whole as they're a single first-name unit)

Same test-runner-or-deferred handling as Phillips News list-service turn.

## Definition of done

- Web part appears in the picker under `Phillips Custom Components`
- Drops onto a page with sensible defaults: Phillips red background, 450px tall, white greeting at 42px medium-weight, left-aligned
- Greeting renders correctly with the current user's first name at the correct time-of-day phrase
- Property pane: Banner group (backgroundType, backgroundColor or backgroundImage, bannerHeight) and Greeting group (greetingColor, greetingSize, greetingWeight, greetingAlignment if included)
- Switching `backgroundType` between Color and Image hides/shows the appropriate downstream property
- Resizing `bannerHeight` updates the rendered banner height live
- Color pickers (background and greeting) are full pickers, not constrained palettes
- Edit-mode placeholder appears when Image is selected with no image picked; runtime falls back to color
- Unit tests for `computeGreeting` and `extractFirstName` pass (or are documented as awaiting test runner)
- Console clean — no warnings, no errors
- Solution version bumped to 1.0.0.20

## Verification (browser walkthrough)

After deployment:

1. Hard-refresh the hub home page (the page where the stock Banner currently lives)
2. Add the new Phillips Personalized Hero web part below or in place of the stock Banner. Confirm it appears in the picker under Phillips Custom Components.
3. Confirm default state: Phillips red banner, 450px tall, "Good {time-of-day}, {your first name}" rendered in white at 42px
4. Open property pane, switch backgroundType to Image, pick an image from site assets, confirm it renders
5. Switch backgroundType back to Color, change `backgroundColor` to something different, confirm it updates
6. Adjust `bannerHeight` — confirm the banner height updates in real time
7. Change `greetingColor` to a contrasting color against the current background, confirm it updates
8. Change `greetingSize` and `greetingWeight`, confirm both update
9. Confirm the greeting text matches the current time of day per the boundaries above (open at 11:55am, refresh at 12:01pm, confirm it changes)
10. Remove the stock Banner web part once the new one is configured satisfactorily
11. Console clean throughout

## Polish / future-turn handoff

After this turn closes, items that may come up later:

- **Font selection** — if content authors want font variants, add a constrained-list property
- **Subtitle line** — free-form text below the greeting, if pages need that affordance
- **Boundary refresh** — re-evaluate the time-of-day on a timer if users keep tabs open across boundaries (probably not worth doing unless it's actually a problem)
- **Microsoft Graph integration** — for richer profile data like photo or job title
- **Multiple greeting styles** — "Welcome back, Dan" / "Hi, Dan" as editor-selectable variants
- **CTA button or links** — for hero-with-action use cases

None of those are scoped for this turn.

## Implementation notes (surfaced during execution, 2026-05-26)

Details that diverged from or refined the spec while building. Captured per the Definition of done.

- **`backgroundImage` uses PnP's `PropertyFieldFilePicker`** (Browse / Upload / From-a-link), added as a refinement in solution version `1.0.0.21`. The control requires a `BaseComponentContext`, and `@pnp/spfx-property-controls@3.23.0` ships its own nested `@microsoft/sp-component-base@1.22.2` — a different TypeScript type identity than the project's `1.23.0`, with no version overlap to dedupe. Rather than an `any` cast (disallowed), the `context` is cast to the picker's *own* declared context type via `type FilePickerContext = Parameters<typeof PropertyFieldFilePicker>[1]['context']` and `this.context as unknown as FilePickerContext` — the runtime object is a valid context; only the compile-time identity differs. The selected file's `fileAbsoluteUrl` is stored in the same `backgroundImage` property the component already reads, so the rendering side was unchanged. Config-dependent tabs (stock/web search, organisational assets, OneDrive, recent) are hidden; Browse (site files), Upload, and From-a-link remain. (Interim note: `1.0.0.20` shipped this as a plain URL `PropertyPaneTextField` before the file picker was wired in.)
- **`@pnp/spfx-property-controls@3.23.0` added as a dependency** for the two full color pickers. Its postinstall added `PropertyControlStrings` to `config/config.json` `localizedResources` (expected).
- **`greetingAlignment` was included** as a `PropertyPaneChoiceGroup` (Left/Center/Right), not hardcoded — it was a small addition, as the spec permitted.
- **`extractFirstName` returns `string | undefined`, not the spec's literal `null`** — to comply with the rig's `@rushstack/no-new-null` rule (same call made in the list-service turn). The parameter is typed `unknown` so a null/non-string `displayName` is handled without a null type. Semantics identical; tests assert `toBeUndefined()`.
- **Test runner already configured.** heft-jest runs the `computeGreeting`/`extractFirstName` tests as part of `npm run build`; **16 tests pass** (33 total across the solution). No deferral needed.

## Decisions captured

| Decision | Choice | Date |
|---|---|---|
| Promote I08 ahead of Phillips News polish and +Add | Yes — see new D-entry in `docs/decisions.md` | 2026-05-26 |
| MVP scope vs. full I08 original scope | MVP only — no Time Zones link (date/time display added later, see date/time row) | 2026-05-26 |
| Name source | `pageContext.user.displayName`, first-name parsed; no Microsoft Graph | 2026-05-26 |
| Time source | Browser local time, computed once on load | 2026-05-26 |
| Time-of-day boundaries | 5–11:59 morning / 12–17:59 afternoon / 18–4:59 evening | 2026-05-26 |
| Background options | Color or Image, default Color | 2026-05-26 |
| Color picker style | Full picker (background and greeting both) | 2026-05-26 |
| Banner height | Editor-adjustable, 200–600px, default 450px | 2026-05-26 |
| Greeting typography controls | Color, size, weight, alignment | 2026-05-26 |
| Font | Hardcoded; deferred | 2026-05-26 |
| Greeting wording | Hardcoded `Good {timeOfDay}, {firstName}`; not editor-editable | 2026-05-26 |
| Refresh on time-boundary cross | No — greeting is page-load only | 2026-05-26 |
| Date/time display (refinement) | Added — bottom-left, `Tuesday, May 26, 2026 · 5:25 PM`, locale-formatted, 16px, inherits greetingColor, updates every minute (minute-aligned `setInterval`, cleaned up on unmount); not editor-toggleable in v1 | 2026-05-27 |
| Background image control (refinement) | PnP `PropertyFieldFilePicker` (Browse / Upload / From-a-link), not a URL text field | 2026-05-27 |
| Banner corner radius (refinement) | 16px rounded corners (`--phil-radius-xl`, matching Phillips News cards); small horizontal margin so the rounding shows in full-width sections | 2026-05-27 |
