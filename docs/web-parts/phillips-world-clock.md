# Phillips World Clock — Web Part Spec

## Purpose

Multi-timezone clock with editor-customizable titles per tile. Built
to replace SharePoint's stock World Clock on the hub home page with
a smaller-tile, brand-consistent design that allows custom display
titles independent of the underlying timezone.

## Property pane

- **Section title** (text, optional) — header above the clock grid
- **Clocks** (reorderable list, 1–12 entries):
    - **Title** (text, required) — free-form display text per tile
    - **Timezone** (text, required) — IANA timezone string
      (validated via `Intl.DateTimeFormat`)

Inline help text on the Timezone field links to the Wikipedia
tz database list for editor reference.

## Runtime behavior

- Renders tiles in CSS Grid with `auto-fit` columns and a 180px
  minimum, producing roughly 5 tiles per row at typical canvas
  widths
- Each tile: 16px rounded corners, Phillips brand styling
- Tile content: title (top), time (large, 12-hour with AM/PM),
  date (M/D format, small)
- Live updates every 60 seconds via setInterval, cleaned up on
  unmount
- Time and date both computed via `Intl.DateTimeFormat` with the
  configured timezone — DST handled correctly by the browser

## Visual specs

- Tile width: 180px minimum, grows in grid track
- Tile padding: 12px
- Title font: 14px, semibold
- Time font: 28px display
- AM/PM font: 12px alongside time
- Date font: 12px, neutral gray
- Container: 16px rounded corners, matches other Phillips web parts
- Section title (if set): red underline, matches other Phillips
  section titles

## Responsive

- Grid auto-wraps via `auto-fit` + `minmax(180px, 1fr)`
- At narrow widths, tiles stack to single column naturally

## Non-features

- No city autocomplete or geocoding
- No "hours ahead/behind" delta
- No seconds, no 24-hour toggle
- No per-tile styling overrides
- No alarms or notifications
- No user-relative timezone awareness

## Solution version

Bump from 1.0.3.0 to 1.0.4.0.

## Definition of done

- [ ] Web part builds and deploys at v1.0.4.0
- [ ] Appears in "Phillips Custom Components" toolbox group
- [ ] Property pane renders correctly with validation
- [ ] Invalid IANA timezone strings rejected with inline error
- [ ] Tiles render at 5-across desktop width
- [ ] Times match stock World Clock for same locations
- [ ] Date renders as "M/D"
- [ ] Time updates across minute boundary
- [ ] Responsive stacking at narrow widths
- [ ] No console errors
- [ ] Human visual verification complete
