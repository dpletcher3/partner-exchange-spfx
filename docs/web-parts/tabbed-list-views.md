# Tabbed List Views — Web Part Spec

## Purpose

Generic SPFx web part that renders a SharePoint list with 2–5 tabs, each
tab driven by a list view. Replaces the stock Awards, Celebrations, and
Tenured Champions list web parts on the Our Partners home page; designed
to be reusable on future pages.

## Property pane

- **Section title** (text, optional) — header rendered above the tab strip
  ("PHILLIPS AWARDS", "CELEBRATIONS", "TENURED CHAMPIONS")
- **List** (dropdown, required) — lists on the current site
- **Layout** (dropdown, required) — `Gallery cards` | `Table rows`
- **Number of tabs** (dropdown, required) — `2` | `3` | `4` | `5`
- **Tabs** (reorderable list, N entries matching tab count) — each entry:
  - **Label** (text, required)
  - **View** (dropdown, required) — views on the selected list
- **See All link** (URL, optional) — when set, renders "SEE ALL" in
  top-right of the web part
- **Show card overlay** (checkbox, only enabled when Layout = Gallery
  cards) — when checked:
  - **Source field** (dropdown, required) — fields on the selected list
  - **Label template** (text, required) — supports `{value}` token
  - **Position** (dropdown, required) — `Top-left` | `Top-right` |
    `Bottom-left` | `Bottom-right`

## Property pane validation

- List required before tabs can be configured
- Every visible tab requires label and view
- If overlay enabled, source field and label template required
- Tab count clamping: reducing the count hides extra entries but
  retains them in state so they reappear if count rises

## Runtime behavior

- Horizontal tab strip above content area. Active tab: Phillips red
  underline (`#c8102e`). Inactive: muted gray text. Click switches the
  active view.
- First tab in configured order is active on page load. No per-user
  persistence.
- Items fetched via REST using the selected view's query (fields,
  filter, sort, row limit). View formatting JSON on the list is
  ignored — we render with our own components.
- **Gallery cards**: image on top, title below, optional secondary line.
  16px rounded corners. Overlay badge rendered per config when source
  field is non-empty for that item.
- **Table rows**: clean rows with column headers, no SharePoint chrome.
- **Empty state**: "No items to display" centered in content area; tab
  strip remains visible.
- **Container**: 16px rounded corners on outer frame.

## Overlay rendering

- Renders only when Layout = Gallery cards AND "Show card overlay" is
  checked AND the item's source field is non-empty
- Background: Phillips red (`#c8102e`), white text, small rounded pill
- `{value}` in the label template is substituted with the source field's
  value (e.g., source field returns `5` → label `{value} YEARS` renders
  as "5 YEARS")
- Position: absolute, anchored to the configured corner of the card
  image

## Non-features (explicit)

- No threshold logic, computation modes, or `{threshold}` token —
  calculated columns on the list handle eligibility and tier values
- No multiple overlays per card
- No persisting last-selected tab per user
- No auto-detect layout from view
- No pagination beyond the view's row limit
- No overlays on table rows
- No cross-site list references
- No native view formatting JSON honored
- No color picker — Phillips red is hardcoded

## Three configured placements on Our Partners

| Placement | Section title | List | Layout | Tabs | Overlay |
| --- | --- | --- | --- | --- | --- |
| Awards | PHILLIPS AWARDS | Awards | Gallery cards | Awards Brand Ambassador (+ others to be decided by editor) | Off |
| Celebrations | CELEBRATIONS | Celebrations | Table rows | Today's Birthdays, This Month's Anniversaries (+ others TBD) | N/A |
| Tenured Champions | TENURED CHAMPIONS | Partner Profiles | Gallery cards | Tenured Champions (+ others TBD) | Source: `Tenured Champion Milestone`, Template: `{value} YEARS`, Position: Bottom-left |

The tab configurations beyond the first tab per placement are left to the
editor to set up post-deployment; the spec covers the web part, not the
specific tab content choices.

## Solution version

Bump from 1.0.0.26 to 1.0.1.0 (new feature warrants minor version bump,
not just patch).

## Definition of done

- [ ] Web part builds without errors
- [ ] Web part deployed to tenant App Catalog
- [ ] Property pane renders correctly with all fields
- [ ] Tab switching works
- [ ] Both layouts render correctly
- [ ] Overlay renders on Tenured Champions (Lauren = "5 YEARS",
      Shelly = "10 YEARS")
- [ ] Empty state renders correctly when a view returns no items
- [ ] No console errors on any of the three placements
- [ ] Human visual verification on Our Partners home page
