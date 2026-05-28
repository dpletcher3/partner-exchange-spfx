# Phillips List View — Web Part Spec

## Purpose

Generic SPFx web part that renders a SharePoint list with Phillips brand
styling. Two layout choices (Gallery cards or Table rows), an **optional**
tab strip driven by list views (2–5 tabs), and an optional configurable
overlay badge on cards. Image columns and Person fields both resolve to
thumbnails so a single web part covers both photo-bearing lists (Awards
with an Image column) and people lists (Partner Profiles where the photo
comes from each row's LinkedUser).

Replaces the four stock list-rendering web parts on the Our Partners home
page — Meet Your New Partners, Phillips Awards, Celebrations, Tenured
Champions — with four configured instances of this one web part.

Originally shipped as "Tabbed List Views" (I11 1.0.1.x). Renamed and
expanded to "Phillips List View" in I11 1.0.2.0 when the tab strip became
optional and Person-field photo support was added. Manifest GUID is
preserved across the rename (`4a8c3b9e-1d6f-4e8a-9b2c-5a7d8e3f1b2c`), so
1.0.1.x instances continue to render.

## Property pane

- **Section title** (text, optional) — header rendered above the tab strip
- **List** (dropdown, required) — lists on the current site
- **Layout** (dropdown, required) — `Gallery cards` | `Table rows`
- **Show tabs** (toggle, default ON) — when ON, the 2–5 tab strip drives
  the renderer; when OFF, a single configured view drives the renderer
- **Number of tabs** (dropdown, only when Show tabs is ON) — `2` | `3` |
  `4` | `5`
- **Tabs** (reorderable list, only when Show tabs is ON) — N entries
  matching tab count; each entry:
  - **Label** (text, required)
  - **View** (dropdown, required) — views on the selected list
- **View** (dropdown, only when Show tabs is OFF, required) — the single
  view that drives the renderer
- **See All link** (URL, optional) — when set, renders "SEE ALL" in
  top-right of the web part
- **Show card overlay** (toggle, only enabled when Layout = Gallery cards)
  — when checked:
  - **Source field** (dropdown, required) — fields on the selected list
  - **Label template** (text, required) — supports `{value}` token
  - **Position** (dropdown, required) — `Top-left` | `Top-right` |
    `Bottom-left` | `Bottom-right`

### Property pane validation

- List required before any view / tab configuration can begin
- When Show tabs is ON: every visible tab requires label and view
- When Show tabs is OFF: the single View dropdown is required
- If overlay enabled, source field and label template required
- Tab count clamping: reducing the count hides extra entries but
  retains them in state so they reappear if count rises
- Changing the **List** clears tabs, viewId, and overlaySourceField so
  stale references to the previous list don't survive the swap

## Runtime behavior

### Tabs ON

Horizontal tab strip above the content area. Active tab: Phillips red
underline. Inactive: muted gray text. Click switches the active view.
First tab in configured order is active on page load. No per-user
persistence of the last selected tab.

### Tabs OFF

No tab strip rendered. Section title still renders above the content
area if set. The renderer is fed the single configured viewId directly.

### Renderer (shared by both modes)

- Items fetched via REST using the selected view's CAML (fields,
  filter, sort, row limit) through `RenderListDataAsStream`. View
  formatting JSON on the list is ignored — we render with our own
  components.
- **Gallery cards**: image on top, title below, optional secondary
  line. 16px rounded corners. Overlay badge rendered per config when
  source field is non-empty for that item.
- **Table rows**: clean rows with column headers, no SharePoint chrome.
- **Empty state**: "No items to display" centered in content area; tab
  strip (when shown) remains visible.
- **Container**: 16px rounded corners on outer frame.

### Component separation

The renderer (`ListViewContent`) is a self-contained component that
takes a single `viewId` plus list / layout / overlay props and handles
data loading, loading/empty/error states, and Gallery/Table dispatch.
The orchestrator (`PhillipsListView`) composes the header, the
**optional** tab strip, and the renderer. This separation is why "Show
tabs" works cleanly: the tab strip is a sibling of the renderer rather
than wrapped around it, and the renderer doesn't know or care whether
its `viewId` came from a tab selection or the property pane.

## Image-field resolution

The Gallery layout auto-detects which field on a row carries the
thumbnail. Detection order:

1. **Image / Thumbnail columns** (field `TypeAsString` ∈ {`Thumbnail`,
   `Image`}). If present in the view, the first one wins. Renders by
   pulling the matching attachment URL from a conditional
   `AttachmentFiles` sidecar fetch (the same path that powers
   PhillipsNews). Lists without an image column skip the sidecar
   entirely.
2. **Person / User columns** (field `TypeAsString` ∈ {`User`,
   `UserMulti`}), priority-ordered by internal-name substring. The
   first priority substring that matches any of the view's Person
   fields wins:
   - `LinkedUser`
   - `User`
   - `Person`
   - `Profile`
3. **Other Person columns** not matching the priority list and not on
   the deprioritized list (`Author`, `Editor`, `CreatedBy`,
   `ModifiedBy` — the admin Person fields present on every SP list).
4. **Row-value sniffing** — last-resort fallback for Hyperlink columns
   serving as ad-hoc image fields. Reads the first row's value for
   each field and matches anything that looks like a serialized image
   payload.

### Person-field rendering

When a Person field is the resolved image source:

- Read the `email` (or `Email` / `EMail`) property from the field's
  value entry.
- Construct `/_layouts/15/userphoto.aspx?accountname={encoded-email}&size=L`
  and render that as the `<img>` src.
- If the field is empty or the entry has no email, fall back to the
  gray placeholder (the same fallback used when an Image column row
  has no attachment).
- `size=L` (~96px) is the right footprint for our 280px card thumb at
  4:3 aspect — smaller sizes pixelate noticeably.

### Why auto-detect (and not a "Photo source" property)

The four current placements span three list shapes (Image column,
LinkedUser Person column, table-only Celebrations with no thumbnail at
all). The priority order encodes the pattern of "an explicit image
column wins; otherwise the row's headline person is the photo;
otherwise no photo". Adding a property would push that pattern onto
the editor for every instance, even though the right answer is the
same one across all lists we render.

## Overlay rendering

- Renders only when Layout = Gallery cards AND "Show card overlay" is
  checked AND the item's source field is non-empty
- Background: Phillips red (`--phil-red`), white text, small rounded pill
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
- No "Photo source" property — image-field auto-detect covers every
  list shape we render
- No Microsoft Graph dependency for Person photos — userphoto.aspx
  resolves photos without any additional permission grant

## Four configured placements on Our Partners

| Placement | Section title | List | Layout | Show tabs | Tabs / View | Overlay |
| --- | --- | --- | --- | --- | --- | --- |
| Meet Your New Partners | MEET YOUR NEW PARTNERS | Partner Profiles | Gallery cards | OFF | View: New Partners (editor configures) | Off |
| Phillips Awards | PHILLIPS AWARDS | Awards | Gallery cards | ON | Brand Ambassador, Virtuoso, Remote Rockstar (editor configures) | Off |
| Celebrations | CELEBRATIONS | Celebrations | Table rows | ON | Today's Birthdays, This Month's Anniversaries (editor configures) | N/A |
| Tenured Champions | TENURED CHAMPIONS | Partner Profiles | Gallery cards | ON | Tenured Champions (+ others TBD by editor) | Source: `TenuredChampionMilestone`, Template: `{value} YEARS`, Position: Bottom-left |

The tab configurations beyond the first tab per placement are left to the
editor to set up post-deployment; the spec covers the web part, not the
specific tab content choices.

Meet Your New Partners and Tenured Champions both point at Partner
Profiles. They differ on the view (New Partners vs Tenured Champions),
on the tab strip (off vs on), and on the overlay (off vs the years-pill).
The image-field auto-detect resolves to `LinkedUser` for both, so face
photos render in both placements without per-instance configuration.

## Solution version

I11 1.0.2.0 (feature expansion: rename + Show tabs toggle + Person-field
photo support). Bumped from 1.0.1.5.

## Definition of done

- [ ] Web part builds without errors
- [ ] Web part deployed to tenant App Catalog at 1.0.2.0
- [ ] Page toolbox shows it as "Phillips List View"
- [ ] Property pane renders correctly with Show tabs toggle
- [ ] Toggle OFF: tab strip hidden, single View dropdown visible
- [ ] Toggle ON: tab strip + tab collection editor visible (existing behavior)
- [ ] Tab switching still works when tabs are on
- [ ] Both layouts render correctly
- [ ] Image-column photos render on Awards (Test Award)
- [ ] Person-column photos render on Partner Profiles (Lauren, Shelly)
- [ ] Overlay renders on Tenured Champions (Lauren = "5 YEARS",
      Shelly = "10 YEARS")
- [ ] Empty state renders correctly when a view returns no items
- [ ] No console errors on any of the four placements
- [ ] Human visual verification on Our Partners home page
