# Phillips News web part — design spec

Design specification for the `PhillipsNews` SPFx web part. This document is the contract between the design decisions captured in chat and the scaffold turn that follows. Update this file when the design changes; do not let the code and the doc drift.

## Purpose

Render branded news cards on the hub home page, sourced from the News Repository list. Replaces the stock SharePoint News rollup for the Phillips Loop and Phillips in the News surfaces. One web part, two instances on the hub home, each configured for its own surface — same code, same rendering, different property values.

Rationale: stock News rollups don't give us the categorical and branded control needed across two distinct news surfaces with consistent styling. Per [decisions.md D027](../../partner-exchange-provisioning/docs/decisions.md) (provisioning repo), the news surface is a custom SPFx web part rather than stock parts.

## Identity

- **Name (internal):** `PhillipsNews`
- **Display name:** Phillips News
- **Description:** Renders branded news cards from the News Repository list, filterable by category and item type.
- **Group in web part picker:** Phillips Custom Components
- **Manifest version:** 1.0
- **Bundled in:** the existing partner-exchange-spfx solution (no new sppkg)
- **Solution version bump:** next after 1.0.0.17 → 1.0.0.18

## Property pane

Single page, two groups.

### Group 1 — Content

| Property | Type | Default | Notes |
|---|---|---|---|
| `sectionTitle` | text | empty | Renders as `<h2>` above the grid. Empty = no header, no underline. Editor sets per instance ("The Phillips Loop", "Phillips in the News"). |
| `categoryFilter` | multi-select | empty | Choices loaded dynamically from the Category column at property-pane open. Empty selection = no category filter. |
| `itemTypeFilter` | dropdown | `(any)` | Choices loaded dynamically from the ItemType column. `(any)` = no item-type filter. |
| `maxItems` | number | 6 | Min 1, max 24. |
| `showViewAllLink` | toggle | on | When on, renders "View all →" to the right of the section title, pointing at the News Repository list view (`{sourceSiteUrl}/Lists/News Repository/AllItems.aspx`). |

### Group 2 — Advanced (collapsed by default)

| Property | Type | Default | Notes |
|---|---|---|---|
| `sourceSiteUrl` | text | hub URL constant from `config/constants.ts` | Absolute site URL where the list lives. Lets the web part be repointed without code changes. |
| `listTitle` | text | `News Repository` | List title at the source site. |

Both fields appear in the Advanced group so editors managing the 90% case (hub-sourced news) don't see configuration they don't need.

## Data source

- `INewsRepositoryService` defines the contract: `getCategories(siteUrl, listTitle)`, `getItemTypes(siteUrl, listTitle)`, `getNewsItems(siteUrl, listTitle, filters, maxItems)`.
- `NewsRepositoryService` implements against SharePoint REST (`SPHttpClient`).
- `MockNewsRepositoryService` returns deterministic fixture data for local dev and future tests.
- Web part component depends on the interface, not the implementation. Service is injected via the web part's `onInit`.

### Query shape

`getNewsItems` builds a REST URL of the form:

```
{sourceSiteUrl}/_api/web/lists/getbytitle('{listTitle}')/items
  ?$select=Id,Title,Category,ItemType,LinkUrl,ThumbnailImage,ShortDescription,PublishedDate
  &$filter=...        (built from categoryFilter and itemTypeFilter)
  &$orderby=PublishedDate desc
  &$top={maxItems}
```

`ThumbnailImage` is a SharePoint Image column, which returns serialized JSON. Service deserializes to `{ serverRelativeUrl, alt }`.

`LinkUrl` is a URL column, returning `{ Url, Description }`. Service normalizes to a plain string URL.

## Rendering

### Layout

- Section sits on a subtle gray surface (`var(--color-background-secondary)` or equivalent), padded `2rem 1.75rem`, with `var(--border-radius-lg)` corners. Cards float on this surface; the surface separates the news section from the surrounding page chrome.
- Section header (`<h2>` when `sectionTitle` is set; nothing when blank). Phillips red 2px underline below the header. "View all →" link aligned right, muted, when `showViewAllLink` is on.
- Grid of cards: `repeat(auto-fit, minmax(190px, 1fr))`, `gap: 1rem`. Resolves to 3 columns on desktop, 2 on tablet, 1 on mobile at typical SharePoint canvas widths.

### Card

- White background, `0.5px` border (default `var(--color-border-tertiary)`), `16px` border radius (`var(--border-radius-lg)` if it's 16px in this codebase; otherwise a literal `16px`).
- 4:3 thumbnail at the top, `overflow: hidden` so the image inherits the rounded top corners.
- If `ThumbnailImage` is empty, render a colored fallback block keyed off the category. (Specific fallback treatment deferred to the polish turn; scaffold can use a solid Phillips-red fallback as placeholder.)
- Content area `padding: 14px 16px 16px`:
  - **Category label** — 11px, weight 500, uppercase, letter-spacing 0.08em, muted text color, 6px margin-bottom.
  - **Title** — 15px, weight 500, line-height 1.35, 8px margin-bottom.
  - **Date** — 12px, muted (`var(--color-text-tertiary)`), formatted as "May 23, 2026", 8px margin-bottom.
  - **Description** — 13px, line-height 1.55, secondary text color, truncated with ellipsis at ~120 chars.
- Entire card is a single `<a href="{LinkUrl}">` — whole card is the click target.

### Hover state

- Border darkens from `var(--color-border-tertiary)` to `var(--color-border-secondary)`.
- Card background shifts subtly (very light tint or no change — pick during polish; scaffold can ship without hover styling).
- No scale. No drop shadow.

### States

The component renders one of four states based on the service result:

1. **Loading** — 6 skeleton cards matching the populated layout. Used during the initial fetch and on property changes.
2. **Populated** — the grid, with one card per item returned.
3. **Empty** — short message "No news to show right now." centered in the section. Used when the query returns zero items (typically because filters exclude everything).
4. **Error** — inline error message with a retry button. Used on REST failure.

All four states must be visible during scaffold dev. After the build, loading and error are harder to hit naturally — that's fine.

## Accessibility

- Section is wrapped in a `<section aria-labelledby="...">` referencing the `<h2>` when present.
- Cards are `<a>` elements with descriptive accessible names (the card's title text is sufficient as the link text via `aria-label` on the anchor).
- Thumbnail images use the `alt` text from `ThumbnailImage`. If alt is empty, the image is treated as decorative (`alt=""`) and the card title carries the meaning.
- "View all →" link has explicit text content, not just an icon.
- Loading skeleton uses `aria-busy="true"` on the grid container.

## What's deferred

- **+Add news item button** — its own turn (second-to-last in the build sequence: `scaffold → list service → cards → polish → +Add → deploy`).
- **Pagination or "load more"** — not in v1. Editor sets `maxItems`.
- **Audience targeting per item** — backlog, not v1.
- **Layout variants** — confirmed not needed for v1; same card design for both surfaces.
- **Visual polish** — separate turn. Scaffold ships functional but not yet brand-polished.
- **Unit tests on the service** — decide during the list-service turn.

## Files (scaffold turn deliverable)

```
src/webparts/phillipsNews/
  PhillipsNewsWebPart.ts
  PhillipsNewsWebPart.manifest.json
  loc/
    en-us.js
    mystrings.d.ts
  components/
    PhillipsNews.tsx
    PhillipsNews.module.scss
    NewsCard.tsx
    NewsCard.module.scss
    NewsGrid.tsx
    EmptyState.tsx
    ErrorState.tsx
    LoadingState.tsx
  services/
    INewsRepositoryService.ts
    NewsRepositoryService.ts
    MockNewsRepositoryService.ts
    models.ts
  config/
    constants.ts
```

Plus `config/config.json` and `config/package-solution.json` updates to register the new web part in the bundle.

## Definition of done — scaffold turn

- Web part appears in the page picker under `Phillips Custom Components` group
- Drops onto the hub home page without error
- Renders loading state, then populated grid of cards from the seeded News Repository items
- Property pane opens; all properties listed above are present; category and item-type choices populate from the list dynamically
- Changing properties triggers re-query and re-render
- Empty state visible when filter set excludes everything
- Error state visible when forced (temporarily point `listTitle` at a non-existent list)
- Card clicks open the `LinkUrl`
- Solution version bumped to 1.0.0.18
- This doc is committed in the SPFx repo and any deviations from it during the scaffold are reflected back here before merge

## Scaffold-turn deviations (2026-05-26)

Discovered against the real seeded `News Repository` list during the scaffold turn. Reflected here per the "do not let the code and the doc drift" rule above.

- **`Category` is a MultiChoice column, not single-value.** The list returns `Category` as an array (e.g. `["Phillips In The News"]`). Two consequences:
  - The `INewsItem` model field is `categories: string[]` (not `category: string`). The card renders the first category as its label; multi-category display is a polish-turn concern.
  - **Category cannot be used in an OData `$filter`** (SharePoint rejects `eq` on multi-value fields). So the query shape above is amended: only `ItemType` (a single-value Choice) is filtered server-side. Category filtering is done **client-side** after the fetch — when a category filter is active the service over-fetches (ceiling 200) ordered by `PublishedDate desc`, filters by category membership, then slices to `maxItems`. With no category filter, `$top` stays exact at `maxItems`.
- **`ThumbnailImage` real-URL resolution deferred to the polish turn.** The seeded items' Image-column JSON is the `Reserved_ImageAttachment` shape (`{ fileName, originalImageName }`) with no `serverRelativeUrl`, so the scaffold renders the spec-sanctioned solid Phillips-red fallback for every card. Resolving the attachment URL (or re-uploading thumbnails so the column serializes a `serverRelativeUrl`) is polish-turn work.
- **`LinkUrl` resolved empty at runtime despite a valid CLI shape.** CLI inspection (`odata.metadata=minimal`) shows `LinkUrl` as a structured `{ Url, Description }` object with `Url` populated, and `SPHttpClient` negotiates the same metadata level — yet the rendered card anchor had an empty `href`, meaning the runtime value reached the mapper in a different shape. The service now extracts the URL **defensively** via `extractUrl()`, handling both a `{ Url }` object and a bare string, and emitting a `console.warn('[PhillipsNews] Unexpected LinkUrl shape', …)` only when it receives an object with no `Url` (silent on the happy path, so the console stays clean). Confirming the precise runtime shape and removing the defensive fallback is polish-turn cleanup.
- **`ShortDescription` is empty on the seeded items** (`null` in REST). Not a bug — a content gap. Cards correctly render no description until the column is populated; truncation is at ~120 chars.
- **List has a leftover default `Choice 3`** in both the Category and ItemType columns (list-creation artifact). Cosmetic; the web part renders whatever choices the columns expose. Worth cleaning up in the list config, out of scope for this turn.

## Decisions captured

| Decision | Choice | Date |
|---|---|---|
| One web part with two instances vs. two specialized web parts | One web part, configurable | (this conversation) |
| Same layout for Phillips Loop and Phillips in the News in v1 | Yes | (this conversation) |
| Card aspect ratio | 4:3 (balanced) | (this conversation) |
| Category badge style | Subtle uppercase text label, no background | (this conversation) |
| Card corner radius | 16px (modern) | (this conversation) |
| Hover treatment | Border darken + light bg shift, no scale, no shadow | (this conversation) |
| Section background | Subtle gray surface (not white-on-white) | (this conversation) |
| Section header underline | Phillips red 2px | (this conversation) |
| Source-site handling | Editor-configurable property with hub-URL default | (this conversation) |
| Web part group in picker | Phillips Custom Components | (this conversation) |
| `+ Add` button | Deferred to its own turn | (this conversation) |
| "View all →" link in v1 | Keep, points at the list view | (this conversation) |
