# Phillips News — polish turn spec

Visual refinement pass on the Phillips News web part. This is the polish turn in the build sequence (`scaffold → list service → cards → polish → +Add → deploy`). List-service turn closed in commit `1869575`; polish was deferred per D028 while I08 jumped ahead.

This turn changes visual state only. No new behavior, no new data flow, no property pane additions. The card design that's live on the hub today should look noticeably better after this turn, but should not behave differently.

## Scope — explicit

What this turn delivers:

1. Grid minmax floor bumped from 190px to 240px so the grid resolves to 3 columns at full canvas width instead of 5
2. Hover treatment refined — border darkens on hover, no background shift, no scale, no shadow
3. Card-height alignment in rows so cards with different description lengths still align cleanly across a row
4. Content guidelines doc created in a new location, separate from the web part spec

What this turn explicitly does NOT deliver:

- No-thumbnail fallback softening (deferred — wait until the grid bump is live and we can evaluate the red treatment against the new visual density)
- +Add button (its own next turn)
- Font selection or new property pane controls
- Color theming or category-keyed treatments
- Any change to the data layer

## Scope: grid minmax bump

Current SCSS:

```scss
.grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(190px, 1fr));
  gap: 1rem;
}
```

After:

```scss
.grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
  gap: 1rem;
}
```

One value changes. The grid now resolves to 3 columns at typical SharePoint canvas widths (≈800–1100px) and 4 columns on very wide canvases. The visual density goes from cramped to comfortable.

No hard column cap — `auto-fit` keeps the grid responsive. If a future use case wants a 3-column hard max, that's a separate change.

## Scope: hover treatment

Current behavior: cards have a static border, no hover affordance.

After: on `:hover`, the border color transitions from `var(--color-border-tertiary)` to `var(--color-border-secondary)`. That's the only change.

SCSS:

```scss
.card {
  border: 0.5px solid var(--color-border-tertiary);
  transition: border-color 0.15s ease-out;
}

.card:hover {
  border-color: var(--color-border-secondary);
}
```

No background shift, no scale, no shadow, no cursor change (the cursor is already a pointer via the parent `<a>`).

Rationale: cards already have a thumbnail and clear content; the border darken alone is enough hover affordance. Background shifts and shadows are useful on text-only cards or list rows; they're visual noise on image-led cards like these. Restrained is right.

Touch-device handling: hover styles don't fire on touch by default (good), no need for media query gating.

## Scope: card-height alignment

Current behavior: cards in a row have natural heights — a card with no ShortDescription is shorter than a card with a two-line description. Across a row of three, the heights are uneven.

After: all cards in a row stretch to match the tallest card in that row.

Two implementation options, pick whichever is cleaner in the existing SCSS:

**(a) Grid alignment**

```scss
.grid {
  align-items: stretch;
}

.card {
  display: flex;
  flex-direction: column;
  height: 100%;
}

.cardContent {
  flex: 1;
}
```

CSS Grid stretches cards to fill their row height; the card uses flex internally to push content properly inside the stretched height. This is the standard "equal-height grid items" pattern.

**(b) Minimum description height**

```scss
.description {
  min-height: 3.6em; /* ~2 lines at line-height 1.55 */
}
```

Reserves space for description even when empty. Simpler but the card stays at a fixed height regardless of how much description content actually exists; longer descriptions overflow or get truncated rather than expanding the card.

Pick (a). It's the correct architectural pattern (cards adapt to content; grid aligns them), and it handles a row of cards with mixed content lengths cleanly. (b) is a hack that papers over the symptom.

## Scope: content guidelines doc

New file: `docs/seed-data/news-repository-guidelines.md`.

This doc lives outside the web part spec because content guidelines outlive specific web parts. If other lists or web parts adopt similar conventions later, the guidelines should be a content-side artifact, not buried in a single component's spec.

Content to include:

```markdown
# News Repository — content guidelines

How to author items in the News Repository list so they render cleanly in the Phillips News web part on the hub home page and any other surface that pulls from this list.

## Title

- Keep it under 60 characters. Longer titles wrap to a second line in the card and start crowding the description.
- Avoid all-caps unless the source uses it (e.g., acronyms). The card's typography handles emphasis; ALL CAPS reads as shouting.
- Title sentence case is recommended over title case for readability ("How the additive team cut prototype cycles" not "How The Additive Team Cut Prototype Cycles").

## ShortDescription

- Keep it under 120 characters. The card truncates with an ellipsis at roughly that point.
- One sentence. If you need two, the article is the place for the second one; the card description is a teaser.
- Lead with the most newsworthy fact, not the setup ("Phillips named to Defense News Top 100 for the seventh year" not "Earlier this month, an industry publication announced...").
- Don't repeat the title. The card shows both; the description should add information.
- Plain text only. No formatting, links, or HTML.
- Optional — items without a description still render, with a slightly shorter card. Prefer including one for visual consistency in a row.

## ThumbnailImage

- Recommended size: at least 800×600 (landscape orientation). Smaller images get upscaled and lose sharpness.
- Aspect ratio: 4:3 is ideal. Other ratios get center-cropped to fit the 4:3 card frame.
- Composition: keep important detail (text, faces, logos) within the **center 80% of the image**. Edges may be cropped to fit a 4:3 frame at various card widths.
- Avoid heavy text overlays on the image; the card has its own text below the thumbnail. Branded headers or watermarks at the image edges are likely to be cropped — re-export with the brand in the center, or rely on the card title.
- Optional — items without a thumbnail render with a solid Phillips-red placeholder block. Prefer including a thumbnail; the placeholder dominates the visual grid when several items lack images.

## Category

Multi-select. Pick one or more from the configured choices:
- Phillips Loop — internal culture, partnership, milestones, behind-the-scenes
- Phillips In The News — external coverage of Phillips by third-party publications
- (other categories as added)

An item can belong to multiple categories. It will surface in any web part instance filtered to a category it belongs to (e.g., the Markforged partnership announcement tagged both "Phillips Loop" and "Phillips In The News" appears in both hub-home instances).

## ItemType

Single value. Picks the kind of item:
- Internal Story — points at a SharePoint or Phillips internal page
- External Link — points at a phillipscorp.com or other external site
- (other types as added)

This drives where clicking a card takes the user. The LinkUrl field controls the actual destination; ItemType is a categorical tag for filtering and (in future) different rendering treatments.

## LinkUrl

- Always populate this. Cards are clickable; without a URL the card has nothing to navigate to.
- Use the canonical URL from phillipscorp.com or the intranet, not a redirected or tracking link.
- The card opens this URL in a new tab. Same-tab navigation isn't currently configurable.

## PublishedDate

- Use the date the news was originally published, not the date you added it to the list.
- The card displays this date prominently. Backdating to "make news look fresh" is a poor practice — surface real news, or remove items that are no longer timely.
- Items are ordered by PublishedDate descending in the web part; the most recent news appears first.

## Editorial cadence

- The web part shows up to N items per instance (default 6, configurable in the property pane).
- Aim for at least 6 active items in the list at any time so the grid is always full.
- Archive or remove items more than ~6 months old unless they are still being referenced. Old items pushed off the visible grid are still in the list and searchable; they just don't appear in the default rendering.
```

That's a complete starter draft. Trim or expand based on what feels right after a few weeks of authoring.

## Files touched

```
src/webparts/phillipsNews/components/
  NewsGrid.module.scss            (minmax bump)
  NewsCard.module.scss            (hover + flex layout for height alignment)
  PhillipsNews.module.scss        (potentially align-items: stretch — depends on which file owns the grid)

docs/seed-data/
  news-repository-guidelines.md   (new file)
```

No changes to TypeScript, no manifest changes, no property pane additions. Solution version bumps 1.0.0.24 → 1.0.0.25.

## Definition of done

- Grid resolves to 3 columns at typical SharePoint canvas widths instead of 5
- Hover on a card transitions the border color smoothly (no flash, no jank); no other hover side effects
- Cards in a row align to equal height even when one card has a long description and others have none
- Content guidelines doc exists at `docs/seed-data/news-repository-guidelines.md`
- Console clean
- Deployed and visually verified

## Verification (browser walkthrough)

After deployment:

1. Hard-refresh the hub home page
2. Confirm the news grid now shows 3 columns across (was 5)
3. Hover over a card — border darkens smoothly, no background change, no card lift, no shadow
4. Hover over another card — same effect; original card's border returns to default
5. Locate a row that has cards with mixed description lengths (one short or empty, one with a full description). Confirm all cards in that row are the same height; the shorter card stretches to match
6. Confirm the section padding and gutters look balanced — the bump to 240px shouldn't break the section layout
7. Resize the browser to narrow widths (~600px, ~400px) and confirm responsive behavior — should drop to 2 columns and then 1
8. Console clean

## Polish-turn handoff

After this turn closes, the +Add button turn picks up. The polish-turn backlog gets updated in `phillips-news.md`:

- Items completed: minmax bump, hover, card-height alignment, content guidelines
- Items deferred: no-thumbnail fallback softening (evaluate after this turn; may no longer be needed)

The +Add turn is then the last turn before "deploy" closes out I06 entirely.

## Decisions captured

| Decision | Choice | Date |
|---|---|---|
| Polish before +Add | Yes — visual locks first, workflow second | 2026-05-27 |
| Grid minmax floor | 240px (was 190px) | 2026-05-27 |
| Hover treatment | Border darken only — no bg shift, no scale, no shadow | 2026-05-27 |
| Card-height alignment | Yes — `align-items: stretch` + flex card | 2026-05-27 |
| No-thumbnail fallback softening | Defer — re-evaluate after grid bump lands | 2026-05-27 |
| Content guidelines home | Separate doc at `docs/seed-data/news-repository-guidelines.md` | 2026-05-27 |
