# Phillips News — +Add button turn spec

Last turn in the build sequence (`scaffold → list service → cards → polish → +Add → deploy`). Adds an editor-only inline affordance that opens the SharePoint new-item form for the News Repository list.

## Scope — explicit

What this turn delivers:

- A `+ Add news item` button that appears inline in the rendered web part
- Visible **only when the page is in edit mode** (`displayMode === DisplayMode.Edit`); invisible in read mode
- Clicking opens the SharePoint new-item form for the configured News Repository list in a new tab
- Distinct from the existing "View all →" link — different action, different placement, different audience
- Visible on both web part instances on the hub home (Phillips Loop and Phillips In The News), regardless of which filter is active

What this turn explicitly does NOT deliver:

- Inline new-item form embedded in the page (the +Add button just opens SharePoint's standard form in a new tab — same as clicking + New on the list itself)
- Automatic web part refresh after a new item is added (editor refreshes the page themselves to see new content)
- Pre-populating the new-item form with the active filter's category or item-type values
- Permission checking (relies on SharePoint to gate the new-item form itself; if the editor doesn't have write access to the list, SharePoint shows the access-denied page in the new tab)
- Editor-configurable button label or position
- A confirmation toast or feedback after the editor returns to the hub home page

## Affordance design

### Placement

The button sits in the section header, on the right side, next to the existing "View all →" link.

Visual layout of the section header when this turn lands:

```
THE PHILLIPS LOOP                        + Add news item    View all →
─────────────                                                          (red 2px underline below)
```

The two right-aligned elements stack horizontally with a small gap (`gap: 1rem` or similar). `+ Add news item` sits to the left of `View all →` — closer to the editor's mental flow of "add things, then see all things."

In read mode, the section header looks the same as today:

```
THE PHILLIPS LOOP                                          View all →
─────────────
```

### Visibility logic

```typescript
const isEditMode = this.context.displayMode === DisplayMode.Edit;
// Pass isEditMode as a prop to the component; render the button conditionally
{isEditMode && <AddNewsItemButton ... />}
```

The button is conditionally rendered, not just visually hidden. In read mode, the DOM element does not exist — no orphan classes, no accessibility-tree clutter.

### Styling

- Text: `+ Add news item` (the + is a literal character, not an icon)
- Font: matches the existing "View all →" link (same family, weight, size)
- Color: same muted gray as "View all →" by default; on hover, transitions to Phillips red to signal the call-to-action
- Cursor: pointer
- No background, no border — same visual weight as the View all link. Both are quiet links, not buttons.

If you want it more prominent (e.g., a pill button with a red background), say so before implementation — it's a small SCSS change but worth being deliberate.

### Click behavior

Opens the SharePoint new-item form for the News Repository list in a new tab:

```
{sourceSiteUrl}/Lists/{listTitle}/NewForm.aspx
```

For the current configuration that resolves to:

```
https://phillipscorp.sharepoint.com/sites/PartnerExchange-Sandbox/Lists/News%20Repository/NewForm.aspx
```

Implementation:

```jsx
<a 
  href={newItemUrl}
  target="_blank"
  rel="noopener noreferrer"
  className={styles.addButton}
>
  + Add news item
</a>
```

Same `target="_blank" rel="noopener noreferrer"` pattern as the news cards — opens in a new tab, doesn't leak the opener.

URL construction is a derived value from the existing `sourceSiteUrl` and `listTitle` property pane values. No new property pane fields. URL-encode the list title to handle spaces (`News Repository` → `News%20Repository`).

## Files touched

```
src/webparts/phillipsNews/
  components/
    PhillipsNews.tsx                  (pass isEditMode prop)
    NewsGrid.tsx (or wherever the section header lives)
                                       (add the AddNewsItemButton inline)
    AddNewsItemButton.tsx             (new — the small link component)
    AddNewsItemButton.module.scss     (new — styling)
    PhillipsNews.module.scss          (small layout adjustment for the header right side if needed)
  PhillipsNewsWebPart.ts              (pass displayMode/isEditMode through to the component)
```

No services changed, no manifest changes, no property pane changes. Solution version bumps 1.0.0.25 → 1.0.0.26.

## Definition of done

- The `+ Add news item` button appears in the section header of each Phillips News web part instance **only in page edit mode**
- The button is positioned to the left of "View all →" in the right side of the section header
- The button is styled as a quiet text link, similar weight to "View all →", with a hover color shift to Phillips red
- Clicking opens the SharePoint new-item form for the configured News Repository list in a new tab
- In read mode (non-editors, or editors viewing without the page in edit), the button is absent from the DOM
- The button respects the `sourceSiteUrl` and `listTitle` property pane values — pointing it at a different list (via the Advanced property pane group) would change where the button takes the editor
- Solution version bumped to 1.0.0.26
- Console clean

## Verification (browser walkthrough)

After deployment:

1. Hard-refresh the hub home page in **read mode** (just viewing, not editing)
2. Confirm the section headers show `View all →` only — no Add button
3. Click `Edit` to enter edit mode on the page
4. Confirm both Phillips News instances now show `+ Add news item` to the left of `View all →`
5. Hover the button — text color shifts to Phillips red, cursor is pointer
6. Click the button — a new tab opens at the News Repository new-item form
7. Fill out and save a test item in the new tab (optional, to confirm full flow)
8. Return to the hub home page tab, refresh — the new item appears in the grid if it matches the instance's filters
9. Exit edit mode on the page (via Discard or Save+Close)
10. Confirm the Add button is gone from the section headers
11. Open DevTools, inspect the section header in edit mode — confirm the link has `target="_blank"` and `rel="noopener noreferrer"`
12. Console clean throughout

## Edge cases worth knowing

- **Multiple instances on the same page**: each instance shows its own Add button when in edit mode, both pointing at the same News Repository list URL. Editor can click either; same destination.
- **Pointing the web part at a different list via Advanced**: if `listTitle` is set to something other than "News Repository", the +Add URL automatically points at that list's NewForm. Same applies to `sourceSiteUrl`.
- **List doesn't exist**: if the editor has misconfigured `listTitle`, the +Add button opens a new tab to a 404. The cards in the web part would also be in their error state, so this isn't a new failure mode.
- **Editor lacks write permission to the list**: SharePoint's NewForm.aspx handles this — the user sees an access-denied page in the new tab. Not the web part's concern.

## Decisions captured

| Decision | Choice | Date |
|---|---|---|
| Where +Add affordance appears | Inline in section header, edit mode only | 2026-05-27 |
| What clicking does | Opens SharePoint NewForm.aspx in a new tab | 2026-05-27 |
| Distinct from View all | Yes — separate affordance, different placement, edit-mode-only | 2026-05-27 |
| Editor visibility | Visible only in `DisplayMode.Edit`; conditionally rendered (not just hidden) | 2026-05-27 |
| Button styling | Quiet text link matching "View all →" weight; hover transitions to Phillips red | 2026-05-27 |
| URL source | Derived from existing `sourceSiteUrl` and `listTitle` property pane values | 2026-05-27 |
| Pre-populate filter values into new-item form | No | 2026-05-27 |
| Inline form vs. SharePoint NewForm | NewForm in new tab — defer inline form indefinitely | 2026-05-27 |

## What closes after this turn

This turn completes the original Phillips News build sequence (`scaffold → list service → cards → polish → +Add`). The remaining sequence item — "deploy" — is the final verified-on-production milestone, which is bundled with I10 (sandbox to production promotion) when that comes around.

For the active I06 increment:

- Scaffold ✓ (commits 2deb7a1, 739ee57, 206adc1, cdc2d8b)
- List-service hardening ✓ (commit 1869575)
- Polish ✓ (this push)
- +Add ✓ (this turn, once deployed and verified)

After +Add lands, I06 is functionally complete on sandbox. Phillips News web part is feature-complete for v1. Anything further (audience targeting, pagination, different layout modes, etc.) is post-v1 work.

`current-increment.md` should be updated post-deploy to reflect I06's full completion.
