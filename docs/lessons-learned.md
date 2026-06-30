# Lessons learned

Hard-won insights from building Partner Exchange SPFx. Add entries here when something took real time to figure out and shouldn't have to be rediscovered. Newer entries at the top; date each section so the timeline is obvious. Keep entries narrow — one root cause per heading, not catch-all summaries.

---

## SPFx property-pane data loaders must self-heal from getPropertyPaneConfiguration

*Discovered 2026-06-30 (PhillipsHighlightVideo property pane). Fixed in solution 1.0.17.6, PhillipsHighlightVideoWebPart.ts.*

**Symptom:** HighlightVideo's property-pane dropdowns — the Featured-item picker and the Title/Video/Info field-mapping dropdowns — intermittently rendered empty and disabled, so the author couldn't select a featured item. Intermittent, not every time.

**Root cause:** the pane loaders (`_loadColumnsForCurrentList` / `_loadItemsForCurrentList`) were triggered ONLY by `onPropertyPaneConfigurationStart` (plus a `listId` change in `onPropertyPaneFieldChanged`). `onPropertyPaneConfigurationStart` races property hydration: if `this.properties.listId` was still falsy when it fired, each loader hit its silent `!listId` early-return and **nothing re-triggered the load** — `getPropertyPaneConfiguration` only *read* `_columnsLoadedFor` / `_itemsLoadedFor`, it never *called* the loaders. The pane then stuck permanently in its 'loading' branch (empty + disabled dropdowns). Intermittent because it depended entirely on whether property hydration beat `onPropertyPaneConfigurationStart` on that render.

**Fix:** self-heal from `getPropertyPaneConfiguration` — the one pane method guaranteed to run on every pane render. It now kicks each loader when the source id is set but `_…LoadedFor !== id` and that loader isn't already in flight. The loaders' own `(_loadedFor === id || _loading)` guard plus their completion `propertyPane.refresh()` keep this to exactly one fetch per id and re-render the pane once data lands. `onPropertyPaneConfigurationStart` stays as a best-effort first attempt; the self-heal is the safety net.

**Cross-web-part caveat — a `getPropertyPaneConfiguration` self-heal is ONLY safe when the loader is idempotent per-source.** It must have BOTH (a) a `_loadedFor === source` "already loaded" guard AND (b) an in-flight guard. Without the "already loaded" guard it re-fetches on every pane render; and because most loaders call `propertyPane.refresh()` on completion, that becomes a refresh → getPropertyPaneConfiguration → load → refresh **loop**.

PhillipsAudienceHero was deliberately **NOT** given this fix:

- **It doesn't have the bug.** Its resolver inputs are default-backed (`_ppSiteUrl` → `DEFAULT_PP_SITE`, `_ppListTitle` → `'Partner Profiles'`), so there is no falsy-source early-return that can race hydration — `onPropertyPaneConfigurationStart`'s resolve always proceeds, and was confirmed firing live (`path=live`).
- **Applying it would be harmful.** `_resolvePartnerProfiles` has only an in-flight `_resolving` guard (no `_resolvedFor === source`) and itself calls `propertyPane.refresh()`, so kicking it from `getPropertyPaneConfiguration` would loop.

**General principle:** before adding a `getPropertyPaneConfiguration` self-heal, confirm the loader is idempotent per-source (has a loaded-for guard). If it isn't, add that guard first — or don't self-heal from there.

---

## SPFx file/image property pane controls — use PropertyFieldFilePicker

*Discovered 2026-05-27 (Phillips Personalized Hero, I08).*

For any web part property that holds a file or image URL, default to
PropertyFieldFilePicker from @pnp/spfx-property-controls, not
PropertyPaneTextField. The picker gives editors three tabs (Browse,
Upload, From a link) and matches the experience editors get from
SharePoint's stock web parts.

A text field for a URL is technically functional but immediately reads
as wrong to editors who expect a Browse button. Discovered during the
Phillips Personalized Hero (I08) build — initial scaffold used a text
field and was visibly off-pattern; the file picker was added as a same-
turn refinement.

Default to PropertyFieldFilePicker for any future file-valued property.

---

## SharePoint REST: CLI output is not the runtime shape

*Discovered 2026-05-26 (PhillipsNews scaffold).*

SP REST responses for typed column fields (URL, Image, multi-choice, lookup)
can have different shapes between m365 CLI inspection and what arrives at
the web part at runtime. The CLI normalizes responses; the SPFx
`SPHttpClient` (with `odata.metadata=minimal`) does not in the same way.

Hit twice during PhillipsNews scaffold:

- `LinkUrl` (URL column): CLI shows `{ Url, Description }` object;
  runtime returned an empty string under `minimal` metadata.
- `ThumbnailImage` (Image column): CLI shows usable structure;
  runtime returns the `Reserved_ImageAttachment` shape without a
  resolvable `serverRelativeUrl`.

Practice: write defensive shape-handling helpers (`extractUrl`,
`extractChoices`, etc.) that accept both shapes and log a warning
with the unexpected shape on the failure path. The warning becomes
the actual diagnostic when the assumption is wrong, instead of
silently returning empty values.

Where you can: hit the live REST endpoint from the browser DevTools
console of an actual SharePoint page, not via the CLI, when shaping
the service mapping for a typed field.

---

## SPFx + Fluent UI v9 in SharePoint surfaces

*Discovered 2026-05-25 (Iter 2c.3, commit b33b858).*

- **Do not wrap SPFx components in a FluentProvider unless you are
  actually consuming v9 theme tokens.** SharePoint's own modern
  surfaces (property panes, new-item panels, callouts) are themselves
  built on Fluent v9 and render into portaled DOM. Their styles depend
  on the class-numbering sequence Griffel produces from the first
  FluentProvider it sees. Inserting an unused provider higher in the
  tree shifts that numbering and silently breaks SharePoint's own
  panels (most visible symptom: transparent backgrounds).

- Diagnosis pattern: if SharePoint's own panels/callouts render
  transparent or unstyled after an SPFx deploy, suspect an added
  provider, theme wrapper, or style injection before suspecting CSS.

- Defensive selectors should be MORE specific than the primary, not
  broader. `[data-sp-feature-tag="Site header host"]` looked like a
  good fallback for hiding the SP site header — it isn't, because the
  same attribute is reused on panel and callout hosts.
