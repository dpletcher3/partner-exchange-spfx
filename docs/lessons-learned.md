# Lessons learned

Hard-won insights from building Partner Exchange SPFx. Add entries here when something took real time to figure out and shouldn't have to be rediscovered. Newer entries at the top; date each section so the timeline is obvious. Keep entries narrow — one root cause per heading, not catch-all summaries.

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
