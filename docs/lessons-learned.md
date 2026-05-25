# Lessons learned

Hard-won insights from building Partner Exchange SPFx. Add entries here when something took real time to figure out and shouldn't have to be rediscovered. Newer entries at the top; date each section so the timeline is obvious. Keep entries narrow — one root cause per heading, not catch-all summaries.

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
