# CLAUDE.md

## What this project is

Partner Exchange SPFx is the custom SharePoint Framework solution that gives the Phillips Corporation intranet its visual identity and adds the one piece of functionality built-in SharePoint web parts cannot reproduce. It contains exactly two SPFx components: an Application Customizer that injects brand CSS site-wide (so all OOTB Hero, News, Quick Links, Highlighted Content, and Button web parts pick up Phillips styling automatically), and a Personalized Hero web part (greeting that varies by time of day, current user's display name, live clock, time-zone selector). Everything else — page composition, news authoring, quick link management, social posts, search — uses standard SharePoint web parts that content authors already know how to use.

## Locked architectural decisions

- **SPFx runtime:** 1.23.x (pinned in `package.json` via `@microsoft/sp-*` deps). *Bumped from 1.20.x on 2026-05-20 because Microsoft moved off Node 18 starting with SPFx 1.21; we'd rather migrate now with zero code than later under deadline.*
- **SPFx generator:** 1.23.x (Yeoman generator package. Verify with `npm view @microsoft/generator-sharepoint version`)
- **React:** 17 (SPFx 1.23.x still pins React 17 — do not upgrade to 18 without confirming SPFx-side support)
- **TypeScript:** 5.x (introduced by SPFx 1.21+, present here via SPFx 1.23.x's bundled rush-stack-compiler-5.x). *Bumped from 4.7.x alongside the runtime bump on 2026-05-20 — single coordinated migration.*
- **UI library:** Fluent UI v9 (`@fluentui/react-components`)
- **Node:** 22 LTS, pinned via `.nvmrc`. *Bumped from 18 on 2026-05-20 — SPFx 1.21+ requires Node 22 LTS; staying on 18 would have stranded us on the unsupported SPFx 1.20.x line.*
- **Styling:** SCSS modules per component, with CSS custom properties for brand tokens defined once in `src/styles/_tokens.scss` and imported by both the customizer and the web part
- **Solution packaging:** one `.sppkg` containing both the Application Customizer and the Personalized Hero web part
- **CI:** GitHub Actions, `.github/workflows/build.yml`, produces `.sppkg` as artifact
- **Deployment:** PnP CLI for Microsoft 365 (`m365` command), not PnP PowerShell — works cross-platform and pairs better with Claude Code on macOS
- **Repo host:** GitHub, `dpletcher3/partner-exchange-spfx`
- **Branch protection:** `main` requires PR + passing CI; no direct pushes

## Hard constraints

These must always be true. If a prompt or change would violate them, stop and raise the conflict instead of working around it.

- **No new SPFx web part is built if the same brand effect can be achieved via CSS in the customizer or via list view formatters.** This is the most important rule. SPFx web parts are a tax on content authors; we only add them when no alternative exists.
- **The Application Customizer's CSS injection is the only global side effect.** No other scripts run on every page. Any new global behavior requires a new design conversation, not a quiet addition.
- **All colors are CSS custom properties.** No hex codes in component code or SCSS files other than `_tokens.scss`. If a new color is needed, add it to tokens first.
- **All fonts come from brand tokens.** Never inline `font-family` declarations in components.
- **Accessibility is non-negotiable.** Focus rings, keyboard navigation, ARIA roles, and color contrast must be preserved or improved by every change. If a brand override removes a focus ring, the change is wrong.
- **Mobile responsiveness must not regress.** SharePoint pages render on phones; the customizer cannot break that.
- **CSS specificity stays as low as possible.** Single-class selectors when feasible. Don't fight SharePoint's own styles with `!important` unless there's no other option; comment why if used.
- **No external network calls from the Application Customizer.** No CDN-loaded fonts at runtime (fonts are deployed via SharePoint Brand Center or bundled), no analytics beacons, no remote config. Keeps Content Security Policy simple and avoids dependency on services.
- **The Personalized Hero web part uses `this.context.pageContext.user.displayName` only.** No Microsoft Graph API calls. Keeps it fast, avoids permission grants, eliminates auth complexity.

## Brand tokens (single source of truth)

Defined in `src/styles/_tokens.scss`. Every component imports from here.

### Colors

| Token | Value | Use |
|---|---|---|
| `--phil-red` | `#F9423A` | Primary brand color, CTAs, accents |
| `--phil-red-tint` | `#fef0ee` | Icon backgrounds, hover surfaces |
| `--phil-gold` | `#F68B33` | Secondary accent, Federal/Defense hub variant |
| `--phil-blue` | `#00AEEF` | Electric blue, hero backgrounds |
| `--phil-black` | `#1a1a1a` | Headings, primary text |
| `--phil-gray-900` | `#323130` | Body text |
| `--phil-gray-600` | `#605e5c` | Secondary text |
| `--phil-gray-400` | `#8a8886` | Tertiary text, captions |
| `--phil-gray-200` | `#e1dfdd` | Borders |
| `--phil-gray-100` | `#edebe9` | Dividers |
| `--phil-gray-50` | `#faf9f8` | Page background |
| `--phil-white` | `#ffffff` | Surface background |

### Typography

| Token | Value | Use |
|---|---|---|
| `--phil-font-display` | `'Barlow Condensed', 'Arial Narrow', sans-serif` | Headings, hero text |
| `--phil-font-body` | `'Montserrat', 'Segoe UI', system-ui, sans-serif` | Body, UI |
| `--phil-eyebrow-letter-spacing` | `2px` | Small caps eyebrow labels |

Barlow Condensed and Montserrat are loaded via SharePoint Brand Center (admin configured separately, not by SPFx).

### Spacing scale

| Token | Value |
|---|---|
| `--phil-space-1` | `4px` |
| `--phil-space-2` | `8px` |
| `--phil-space-3` | `12px` |
| `--phil-space-4` | `16px` |
| `--phil-space-5` | `24px` |
| `--phil-space-6` | `32px` |
| `--phil-space-7` | `48px` |

### Border radius scale

| Token | Value | Use |
|---|---|---|
| `--phil-radius-sm` | `4px` | Inline elements, small badges |
| `--phil-radius-md` | `8px` | Icon backgrounds, small cards |
| `--phil-radius-lg` | `12px` | Cards, tiles, news cards |
| `--phil-radius-xl` | `16px` | Hero, featured content, CTA sections |
| `--phil-radius-pill` | `999px` | Pill buttons |

## Domain rules

- **Eyebrow pattern:** A red 16px-wide 1.5px-tall horizontal bar followed by uppercase 9–10px text in `--phil-red` with 2px letter-spacing. Used above every section heading. The customizer exposes a utility class `.phil-eyebrow` for this; components that need to render an eyebrow programmatically use the same class.
- **Greeting by time of day in the Personalized Hero:**
  - 5:00–11:59 → "Good morning"
  - 12:00–16:59 → "Good afternoon"
  - 17:00–21:59 → "Good evening"
  - 22:00–4:59 → "Working late" (a friendly nod, not an interrogation)
- **Pill button:** background `--phil-red`, color white, padding 7px 14px, border-radius `--phil-radius-pill`, includes a right-arrow icon (`→` or Fluent's `ChevronRight`).
- **Icon-in-square pattern (Quick Links):** 32px square with `--phil-red-tint` background, `--phil-radius-md` corners, brand-red icon centered.
- **Card pattern:** white background, 0.5px `--phil-gray-200` border, `--phil-radius-lg` corners, no shadow.
- **Hero pattern:** `--phil-radius-xl` corners, overlay color `rgba(0, 50, 80, 0.35)` for darkening photo backgrounds.

## Known gotchas

- **SPFx ApplicationCustomizers do NOT run on classic site collections** even when individual pages render with modern-looking chrome. The site collection itself must be group-connected (`GroupId != 00000000-0000-0000-0000-000000000000`) for the SPFx extension pipeline to engage. On a classic site, `_spPageContextInfo` is `undefined` on the page, the customaction is silently delivered in the page payload but the component manifest is never resolved, and no amount of correct registration/packaging will make the customizer load.
  - **`PartnerExchange-DanSandbox` is a classic site collection and is permanently unusable for SPFx extension dev/debug.** Confirmed via `m365 spo site get --url https://phillipscorp.sharepoint.com/sites/PartnerExchange-DanSandbox` showing `GroupId: 00000000-0000-0000-0000-000000000000`. Do not try to use it as a target site again.
  - **Current dev site:** `https://phillipscorp.sharepoint.com/sites/spfx-extension-test` (group-connected team site, SPFx-eligible).
  - **Always validate the target site type with `m365 spo site get --url <siteUrl>` before deploying** — check `GroupId`. The all-zero GUID is the smoking gun.
- **The diagnostic banner from Prompt 4 is still in `onInit()`.** [src/extensions/phillipsBrand/PhillipsBrandApplicationCustomizer.ts](src/extensions/phillipsBrand/PhillipsBrandApplicationCustomizer.ts) currently appends a red `<div id="phil-test-banner">🟥 PARTNER EXCHANGE CUSTOMIZER LOADED 🟥</div>` to `<body>` and logs `[PhilCustomizer] Banner injected, onInit complete`. This was added to prove end-to-end bundle loading on the new dev site and **must be removed before Prompt 5 starts** — the banner is not a product feature, only the brand CSS injection is. Search for `phil-test-banner` and `[PhilCustomizer]` and delete the block (the `TEMPORARY DIAGNOSTIC` comment in the source marks the boundary).

## How to work in this codebase

- **Every prompt completion in this project ends with a Manual Test section — numbered steps the developer runs to verify the work, written so they could be followed by someone who didn't see the code change. No prompt is considered complete without it.**
- **Clarity over cleverness in CSS.** Verbose selectors that future Dan can read beat tight ones future Dan can't.
- **Comments in SCSS explain *why* a rule exists,** especially overrides of SharePoint defaults. Future maintainers won't know that `.ms-CommandBar` overrides exist to remove the dated drop shadow unless we tell them.
- **Component files stay under 200 lines.** If a hero component grows past that, split it into subcomponents.
- **Git workflow.** Dan has delegated git ops to Claude. At the end of every prompt's Manual Test, *after Dan confirms it passed*, commit and push without asking. The discipline:
  - Run `git status` first and summarize what changed in the response — short, no walls of text.
  - Group changes into logical commits, not one giant commit per prompt. Typical pattern: one commit for the substantive code change (`feat`/`fix`/`refactor`), one for any docs updates (`docs`), one for any chore work (`chore`). If the prompt only produced one logical change, one commit is fine.
  - Use conventional commit prefixes: `feat`, `fix`, `docs`, `chore`, `refactor`, `test`, `style`. Scope to the area changed when useful: `feat(customizer):`, `docs(claude):`.
  - Commit message body references the prompt number completed: `Completes Prompt N.` on its own line.
  - Push after committing — don't leave commits sitting locally.
  - Never `git add` `node_modules/`, `lib/`, `lib-commonjs/`, `.claude/`, build artifacts, or anything in `.gitignore`. If something unexpected shows up in `git status`, stop and surface it.
  - Never `git push --force`. If a push is rejected (remote ahead, branch protection, anything), stop and ask — don't try to resolve it solo.
  - After push, report: number of commits, short commit hashes, and confirmation the push succeeded.
  - If you need a decision about how to split commits or what message to use, ask. If you hit a merge conflict, stop and surface it.
- **PRs reference the prompt number** they completed: "Closes Prompt 4 (CSS injection)." (Commit bodies say "Completes Prompt N"; PR titles/descriptions say "Closes Prompt N" — both can be true for the PR that lands a prompt's work.)
- **Don't bypass type errors.** No `@ts-ignore`, no `any`. If TypeScript complains, fix the underlying problem.
- **Test in the SharePoint Workbench first**, then on a real test site. Workbench catches structural issues; only real sites catch CSP, theme, and OOTB-web-part interaction issues.
- **Site type validation comes FIRST when SPFx extensions don't load.** Before opening DevTools, inspecting bundles, or auditing manifests for any "extension isn't running" or "bundle isn't loading" symptom, run these three console checks on the affected page:
  ```js
  window._spPageContextInfo?.isSPO        // must be true
  window._spPageContextInfo?.pageItemId   // must be a GUID
  typeof window.SPClientPlugin            // must be 'undefined' (modern)
  ```
  If `isSPO` is `undefined` or `false`, the page is not SPFx-eligible regardless of how correct the package, manifest, or customaction registration is — SharePoint will silently include the customaction in the page payload but never resolve the component manifest because SPFx itself never initialized. This is the **first** thing to check, not the last. Also confirm the underlying site collection with `m365 spo site get --url <siteUrl>` and verify `GroupId` is not `00000000-0000-0000-0000-000000000000` (the all-zero GUID indicates a classic, non-group-connected site, which doesn't run SPFx extensions even when pages look modern).

## Graduation readiness

What would have to change for this solution to support a wider audience than just Phillips:

- All brand tokens would need to be externalized to a configuration file (JSON or environment-based) so the solution can be re-themed per tenant
- The customizer's English-only UI strings would need internationalization (resource files per locale, SPFx supports this natively)
- The Personalized Hero's time-zone default (currently Eastern) would need to come from user profile or site config
- The hardcoded "Phillips" references in greetings and copy would need to be tokenized
- The provisioning JSON files would need a tenant-substitution step in CI

These are not in v1 scope. Document them here so future-us doesn't have to rediscover them.

## Verified behavior

To be filled in during Phase 3 as prompts complete and behavior is confirmed. Format: short bullet under a sub-heading per prompt.

### Prompt 3 — Application Customizer scaffold

- **clientSideComponentId**: `e37132c1-6b4b-4c0c-9d59-2e35c666cb8f`
  - Source: [src/extensions/phillipsBrand/PhillipsBrandApplicationCustomizer.manifest.json](src/extensions/phillipsBrand/PhillipsBrandApplicationCustomizer.manifest.json) line 4
  - Needed in Prompt 4 for `m365 spo customaction add --clientSideComponentId ...` and in CI / deployment scripts
- Entry point class: `PhillipsBrandApplicationCustomizer` (default export, extends `BaseApplicationCustomizer<Record<string, never>>`)
- Customizer takes no configurable properties — `ClientSideComponentProperties` is `{}` in both `sharepoint/assets/elements.xml` and `sharepoint/assets/ClientSideInstance.xml`
- Debug URL: `https://phillipscorp.sharepoint.com/sites/spfx-extension-test/SitePages/Home.aspx` (migrated from `PartnerExchange-DanSandbox` on 2026-05-20 after that site was confirmed unusable for SPFx — see "Known gotchas"). Note: [config/serve.json](config/serve.json) and several `PROMPTS.md` references still hard-code the old URL and need updating in a separate non-docs commit.

### Prompt 4 — Brand CSS injection (working, banner still in place)

- **Status:** Working end-to-end on `https://phillipscorp.sharepoint.com/sites/spfx-extension-test`. Confirmed via the diagnostic banner on `SitePages/Home.aspx` and `SitePages/testpage.aspx`. `getComputedStyle(document.documentElement).getPropertyValue('--phil-red')` returns `" #F9423A"`; `<style id="phil-brand">` is present in `<head>`.
- **Solution version deployed:** `1.0.0.1` (the bump from `1.0.0.0` was required so SharePoint would recognize the new bundle instead of serving cached metadata for the old hash)
- **Bundle hash:** `4fd914910a7d83b565a9` (file in `ClientSideAssets/` is `phillips-brand-application-customizer_4fd914910a7d83b565a9.js`)
- **App ID:** `7b2e6ef9-8db7-41cd-9660-3aee7feb8f63` (matches `ProductID` in `AppManifest.xml`)
- **Installed at:** `https://phillipscorp.sharepoint.com/sites/spfx-extension-test`
- **CSS approach:** SCSS compiled at build time into a string export (`src/extensions/phillipsBrand/generated/phillipsBrandCss.ts`, gitignored), injected manually via `document.createElement('style')` with `id="phil-brand"`. The sp-css-loader auto-injection path was abandoned because it routes CSS through `window.__themeState__.loadStyles`, which silently drops non-themable `:root` declarations. Full diagnostic in commit `cfdce55`.
- **Diagnostic banner still in `onInit()`** — see "Known gotchas" for the removal checklist before Prompt 5.

## Lessons learned

Add insights here as they emerge from the build process — things that took time to figure out and shouldn't have to be rediscovered.

### Validate the target site type *before* deep-diving the code

The Prompt 4 CSS-injection failure consumed a long debugging session: I audited the build pipeline, manifest schema, package layout, feature XML, ClientSideAssets folder, sp-css-loader output, and runtime CSS injection — all of which were correct. The actual cause was that the original target site (`PartnerExchange-DanSandbox`) is a classic site collection, and SPFx extensions simply don't run on classic sites. The customaction got registered cleanly, SharePoint dutifully shipped it in the page payload, and then the SPFx runtime — which never initialized on that page — never resolved the component manifest. Symptom: zero network requests for the bundle, zero console output, registration looks perfect.

**Rule of thumb for next time:** when SPFx behavior is inexplicable, spin up a fresh modern team site (`m365 spo site add --type TeamSite --url https://<tenant>.sharepoint.com/sites/<name> --title <name> --alias <name>`), deploy the app there, and see if the symptom reproduces. This takes ~2 minutes and definitively answers "is the code the problem or is the target site the problem." Make it one of the first diagnostic moves, not the last. The "Site type validation" bullet in "How to work in this codebase" is the inlined version of this principle.
