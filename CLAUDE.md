# CLAUDE.md

## What this project is

Partner Exchange SPFx is the custom SharePoint Framework solution that gives the Phillips Corporation intranet its visual identity and adds the one piece of functionality built-in SharePoint web parts cannot reproduce. It contains exactly two SPFx components: an Application Customizer that injects brand CSS site-wide (so all OOTB Hero, News, Quick Links, Highlighted Content, and Button web parts pick up Phillips styling automatically), and a Personalized Hero web part (greeting that varies by time of day, current user's display name, live clock, time-zone selector). Everything else — page composition, news authoring, quick link management, social posts, search — uses standard SharePoint web parts that content authors already know how to use.

## Locked architectural decisions

- **SPFx runtime:** 1.23.x (pinned in `package.json` via `@microsoft/sp-*` deps). *Bumped from 1.20.x on 2026-05-20 because Microsoft moved off Node 18 starting with SPFx 1.21; we'd rather migrate now with zero code than later under deadline.*
- **SPFx generator:** 1.23.x (Yeoman generator package. Verify with `npm view @microsoft/generator-sharepoint version`)
- **React:** 17.0.1, pinned in `package.json` `dependencies` (`react`, `react-dom`). SPFx 1.23.x still requires React 17 — do not upgrade to 18 (Microsoft's compatibility table is explicit, and SPFx's webpack pipeline will silently fail at runtime if React 18 is installed). Added 2026-05-24 in Iteration 2b when the customizer began rendering React into the Top placeholder; see "Iteration 2b" under "Verified behavior" and D026 in the `partner-exchange-provisioning` repo for the design rationale.
- **TypeScript:** 5.x (introduced by SPFx 1.21+, present here via SPFx 1.23.x's bundled rush-stack-compiler-5.x). *Bumped from 4.7.x alongside the runtime bump on 2026-05-20 — single coordinated migration.*
- **UI library:** None at runtime. The customizer's `BrandedHeader` renders plain HTML elements (`<div>`, `<img>`, `<nav>`, `<a>`) styled by CSS modules. `@fluentui/react-components` (`^9.46.0`) is still in `dependencies` but is **not imported anywhere in `src/` after Iter 2c.3** — the Iter 2b `<FluentProvider>` wrapper was removed on 2026-05-25 because it triggered a `fui-FluentProvider#` class-numbering collision that left SharePoint's portaled v9 panels (New Item form, web part property pane, Site Information, Highlighted Content config) with transparent surfaces (see `microsoft/fluentui#23821`, `SharePoint/sp-dev-docs#9847`). The package is kept as a dependency only so a future component that genuinely needs v9 doesn't require reinstalling the lockfile; do not mount another `<FluentProvider>` on a page where SharePoint also renders v9 panels (i.e. anywhere in modern SPO) without first verifying it doesn't reproduce the collision. Fluent UI v8 (`@fluentui/react`) is still in `devDependencies` for compatibility with SPFx-shipped types but new components should not import it either.
- **Externals:** `config/config.json` `externals: {}` is intentionally empty. SPFx 1.22+ moved external-library configuration out of `config.json` and into the Heft toolchain's webpack plugins, which handle React externalization automatically. Do not add a `react` / `react-dom` block to `config.json` — it is silently ignored by the Heft rig and creates confusion about who owns externalization. See Microsoft Learn "SharePoint Framework Toolchain: Heft & Webpack" for the official statement.
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
- **The Application Customizer has exactly TWO global side effects: CSS injection into `<head>` (the `<style id="phil-brand">` tag), and a React tree rendered into the SharePoint `PlaceholderName.Top` slot (the `BrandedHeader` component, rendered as a plain `<div>` tree styled by CSS modules — NO `FluentProvider` wrapper).** Both were authorized by an explicit design conversation — the CSS side effect by the original brand work, and the React render by D026 in the `partner-exchange-provisioning` repo (Iteration 2b, 2026-05-24). The Iter 2b/2c FluentProvider wrapper was removed in Iter 2c.3 (2026-05-25) because it broke SharePoint's portaled v9 panels via a `fui-FluentProvider#` class-numbering collision; the BrandedHeader uses no Fluent v9 components at runtime, so the wrapper was load-bearing for nothing. Any further global behavior on top of these two requires another design conversation, not a quiet addition.
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
- **The diagnostic banner from Prompt 4 has been removed.** Historical note: during Prompt 4 the customizer briefly appended a red `<div id="phil-test-banner">🟥 PARTNER EXCHANGE CUSTOMIZER LOADED 🟥</div>` to `<body>` (plus a `[PhilCustomizer] Banner injected, onInit complete` console log) to prove end-to-end bundle loading on the new dev site. That block has since been stripped — [src/extensions/phillipsBrand/PhillipsBrandApplicationCustomizer.ts](src/extensions/phillipsBrand/PhillipsBrandApplicationCustomizer.ts) now only injects the brand `<style id="phil-brand">` tag. If a future change reintroduces a similar banner for diagnostics, the same removal discipline applies: it is not a product feature and must be excised before the change ships.

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

### Prompt 4 — Brand CSS injection (working, banner removed)

- **Status:** Working end-to-end on `https://phillipscorp.sharepoint.com/sites/spfx-extension-test`. Originally confirmed via a temporary diagnostic banner on `SitePages/Home.aspx` and `SitePages/testpage.aspx` (since removed); the banner-free runtime check is `getComputedStyle(document.documentElement).getPropertyValue('--phil-red')` returning `" #F9423A"` and `<style id="phil-brand">` being present in `<head>`.
- **Solution version:** Now `1.0.0.12` in [config/package-solution.json](config/package-solution.json) (bumped during Iteration 2b; previously `1.0.0.9` at the time of Prompt 4's verified state, then `1.0.0.11` after intervening deploys). The initial `1.0.0.0` → `1.0.0.1` bump during Prompt 4 was required so SharePoint would recognize the new bundle instead of serving cached metadata for the old hash; every subsequent .sppkg upload to the App Catalog continues to require a version bump for the same cache-busting reason.
- **App ID:** `7b2e6ef9-8db7-41cd-9660-3aee7feb8f63` (matches `ProductID` in `AppManifest.xml` and the `solution.id` in [config/package-solution.json](config/package-solution.json)).
- **Installed at:** `https://phillipscorp.sharepoint.com/sites/spfx-extension-test` plus the five I01 sandbox sites (`PartnerExchange-Sandbox`, `-OurCulture`, `-OurPartners`, `-Dashboard`, `-TheHub`) — installed and customaction-registered via `scripts/i01-foundation/install-customizer-app.ps1` and `scripts/i01-foundation/register-customizer.ps1` in the `partner-exchange-provisioning` repo.
- **CSS approach:** SCSS compiled at build time into a string export (`src/extensions/phillipsBrand/generated/phillipsBrandCss.ts`, gitignored), injected manually via `document.createElement('style')` with `id="phil-brand"`. The sp-css-loader auto-injection path was abandoned because it routes CSS through `window.__themeState__.loadStyles`, which silently drops non-themable `:root` declarations. Full diagnostic in commit `cfdce55`.
- **Diagnostic banner:** Removed — `onInit()` invoked only `injectBrandStyles()` at the end of Prompt 4. (As of Iteration 2b, `onInit()` also wires up the Top-placeholder render; as of Iteration 2c that placeholder renders the real `BrandedHeader` — wordmark, "PARTNER EXCHANGE" subtitle, hub nav strip with active-state underline — not the red placeholder bar.) The original red banner remains historical; see "Known gotchas" for the full record.

### Iteration 2b — Branded header placeholder render (CSS + React, working)

- **Status:** Working end-to-end. The customizer now performs **two** global side effects on every modern page where it's registered: (1) injects `<style id="phil-brand">` into `<head>` (unchanged from Prompt 4), and (2) renders a `BrandedHeader` React component — wrapped in `FluentProvider` with `webLightTheme` — into `PlaceholderName.Top`. Iteration 2b's BrandedHeader is a styled red 80px placeholder reading "BRANDED HEADER PLACEHOLDER"; the real header design lands in Iteration 2c.
- **Solution version at deploy:** `1.0.0.12`. Bundle hash: `phillips-brand-application-customizer_9157aef177473f466b7e.js`.
- **Design authorization:** D026 in the `partner-exchange-provisioning` repo. This is the design conversation that authorized expanding the customizer from one global side effect (CSS) to two (CSS + Top-placeholder React render). Future expansions need their own decision record.
- **Stack added in this iteration:** `react@17.0.1`, `react-dom@17.0.1`, `@fluentui/react-components@^9.46.0` as production dependencies; `@types/react@17.0.45`, `@types/react-dom@17.0.17` as devDependencies. `npm install` raised peer-dep override warnings because some transitive packages want `react@17.0.2`; the override is benign because SPFx 1.22+ requires `react@17.0.1` exactly (Microsoft's compatibility table is the authority).
- **Externals:** Not changed. `config/config.json` `externals: {}` is intentionally empty — see "Locked architectural decisions" above for why SPFx 1.23's Heft toolchain handles React externalization itself.
- **Placeholder pattern:** `placeholderProvider.tryCreateContent(PlaceholderName.Top, { onDispose })` plus `placeholderProvider.changedEvent.add(this, ...)` — placeholders aren't guaranteed to be available at `onInit` time, so subscribing to `changedEvent` is the canonical way to handle late availability. `_onDispose` calls `ReactDOM.unmountComponentAtNode` to clean up the React tree if the placeholder is removed. The customizer's `_injectBrandStyles` remains idempotent.
- **Runtime check (banner-free):** `document.getElementById('phil-brand')` is non-null (CSS), and the `PlaceholderName.Top` slot contains a `<div>` with a red 80px child reading "BRANDED HEADER PLACEHOLDER" (React render).

### Iteration 2c — Real BrandedHeader (wordmark + subtitle + hub nav + hide SP header)

- **Status:** Deployed; awaiting browser verification on the five sandbox sites. The Iteration 2b red placeholder bar is replaced by the production-shape `BrandedHeader`: Phillips wordmark on the left (PNG inlined as a base64 data URI from `src/assets/Logo.png`), "PARTNER EXCHANGE" subtitle separated by a 1px gray rule, and a right-aligned uppercase nav strip whose active item gets a 2px red underline.
- **Solution version at deploy:** `1.0.0.13`. Bundle hash: `phillips-brand-application-customizer_5d40422996b126d2fc83.js`.
- **Design authorization:** Still D026 (Iteration 2b's decision record). Iteration 2c is the visual realization of the same architectural change, no new global side effect.
- **Hub nav source:** `spHttpClient.get('<hub-root>/_api/web/Navigation/TopNavigationBar', SPHttpClient.configurations.v1)`. The hub root is hardcoded to `https://phillipscorp.sharepoint.com/sites/PartnerExchange-Sandbox` per D023 (sandbox specifics in code; production deploy updates the constant). Result is cached in `sessionStorage` under key `phil-hub-nav-cache` for 5 min. On fetch failure (network, permissions, hub re-org), a 5-item `FALLBACK_NAV` constant takes over and a `console.warn` is logged.
- **Active-state rule:** Strip protocol+host from both the link's `SimpleUrl` and `window.location.href`, then mark active iff `currentPath === linkPath` OR `currentPath.startsWith(linkPath + '/')`. The trailing-slash guard is the load-bearing bit — it prevents `/sites/PartnerExchange-Sandbox` from being treated as a prefix of `/sites/PartnerExchange-Sandbox-OurCulture` (the bug-shape that motivated this exact comparison).
- **Loading state:** Five skeleton placeholders (`80px × 12px`, `--phil-gray-100` background, 1.5s opacity-pulse animation) render until either the cache hit or the fetch resolves. Cache hit is synchronous from React's perspective, so users with a warm session see no skeletons.
- **Hide SP site header:** New rule at the bottom of [src/styles/_pass3-overlay.scss](src/styles/_pass3-overlay.scss): `[data-automation-id="SiteHeader"] { display: none !important; }`. Hides the redundant SharePoint site-header row (site logo + site title + Home/Documents/Pages/Site contents/Edit nav) now that our BrandedHeader owns the top of the page. `!important` is acceptable here because the rule lives in pass3-overlay, which is the documented place for that pattern.
- **Logo inline encoding:** `src/assets/Logo.png` (22,520 bytes) base64-encodes to 30,028 chars, embedded directly into [BrandedHeader.tsx](src/extensions/phillipsBrand/components/BrandedHeader.tsx) as a `LOGO_DATA_URI` constant. Inlining avoids an extra network round-trip and keeps CSP clean (no `img-src` allowlist for a CDN). Regenerate via `base64 -w0 src/assets/Logo.png` if the PNG changes.
- **2c.1 selector fix:** `data-automation-id` was the wrong attribute spelling; SharePoint uses `data-automationid` (no hyphen). The Iter 2c selector `[data-automation-id="SiteHeader"]` matched nothing and the SP site-header chrome stayed visible. Replaced with `#spSiteHeader, [data-sp-feature-tag="Site header host"]` — the wrapper id catches both the hub-nav row and the site-title/admin row in a single rule, and the feature-tag attribute is a documented fallback if Microsoft renames the id. Solution version `1.0.0.14`. Bundle hash: `phillips-brand-application-customizer_4996d6031892637007a3.js`.
- **2c.2 selector fix:** Removed `[data-sp-feature-tag="Site header host"]` from the hide-SP-header rule in [src/styles/_pass3-overlay.scss](src/styles/_pass3-overlay.scss). The attribute is generic — SharePoint applies `data-sp-feature-tag` to many feature hosts beyond the site header, including panel, callout, and dialog hosts. The Iter 2c.1 "defensive fallback" was matching those unrelated hosts and applying `display: none`, which the diagnostic at the time **mistakenly identified as the cause of the transparent panel bug**. Solution version `1.0.0.16`. Bundle hash: `phillips-brand-application-customizer_4574a0a8dfbbd0771197.js`. **Lesson learned:** defensive fallback selectors should be MORE specific than the primary, not broader. **Important post-script:** the panel-transparency symptom **persisted after 2c.2 deployed** — proving the selector wasn't the actual cause; see 2c.3 for the real root cause. The 2c.2 selector change is still correct on its own merits (the feature-tag fallback was unsound), it just wasn't the bug we were chasing.
- **2c.3 FluentProvider removal (real fix for transparent SP panels):** Removed the `<FluentProvider theme={webLightTheme}>` wrapper from [src/extensions/phillipsBrand/components/BrandedHeader.tsx](src/extensions/phillipsBrand/components/BrandedHeader.tsx) and dropped the `import { FluentProvider, webLightTheme } from '@fluentui/react-components'`. Solution version `1.0.0.17`. Bundle hash: `phillips-brand-application-customizer_3a85a3f22fdefd7a9ed2.js`.
  - **Root cause:** SharePoint Online's modern UI has been migrating panels (New Item form, web part property pane, Site Information, Highlighted Content config) to Fluent UI v9 internally. Those panels render as portals (siblings of the page DOM, not descendants of our placeholder tree) and resolve their `background-color` via a `--colorNeutralBackground1` theme token cascaded from a `fui-FluentProvider#` class on a SharePoint-owned root `<FluentProvider>`. When our customizer mounted a second `<FluentProvider>` in the Top placeholder, the v9 runtime assigned conflicting `fui-FluentProvider#` class numbers (documented in [microsoft/fluentui#23821](https://github.com/microsoft/fluentui/issues/23821) and [SharePoint/sp-dev-docs#9847](https://github.com/SharePoint/sp-dev-docs/issues/9847)); SP's portaled panels then looked up the wrong provider class, resolved `--colorNeutralBackground1` to nothing, and painted their surface as `transparent`. Compounded by `applyStylesToPortals={true}` (the v9 default), which sent our provider's variables into SP's panel portals.
  - **Why the fix is correct:** `BrandedHeader` does not use a single `@fluentui/react-components` component inside the provider — only native `<div>`, `<img>`, `<nav>`, `<a>` styled by CSS modules. The `<FluentProvider>` was load-bearing for nothing in our code but actively load-bearing for SP's portaled-panel bug. Removing it restores SP's panel backgrounds without changing any visible aspect of our BrandedHeader.
  - **What stayed:** the two global side effects (CSS injection + React render into Top) are unchanged. The component still subscribes to `placeholderProvider.changedEvent`, still caches hub nav in `sessionStorage`, still falls back to `FALLBACK_NAV` on fetch failure.
  - **What to watch for in future work:** if we ever add a real Fluent v9 component to this codebase (Dialog, Drawer, Combobox, etc.), do NOT just re-add `<FluentProvider>` around the whole `BrandedHeader`. Mount the provider as locally as possible (around the specific v9 component) and pass `applyStylesToPortals={false}` to keep our theme from leaking into SP's portals. Verify by opening the List New Item form and the web part property pane after deploy — if their backgrounds go transparent, the provider is conflicting again.
  - **Lesson learned:** when an SPFx Application Customizer mounts a `<FluentProvider>` (v9) inside the SharePoint shell, it isn't a no-op. The provider registers a numbered `fui-FluentProvider#` class globally; SharePoint's own v9 components (in portals) may resolve their theme tokens through that class and break when the numbering doesn't match. The diagnostic principle: when something visible breaks across many SP panels simultaneously after deploying a customizer, suspect the customizer's React tree (especially any provider it mounts) before suspecting CSS selectors. A second-FluentProvider collision causes a broad symptom from a tiny code change.

### Iteration 2c.4 — BrandedHeader active-nav underline reactive to client-side navigation (1.0.17.3)

- **Status:** Deployed and browser-verified. Single-clicking a top-nav item now moves the red active-underline to the clicked item, and browser back/forward updates it too. Previously only a double-click (which degrades to a hard reload) moved it.
- **Solution version at deploy:** `1.0.17.3` (bump from `1.0.17.2`). Bundle hash: `phillips-brand-application-customizer_b07b69f6f3238684b8a9.js`.
- **Symptom & root cause:** The active item was computed from a render-time snapshot `const currentUrl = window.location.href`. The `BrandedHeader` renders once into the persistent `PlaceholderName.Top` slot and is NOT unmounted/remounted on SharePoint modern SPA navigation, so the snapshot went stale and `isNavItemActive` was never re-evaluated. The underline itself (CSS `.active::after`) was always correct — only its input was stale.
- **Fix:** Made the current URL reactive React state — `const [currentUrl, setCurrentUrl] = React.useState(window.location.href)` — updated from a NEW, separate `useEffect` (the nav-fetch effect is untouched) that subscribes to `context.application.navigatedEvent` (SP client-side nav) plus a window `popstate` listener (browser back/forward), with teardown on unmount. The handler reads `window.location.href` fresh on each call (no stale closure). `isNavItemActive`, the per-item call, the plain `<a href>` markup, and `BrandedHeader.module.scss` are all unchanged. Only file changed: [src/extensions/phillipsBrand/components/BrandedHeader.tsx](src/extensions/phillipsBrand/components/BrandedHeader.tsx).
- **`navigatedEvent` subscription gotcha (SPFx 1.23):** `SPEvent.add(observer, handler)` requires the owner to be an `ISPEventObserver`. `context` itself does NOT qualify — `BaseComponentContext` (→ `ExtensionContext` → `ApplicationCustomizerContext`) has no `componentId`. Build the observer from `context.instanceId` + `context.manifest.id`, and do NOT include `manifest` on the observer object — it is `@internal` and absent from the rolled-up public `.d.ts`, so including it is a build-time error (`TS2353`). `add` and `remove` must receive the **same** observer and handler instances — keep both calls inside one effect closure so identity is guaranteed.
- **Design authorization:** No new global side effect — still the two authorized side effects (CSS injection + React render into Top, D026). This only makes the existing Top render reactive to navigation.

## Lessons learned

Add insights here as they emerge from the build process — things that took time to figure out and shouldn't have to be rediscovered.

### Validate the target site type *before* deep-diving the code

The Prompt 4 CSS-injection failure consumed a long debugging session: I audited the build pipeline, manifest schema, package layout, feature XML, ClientSideAssets folder, sp-css-loader output, and runtime CSS injection — all of which were correct. The actual cause was that the original target site (`PartnerExchange-DanSandbox`) is a classic site collection, and SPFx extensions simply don't run on classic sites. The customaction got registered cleanly, SharePoint dutifully shipped it in the page payload, and then the SPFx runtime — which never initialized on that page — never resolved the component manifest. Symptom: zero network requests for the bundle, zero console output, registration looks perfect.

**Rule of thumb for next time:** when SPFx behavior is inexplicable, spin up a fresh modern team site (`m365 spo site add --type TeamSite --url https://<tenant>.sharepoint.com/sites/<name> --title <name> --alias <name>`), deploy the app there, and see if the symptom reproduces. This takes ~2 minutes and definitively answers "is the code the problem or is the target site the problem." Make it one of the first diagnostic moves, not the last. The "Site type validation" bullet in "How to work in this codebase" is the inlined version of this principle.

### Persistent-placeholder UI that depends on the URL must subscribe to `navigatedEvent` (+ `popstate`)

Anything rendered into a persistent Application Customizer placeholder (`PlaceholderName.Top`/`.Bottom`) that depends on the current URL or page must subscribe to `context.application.navigatedEvent` (and a window `popstate` listener for browser back/forward) to stay correct across SharePoint modern SPA navigation. The placeholder React tree is mounted **once** and is NOT remounted on client-side navigation, so a render-time `window.location` read is only correct until the first client-side hop — after that it is stale with no re-render to refresh it. This is exactly what broke the I07 branded-header active-nav underline (fixed in 1.0.17.3 — see "Iteration 2c.4" under Verified behavior): single-click navigated but the underline didn't move; only a double-click, which degrades to a hard reload, appeared to "work."

**Subscription gotcha (SPFx 1.23):** `SPEvent.add(observer, handler)` needs an `ISPEventObserver` owner. `context` does not qualify (no `componentId`); from a functional component, build a minimal stable observer `{ instanceId: context.instanceId, componentId: context.manifest.id, isDisposed: false, dispose: () => {} }`. Do NOT set `manifest` on it (it's `@internal`, absent from the public `.d.ts` → `TS2353`). Pass the SAME observer and handler instances to `add` and `remove` — keep both in one `useEffect` closure so identity matches and teardown actually unsubscribes.
