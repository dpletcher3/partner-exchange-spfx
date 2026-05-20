# CLAUDE.md

## What this project is

Partner Exchange SPFx is the custom SharePoint Framework solution that gives the Phillips Corporation intranet its visual identity and adds the one piece of functionality built-in SharePoint web parts cannot reproduce. It contains exactly two SPFx components: an Application Customizer that injects brand CSS site-wide (so all OOTB Hero, News, Quick Links, Highlighted Content, and Button web parts pick up Phillips styling automatically), and a Personalized Hero web part (greeting that varies by time of day, current user's display name, live clock, time-zone selector). Everything else — page composition, news authoring, quick link management, social posts, search — uses standard SharePoint web parts that content authors already know how to use.

## Locked architectural decisions

- **SPFx version:** 1.20.x (latest stable at build time — verify with `npx @microsoft/generator-sharepoint --version` before scaffolding)
- **React:** 17 (SPFx's pinned version — do not upgrade to 18)
- **TypeScript:** 5.x
- **UI library:** Fluent UI v9 (`@fluentui/react-components`)
- **Node:** 18 LTS, pinned via `.nvmrc`
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

## How to work in this codebase

- **Every prompt completion in this project ends with a Manual Test section — numbered steps the developer runs to verify the work, written so they could be followed by someone who didn't see the code change. No prompt is considered complete without it.**
- **Clarity over cleverness in CSS.** Verbose selectors that future Dan can read beat tight ones future Dan can't.
- **Comments in SCSS explain *why* a rule exists,** especially overrides of SharePoint defaults. Future maintainers won't know that `.ms-CommandBar` overrides exist to remove the dated drop shadow unless we tell them.
- **Component files stay under 200 lines.** If a hero component grows past that, split it into subcomponents.
- **Conventional commits:** `feat:`, `fix:`, `chore:`, `docs:`, `refactor:`, `style:`. Prefix with the scope when useful: `feat(hero): add time-zone selector`.
- **PRs reference the prompt number** they completed: "Closes Prompt 4 (CSS injection)."
- **Don't bypass type errors.** No `@ts-ignore`, no `any`. If TypeScript complains, fix the underlying problem.
- **Test in the SharePoint Workbench first**, then on a real test site. Workbench catches structural issues; only real sites catch CSP, theme, and OOTB-web-part interaction issues.

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

(empty)
