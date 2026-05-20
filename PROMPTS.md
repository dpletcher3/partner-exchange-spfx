# PROMPTS.md

Numbered execution sequence for the Partner Exchange SPFx build. Claude Code executes one prompt at a time and stops at the end of each, outputting Manual Test steps for the developer to run before continuing.

Every prompt ends with a Manual Test section that the developer must actually run. Claude Code cannot independently verify its own output.

---

## Prompt 1 — Scaffold the SPFx solution

Run the SharePoint Framework Yeoman generator to scaffold the solution into the current directory. Use these exact answers:

- Solution name: `partner-exchange-spfx`
- Target environment: `SharePoint Online only (latest)`
- Component to create: `No JavaScript framework` initially (we'll add React per-component to keep the base lean — confirm this is still correct against current SPFx 1.20 conventions before proceeding)
- Skip the first component prompt (we'll generate components individually in later prompts)

After scaffolding:

1. Verify the generated files match what SPFx 1.23.x produces: `package.json`, `tsconfig.json`, `eslint.config.js`, `config/config.json`, `config/package-solution.json`, `config/serve.json`, `config/rig.json`, `.gitignore` (note: no `gulpfile.js` — SPFx 1.23 uses Heft, not gulp)
2. Update `config/package-solution.json` so the solution name is `partner-exchange-spfx`, the solution title is `Partner Exchange Brand & Components`, and the developer name is `Phillips Corporation`
3. Run `npm install` and confirm it completes without errors
4. Run `npm audit` and resolve any high or critical vulnerabilities (do not auto-fix breaking changes; surface them as a question if `npm audit fix` would bump majors)

### Manual Test

1. Run `ls` in the repo root (PowerShell) or `ls -la` (bash). Confirm the SPFx files listed above all exist.
2. Open `config/package-solution.json` in your editor. Confirm the `name`, `title`, and `developer.name` fields match the values above.
3. Run `npx heft trust-dev-cert` and accept the Windows certificate prompt. Console prints `Successfully trusted development certificate.`
4. Run `npm start` (which executes `heft start --clean`). Wait until you see `[start] Server started at https://localhost:4321` in the terminal.
5. Open `https://phillipscorp.sharepoint.com/_layouts/15/workbench.aspx` in a browser. Accept the "Load debug scripts?" prompt. Confirm the workbench canvas loads.
6. Stop the dev server with Ctrl+C in the terminal.

If all six steps pass, Prompt 1 is complete.

---

## Prompt 2 — Create the brand token foundation

Create `src/styles/_tokens.scss` containing all brand tokens defined in `CLAUDE.md` (colors, typography, spacing scale, border radius scale). Export them as:

- CSS custom properties on `:root` (so they cascade to OOTB SharePoint web parts via the customizer's stylesheet)
- SCSS variables (so they can be used in components' SCSS modules without going through `var()`)

Create `src/styles/_mixins.scss` containing the following mixin:

- `@mixin phil-eyebrow` — produces the eyebrow label pattern (red 16px dash + uppercase red text with 2px letter spacing, 9px font size)

Create `src/styles/index.scss` that imports both files. This will be the entry point the customizer uses.

### Manual Test

1. Open `src/styles/_tokens.scss`. Confirm every token from CLAUDE.md's Brand Tokens section is present, with the correct value, as both a CSS custom property and a SCSS variable.
2. Open `src/styles/_mixins.scss`. Confirm the `phil-eyebrow` mixin is present.
3. Create a temporary test file `src/styles/test.scss` that imports `_tokens.scss`, applies `--phil-red` and the `phil-eyebrow` mixin to a test class. Compile it with `npx sass src/styles/test.scss`. Confirm the output includes `#F9423A` and the eyebrow rules.
4. Delete `src/styles/test.scss`.

If all four steps pass, Prompt 2 is complete.

---

## Prompt 3 — Generate the Application Customizer

Run the SPFx Yeoman generator to add an Application Customizer extension to the existing solution:

```
yo @microsoft/sharepoint --solution-name partner-exchange-spfx --component-type extension --extension-type ApplicationCustomizer --component-name PhillipsBrand
```

Verify the generator created `src/extensions/phillipsBrand/PhillipsBrandApplicationCustomizer.ts` and supporting files. Open the main `.ts` file and:

1. Remove the default "Top" and "Bottom" placeholder code — the customizer will not render any UI, only inject CSS
2. In the `onInit()` method, log a message confirming the customizer initialized (this gets removed in Prompt 4 but proves wiring works now)
3. Update `config/package-solution.json` so the feature framework references this extension correctly
4. Update `config/serve.json` so debugging the extension works against the dev sandbox: `https://phillipscorp.sharepoint.com/sites/spfx-extension-test`

### Manual Test

1. Run `npm start` (which executes `heft start --clean`). Wait for `[start] Server started at https://localhost:4321`. (Heft does not auto-open the browser — that was a gulp behavior.)
2. Open the URL from `serve.json` `initialPage` in your browser. Accept the "Load debug scripts?" prompt.
3. Open the browser's developer console.
4. Confirm the console shows the log message you added in `onInit()`.
5. Stop the dev server with Ctrl+C.

If all five steps pass, Prompt 3 is complete.

---

## Prompt 4 — Implement the CSS injection

This is the visible-progress prompt. After this, OOTB SharePoint web parts on the test site will look Phillips-branded.

In the customizer's `onInit()` method:

1. Rename the brand stylesheet entry point from `src/styles/index.scss` to `src/styles/index.global.scss`. The `.global.scss` extension matches Heft's `nonModuleFileExtensions` pattern (configured in `@microsoft/spfx-web-build-rig`), so the compiled CSS is auto-injected into `<head>` by sp-css-loader at runtime instead of being treated as a CSS Module (which would hash class names and break the OOTB selectors).
2. Side-effect import from the customizer TS: `import '../../styles/index.global.scss';`. sp-css-loader handles injection and dedup automatically — the same SCSS module is loaded exactly once per page session, even across SPA navigations. No manual `<style>` tag creation needed.
3. Inside SCSS partials, always use `@use './tokens' as *;` (or `@use './tokens' as t;` for namespaced access) — never `@import` (Dart Sass 3.0 removes it and double-emits `:root`).

In `src/styles/_overrides.scss` (new partial, added to `index.global.scss` via `@use './overrides';` and consuming tokens via `@use './tokens' as *;`), write the brand overrides. Cover these built-in web parts at minimum:

- **Hero web part:** `--phil-radius-xl` corners, ensure overlay is `rgba(0, 50, 80, 0.35)` over images
- **News web part:** `--phil-radius-lg` corners on cards, eyebrow pattern applied to section titles
- **Quick Links web part:** card backgrounds `--phil-white`, `--phil-gray-200` border, `--phil-radius-lg` corners (note: even with this, the JSON list-formatter approach in `provisioning/quick-links-tile-formatting.json` produces a better result — the OOTB Quick Links override here is a fallback)
- **Highlighted Content web part:** `--phil-radius-lg` corners on cards
- **Button web part:** pill shape (`--phil-radius-pill`), `--phil-red` background for primary, white outline for secondary
- **Section backgrounds:** ensure colored section backgrounds use `--phil-red-tint` for "Neutral" alternate sections

Each override block must include a comment explaining what SharePoint default it's overriding and why. Selectors stay as low-specificity as possible.

### Manual Test

1. Run `npm run build` (which executes `heft test --clean --production && heft package-solution --production`). Confirm a `.sppkg` is produced under `sharepoint/solution/`.
2. Upload the `.sppkg` to the App Catalog using PnP CLI: `m365 spo app add --filePath sharepoint/solution/partner-exchange-spfx.sppkg --overwrite --appCatalogScope tenant`
3. Deploy it: `m365 spo app deploy --name partner-exchange-spfx.sppkg --appCatalogScope tenant`
4. Add the customizer to your test site: `m365 spo customaction add --webUrl https://phillipscorp.sharepoint.com/sites/spfx-extension-test --name "Phillips Brand" --location "ClientSideExtension.ApplicationCustomizer" --clientSideComponentId YOUR_COMPONENT_ID`
5. Open the test site home page. Add (or confirm already present) a Hero web part, News web part, Quick Links web part, and a Button. Configure each with placeholder content.
6. Confirm visually: Hero has rounded XL corners. News cards have rounded LG corners. Quick Links cards have brand styling. Button is pill-shaped with Phillips Red background.
7. Open DevTools → Elements → `<head>`. Confirm at least one `<style>` element contains `--phil-red` (search the inline style content). The element's `id` will be a sp-css-loader-generated hash, not `phil-brand` — that's expected; sp-css-loader auto-injects and there's no manual DOM call to set a custom id.
8. Navigate to a different page on the site. Confirm the brand `<style>` is still present and the page hasn't accumulated a second copy of it — sp-css-loader's module cache prevents duplicate injection across SPA navigations.

If all eight steps pass, Prompt 4 is complete. **This is the moment to commit and PR — the customizer alone is a meaningful v1 deliverable.**

---

## Prompt 5 — Generate the Personalized Hero web part scaffold

Run the SPFx Yeoman generator to add a React-based web part to the existing solution:

```
yo @microsoft/sharepoint --solution-name partner-exchange-spfx --component-type webpart --framework react --component-name PersonalizedHero
```

After scaffolding:

1. Confirm `src/webparts/personalizedHero/` exists with the standard SPFx structure
2. Update the web part's display name to "Personalized Hero" and description to "Greeting that varies by time of day, with the current user's name, a live clock, and a time-zone selector. Designed for use at the top of the home page."
3. Add the web part to the appropriate group in the web part picker (e.g., "Phillips" group — configure in the web part manifest)
4. Replace the default React component with a minimal placeholder that renders the user's display name from `this.context.pageContext.user.displayName`
5. In the web part's `.module.scss`, add `@use '../../styles/tokens' as *;` at the top so the brand SCSS variables (`$phil-red`, `$phil-radius-xl`, etc.) are available. Do not use the deprecated `@import` form. For values that don't need compile-time resolution, prefer the `var(--phil-*)` CSS custom properties — they cascade from `:root` and don't require an `@use` declaration.

### Manual Test

1. Run `npm start` against the test site (the URL in `config/serve.json` should point to it).
2. Open the test site's home page in edit mode.
3. Add a new web part — search the picker for "Personalized Hero".
4. Confirm the web part appears in the picker under the "Phillips" group with the description from step 2.
5. Add it to the page. Confirm the placeholder renders the current user's display name.
6. Stop the dev server with Ctrl+C.

If all six steps pass, Prompt 5 is complete.

---

## Prompt 6 — Implement the Personalized Hero component

Replace the placeholder React component with the full Personalized Hero, matching the mockup from the design phase:

- Background: `--phil-blue` with `rgba(0, 50, 80, 0.35)` darkening overlay
- Top-left: Eyebrow pattern reading "GOOD MORNING" / "GOOD AFTERNOON" / "GOOD EVENING" / "WORKING LATE" based on local time (rules in CLAUDE.md → Domain rules → Greeting by time of day)
- Top-right: Time-zone selector pill button (US Eastern default, dropdown with Eastern, Central, Mountain, Pacific, India Standard, GMT — Phillips operates across these)
- Center-bottom: Large display heading "Welcome back, {firstName}." (extract first name from display name; if display name has no space, fall back to the full name)
- Below heading: Live clock showing day-of-week, date, and time in selected time zone, updating every minute (not every second — every-second updates are visually noisy)
- Border-radius: `--phil-radius-xl`
- Height: 160px on desktop, 120px on mobile

Use `react-hooks` (`useState`, `useEffect`) for the clock and time-zone state. No external date libraries — use the browser's `Intl.DateTimeFormat` API with the selected time zone.

Make the time-zone selection persist via `localStorage` so the user's choice survives reloads (key: `phil-hero-tz`).

### Manual Test

1. Run `npm start` against the test site.
2. Open the home page in edit mode. The Personalized Hero web part should already be on the page from Prompt 5; if not, add it.
3. Verify the greeting reads correctly for the current time (e.g., "Good morning" if it's 9am local).
4. Verify your name appears as "Welcome back, {firstName}."
5. Verify the clock shows the current day, date, and time.
6. Click the time-zone selector. Verify the dropdown shows all six time zones.
7. Select a different time zone. Verify the clock updates to that time zone immediately.
8. Reload the page. Verify the time zone you selected persists (`localStorage`).
9. Wait 60 seconds. Verify the clock advances by one minute (not 60 seconds — minutes only).
10. Resize the browser window to ~400px wide. Verify the hero remains readable and the height shrinks to 120px.
11. Stop the dev server with Ctrl+C.

If all eleven steps pass, Prompt 6 is complete.

---

## Prompt 7 — GitHub Actions CI

Create `.github/workflows/build.yml` that:

1. Triggers on push to any branch and on PR to `main`
2. Runs on `ubuntu-latest`
3. Uses `actions/setup-node@v4` with `node-version-file: '.nvmrc'` (which currently resolves to Node 22; ensures CI tracks the repo pin automatically). `ubuntu-latest` runners have Node 22 LTS pre-installed but `setup-node` makes the version explicit and lockable.
4. Caches `~/.npm` (the npm cache) keyed on `package-lock.json` hash — this is the modern equivalent of caching `node_modules` and pairs better with `npm ci`
5. Runs `npm ci`
6. Runs `npm run build` — this is the single Heft command that replaces the old `gulp build && gulp bundle --ship && gulp package-solution --ship` chain. The npm script is `heft test --clean --production && heft package-solution --production`, which runs lint, type-check, bundle, and package in one pass.
7. Uploads the `.sppkg` from `sharepoint/solution/` as a GitHub Actions artifact named `partner-exchange-spfx-sppkg`
8. On `main` only: tags the artifact with the commit SHA so deploys are traceable

Also create `.github/dependabot.yml` configured to weekly-check npm dependencies. Group SPFx-related packages (anything starting with `@microsoft/sp-`) into a single PR so SPFx upgrades happen atomically.

### Manual Test

1. Commit and push the workflow files to a new branch.
2. Open GitHub. Confirm the Actions tab shows the workflow running.
3. Wait for the workflow to complete. Confirm it succeeded.
4. Click into the run. Confirm the `partner-exchange-spfx-sppkg` artifact is downloadable.
5. Download the artifact, unzip it, and confirm the `.sppkg` is inside.
6. Open the `.sppkg` (it's a zip) and confirm the components are correctly bundled.

If all six steps pass, Prompt 7 is complete.

---

## Prompt 8 — First production deployment

This prompt is operational, not code. Document the deployment runbook as `docs/DEPLOY.md` so future deploys are repeatable.

Contents of `docs/DEPLOY.md`:

1. Prerequisites: PnP CLI for M365 installed, authenticated as a tenant admin
2. Step-by-step PnP CLI commands to:
   - Upload the `.sppkg` to the App Catalog
   - Deploy the solution tenant-wide
   - Add the Application Customizer to the target site(s)
   - Verify the customizer loads on the site
3. Rollback procedure:
   - Remove the custom action: `m365 spo customaction remove`
   - Retract the solution: `m365 spo app retract`
   - Remove the solution from the catalog: `m365 spo app remove`
4. Troubleshooting common issues:
   - CSP errors (usually CDN paths)
   - Custom action not appearing (cache, app catalog scope)
   - Personalized Hero not appearing in web part picker (solution-level vs site-level web parts)

### Manual Test

1. Build a fresh `.sppkg` locally: `npm run build` (executes `heft test --clean --production && heft package-solution --production`)
2. Following only the steps in `docs/DEPLOY.md` (don't reference anything outside that file), deploy to a clean test site.
3. Confirm the customizer activates and styling appears.
4. Confirm the Personalized Hero web part can be added to a page.
5. Following only the rollback procedure in `docs/DEPLOY.md`, retract the solution and remove it from the catalog.
6. Confirm the test site returns to default SharePoint styling.

If all six steps pass, Prompt 8 is complete.

---

## Prompt 9 — Update the Quick Links list formatter to match the new design

Edit `provisioning/quick-links-tile-formatting.json` to match the rounded-corner, icon-in-square design from the second mockup (not the original circle-icon version):

- Tile border-radius: 12px
- Icon background: rounded square 8px, `--phil-red-tint` color
- Icon color: `--phil-red`
- Title: sentence case (not all-caps), `--phil-black` color, font-size 11px, weight 500
- Hover: subtle border color change to `--phil-red`, no transform

Re-apply the formatter to the Quick Links list on the test site using PnP CLI:

```
m365 spo list view set --webUrl https://phillipscorp.sharepoint.com/sites/spfx-extension-test --listTitle "Quick Links" --title "All Items" --CustomFormatter "@./provisioning/quick-links-tile-formatting.json"
```

### Manual Test

1. Open the test site's Quick Links list in the browser.
2. Confirm the tiles render with: 12px corners, 8px rounded-square icon backgrounds in pink-tint, brand-red icons, sentence-case titles.
3. Hover over a tile. Confirm the border color changes to Phillips Red without any transform/movement.
4. Open the home page. Confirm the embedded Quick Links view shows the new design.

If all four steps pass, Prompt 9 is complete. **All v1 prompts done.**

---

## Kickoff message

Copy this exact text into Claude Code in the repo directory to begin Prompt 1:

> Read CLAUDE.md and PROMPTS.md in full before doing anything. Then execute Prompt 1 only. When Prompt 1 is complete and its acceptance criteria pass, output a "Manual test" section with specific, ordered steps I should take to verify the work myself — exact commands to run, files to open, buttons to click, inputs to try, and the observable result I should see at each step. Do not describe what should happen in the abstract; tell me what to do and what I'll see if it worked. Then stop and wait for me to run those steps and tell you to continue with Prompt 2. Do not auto-continue to subsequent prompts.
