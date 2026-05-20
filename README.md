# Partner Exchange SPFx

Custom SharePoint Framework solution for the Phillips Corporation intranet ("Partner Exchange"). Contains one Application Customizer that injects brand CSS site-wide, and one custom web part (a personalized hero) for content that built-in SharePoint web parts can't reproduce.

This repository is built using the Building Apps with Claude methodology. The build is executed by Claude Code following the numbered prompts in `PROMPTS.md`. The architectural decisions, brand tokens, and constraints are locked in `CLAUDE.md`.

## What's in this repo

| File / folder | Purpose |
|---|---|
| `CLAUDE.md` | Locked decisions, brand tokens, hard constraints. Read first. |
| `PROMPTS.md` | Numbered build sequence. Claude Code executes one at a time. |
| `provisioning/` | Theme, site script, list formatters (JSON). Deploy with PnP CLI. |
| `src/` | SPFx solution source (created by Prompt 1). |
| `config/` | SPFx config files (created by Prompt 1). |
| `.github/workflows/` | GitHub Actions CI (created by Prompt 7). |

## Setup

```bash
# Clone
git clone https://github.com/dpletcher3/partner-exchange-spfx.git
cd partner-exchange-spfx

# Node version (uses .nvmrc)
nvm install
nvm use

# Install global tooling once per machine
npm install -g yo @microsoft/generator-sharepoint gulp-cli @pnp/cli-microsoft365

# Install Claude Code if not present
# Follow https://docs.claude.com/en/docs/claude-code/quickstart
```

## The build workflow

1. Open Claude Code in this directory: `claude`
2. Paste the kickoff message from the bottom of `PROMPTS.md`
3. Claude Code reads `CLAUDE.md` and `PROMPTS.md` in full, then executes Prompt 1 only
4. When Prompt 1 finishes, it outputs explicit manual test steps — you run them
5. Confirm verification, then tell Claude Code "continue with Prompt 2"
6. Repeat until all prompts are complete

The discipline matters: don't let Claude Code chain multiple prompts without verification. Each prompt's manual test steps are the only independent check that the work actually landed.

## Provisioning files (existing JSON)

The JSON artifacts produced before this repo was started (theme, site script, list formatters) should be copied into `provisioning/` before Prompt 1 begins:

```bash
mkdir -p provisioning
# Move them in:
cp /path/to/phillips-theme.json provisioning/
cp /path/to/phillips-site-script.json provisioning/
cp /path/to/quick-links-tile-formatting.json provisioning/
cp /path/to/quick-links-seed-data.json provisioning/
cp /path/to/homepage-canvas.json provisioning/
cp /path/to/provisioning-deployment-guide.md provisioning/README.md
```

These get deployed once (tenant theme, site script registration, list provisioning) via PnP CLI — see `provisioning/README.md`. They are not part of the SPFx package.

## What's in scope for v1

- Application Customizer that styles all built-in SharePoint web parts to match the Phillips brand (rounded corners, pill buttons, eyebrow pattern, pink-tinted icon backgrounds)
- One custom web part: personalized hero (greeting + clock + time-zone selector)
- GitHub Actions build pipeline producing `.sppkg` artifacts
- Manual deploy to App Catalog via PnP CLI

## What's deferred to v1.5

- Featured Story web part (the Machining-pattern card from phillipscorp.com)
- PnP Modern Search
- Alert banner driven by a SharePoint list
- People directory upgrade
- Multi-language support

## Verifying things actually work

Claude Code cannot independently verify its own output. The manual test steps at the end of every prompt are not optional — they're the only real check. If a prompt finishes without producing manual test steps, that's a defect: stop and ask Claude Code to add them before continuing.
