# Provisioning Runbook

How to deploy the artifacts in this folder to the Phillips Microsoft 365 tenant.

The companion file [provisioning-deployment-guide.md](provisioning-deployment-guide.md) was written assuming PnP PowerShell. This README is the **canonical** runbook for the PnP CLI for Microsoft 365 (`m365`) toolchain that we standardized on per CLAUDE.md ("Deployment: PnP CLI for Microsoft 365 — works cross-platform"). When the two disagree, follow this file.

---

## Step 1 — Register the "Phillips Brand" theme tenant-wide

**Use [phillips-theme-palette-v2.json](phillips-theme-palette-v2.json)** (the 22-slot palette — see "Palette schema" below for why).

### The obvious command does not work

```bash
# DOES NOT WORK for first-time registration on PnP CLI for M365 v9.1.0:
m365 spo theme set --name "Phillips Brand" --theme "$(cat phillips-theme-palette-v2.json)"
# Error: "We couldn't find the theme. Verify the name and try again."
```

**Why:** `m365 spo theme set` in v9.1.0 sends the CSOM `UpdateTenantTheme` method unconditionally — it never calls `AddTenantTheme` first. `UpdateTenantTheme` errors out when the theme does not yet exist. Confirmed by running with `--debug` and inspecting the outgoing `ProcessQuery` body. `--verbose` is not enough to see this; you need `--debug`.

Status: bug in PnP CLI v9.x. Once the theme exists, `m365 spo theme set` works correctly for updates — the regression is only on first-time creation.

### Working approach: call `AddTenantTheme` via the raw CSOM endpoint

Run these from the `provisioning/` directory after `m365 login`:

```bash
# 1. Fetch a fresh request digest from the SPO admin site
DIGEST=$(m365 request \
  --url "https://phillipscorp-admin.sharepoint.com/_api/contextinfo" \
  --method post \
  --resource "https://phillipscorp-admin.sharepoint.com" \
  --accept "application/json;odata=nometadata" \
  | grep -oP '"FormDigestValue":\s*"\K[^"]+')

# 2. Build the AddTenantTheme CSOM payload from the palette JSON
PALETTE=$(cat phillips-theme-palette-v2.json | tr -d '\n' | tr -s ' ')
THEME_JSON="{\"isInverted\":false,\"name\":\"Phillips Brand\",\"palette\":${PALETTE}}"
cat > /tmp/csom-add-theme.xml <<EOF
<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="CLI for Microsoft 365" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><Method Name="AddTenantTheme" Id="11" ObjectPathId="9"><Parameters><Parameter Type="String">Phillips Brand</Parameter><Parameter Type="String">${THEME_JSON}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="9" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/></ObjectPaths></Request>
EOF

# 3. POST the CSOM call
m365 request \
  --url "https://phillipscorp-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery" \
  --method post \
  --resource "https://phillipscorp-admin.sharepoint.com" \
  --body "@/tmp/csom-add-theme.xml" \
  --content-type "text/xml" \
  --x-requestdigest "$DIGEST" \
  --accept "application/json"

# Expected success response (note ErrorInfo: null):
#   [{"SchemaVersion":"15.0.0.0", ..., "ErrorInfo":null, ...}, 10, {"IsNull":false}, 11, true]
```

After it succeeds, verify with:

```bash
m365 spo theme list --output json
# Expect a single entry { "name": "Phillips Brand", "themeJson": "..." }
```

### Updating the palette later

Once "Phillips Brand" exists in the tenant, `m365 spo theme set` works normally for updates:

```bash
m365 spo theme set --name "Phillips Brand" --theme "$(cat phillips-theme-palette-v2.json)"
```

The CSOM workaround above is only needed for the **first-time** registration on a fresh tenant (or after the theme has been deleted).

### Fallback: SharePoint Admin Center UI

If the CLI breaks again on a future tenant or release:

1. Open `https://phillipscorp-admin.sharepoint.com`
2. Settings (top-right gear) → **Site themes** → **+ Add theme**
3. Name: `Phillips Brand`
4. Paste the contents of `phillips-theme-palette-v2.json` into the palette JSON box
5. Save

The admin UI invokes `AddTenantTheme` correctly and is not subject to the CLI bug.

### Palette schema (22 slots vs 30 slots — which file?)

SharePoint's tenant theme API accepts **exactly the 22 slot names** in [phillips-theme-palette-v2.json](phillips-theme-palette-v2.json). The 30-slot variants in `phillips-theme-palette.json` and `phillips-theme.json` add Fluent-UI-specific slots (`accent`, `primaryBackground`, `bodyText`, etc.) that the tenant theme API silently rejects on some builds and accepts on others — behavior is inconsistent across SharePoint releases.

- Use **`phillips-theme-palette-v2.json`** for `AddTenantTheme` / `UpdateTenantTheme` / `m365 spo theme set`
- Keep `phillips-theme.json` only as the design reference (it documents `accent: #00AEEF` and the gold variant — context the 22-slot file does not preserve)

---

## Step 2 — Apply the theme to the dev sandbox site

```bash
m365 spo theme apply \
  --webUrl https://phillipscorp.sharepoint.com/sites/PartnerExchange-DanSandbox \
  --name "Phillips Brand"
```

To revert the site to a stock theme, apply one of the built-ins (e.g. `Blue`).

---

## Step 3 — Site script, Quick Links formatter, list seeding, home page

These steps remain as documented in [provisioning-deployment-guide.md](provisioning-deployment-guide.md) Steps 2–7, with the PowerShell `Connect-PnPOnline` / `Add-PnPListItem` / `Set-PnPView` cmdlets translated to their `m365` equivalents (`m365 spo sitescript add`, `m365 spo listitem add`, `m365 spo list view set`). Translate as needed; the JSON artifacts themselves do not change.

---

## Manual test — verify Step 1 worked

1. Run `m365 spo theme list --output json`. Confirm exactly one entry named `Phillips Brand` is returned with a `themeJson` that includes `"themePrimary":"#F9423A"`.
2. Open `https://phillipscorp.sharepoint.com/sites/PartnerExchange-DanSandbox` in a browser, logged in as a site admin.
3. Click the gear icon (top right) → **Change the look** → **Theme**.
4. Confirm **Phillips Brand** appears in the theme picker alongside the built-in themes.
5. Click **Phillips Brand**, then **Save**. The site re-renders with red accents.
6. Inspect any link or primary button on the page — the primary color should be `#F9423A` (Phillips Red), not the default SharePoint blue.
7. (Optional rollback) Re-open **Change the look → Theme**, pick `Blue`, click **Save** to restore the default for this dev sandbox site.

If all seven steps pass, Step 1 is complete and Step 2 has been validated end-to-end.
