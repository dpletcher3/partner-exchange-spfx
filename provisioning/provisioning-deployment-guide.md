# Phillips Intranet — Deployment Guide

JSON artifacts in this folder, in the order you'll use them:

| File | Purpose |
|---|---|
| `phillips-theme.json` | Fluent UI theme palette — applied tenant-wide via `Add-SPOTheme` |
| `phillips-site-script.json` | Site script — theme + navigation + lists + branding |
| `quick-links-tile-formatting.json` | View formatter to render Quick Links as branded tiles |
| `quick-links-seed-data.json` | The 5 starter tiles from the screenshot |
| `homepage-canvas.json` | Reference layout for the home page (sections, columns, web parts) |

## Prerequisites

```powershell
# One-time install (PowerShell 7+ recommended)
Install-Module -Name "PnP.PowerShell" -Scope CurrentUser
Install-Module -Name "Microsoft.Online.SharePoint.PowerShell" -Scope CurrentUser

# Register a PnP app registration if you don't already have one
# (one-time per tenant — you likely have one from the D365 work)
Register-PnPManagementShellAccess
```

## Step 1 — Apply the theme tenant-wide

```powershell
Connect-SPOService -Url "https://phillipscorp-admin.sharepoint.com"

$themeJson = Get-Content -Path ".\phillips-theme.json" -Raw | ConvertFrom-Json
$palette = @{}
$themeJson.palette.PSObject.Properties | ForEach-Object { $palette[$_.Name] = $_.Value }

Add-SPOTheme -Identity "Phillips Brand" -Palette $palette -IsInverted $false -Overwrite
```

After this, "Phillips Brand" is selectable under **Change the look → Theme** on any site, and the site script will reference it by name.

> **Brand Center note**: If your tenant has SharePoint Premium / Microsoft 365 Copilot licensing, also set up Brand Center (admin center → Brand Center) and upload Barlow Condensed + Montserrat as the corporate fonts. The theme palette and the Brand Center fonts together get you most of the way to the screenshot without any CSS.

## Step 2 — Register the site script and site design

```powershell
Connect-SPOService -Url "https://phillipscorp-admin.sharepoint.com"

$siteScriptContent = Get-Content -Path ".\phillips-site-script.json" -Raw

$siteScript = Add-SPOSiteScript `
  -Title "Phillips Intranet Provisioning" `
  -Description "Applies Phillips brand theme, navigation, and core lists" `
  -Content $siteScriptContent

Add-SPOSiteDesign `
  -Title "Phillips Intranet Home Site" `
  -WebTemplate "68" `
  -SiteScripts $siteScript.Id `
  -Description "Phillips-branded communication site for intranet hubs" `
  -PreviewImageUrl "https://phillipscorp.sharepoint.com/sites/intranet/SiteAssets/site-design-preview.png"

# WebTemplate IDs:
#   64 = Team Site
#   68 = Communication Site  (use this for the home site and hub sites)
```

## Step 3 — Create the home site and apply the design

```powershell
# Create the communication site that will become the home site
New-SPOSite `
  -Url "https://phillipscorp.sharepoint.com/sites/intranet" `
  -Owner "dpletcher@phillipscorp.com" `
  -StorageQuota 1024 `
  -Title "Phillips Portal Exchange" `
  -Template "SITEPAGEPUBLISHING#0"

# Promote it to the org-wide home site
Set-SPOHomeSite -HomeSiteUrl "https://phillipscorp.sharepoint.com/sites/intranet"

# Apply the site design (re-applicable; safe to run again after edits)
Invoke-SPOSiteDesign `
  -Identity $siteDesignId `
  -WebUrl "https://phillipscorp.sharepoint.com/sites/intranet"
```

## Step 4 — Format the Quick Links list as tiles

```powershell
Connect-PnPOnline -Url "https://phillipscorp.sharepoint.com/sites/intranet" -Interactive

# Apply the tile view formatter to the default view of the Quick Links list
$tileFormatJson = Get-Content -Path ".\quick-links-tile-formatting.json" -Raw

Set-PnPView `
  -List "Quick Links" `
  -Identity "All Items" `
  -Values @{ CustomFormatter = $tileFormatJson }
```

## Step 5 — Seed the Quick Links list

```powershell
$seed = Get-Content -Path ".\quick-links-seed-data.json" -Raw | ConvertFrom-Json

foreach ($item in $seed.items) {
  Add-PnPListItem -List "Quick Links" -Values @{
    Title             = $item.Title
    QLink             = "$($item.QLink.Url), $($item.QLink.Description)"
    IconName          = $item.IconName
    ShortDescription  = $item.ShortDescription
    SortOrder         = $item.SortOrder
    AudienceGroup     = $item.AudienceGroup
  }
}
```

## Step 6 — Build the home page

The site design only provisions structure, not page content. Use the canvas reference to build the home page:

```powershell
# Create the modern page
Add-PnPPage -Name "Home" -LayoutType Home -PromoteAs HomePage

# Add a two-column-right section
Add-PnPPageSection -Page "Home" -SectionTemplate TwoColumnRight -Order 1

# Add the Hero web part to column 1 (configure tiles in the UI afterward)
Add-PnPPageWebPart -Page "Home" -DefaultWebPartType Hero -Section 1 -Column 1

# Add the List web part referencing Quick Links
$quickLinksList = Get-PnPList -Identity "Quick Links"
Add-PnPPageWebPart -Page "Home" `
  -DefaultWebPartType List `
  -Section 1 -Column 1 `
  -WebPartProperties @{ selectedListId = $quickLinksList.Id; selectedView = "All Items" }

# Add Highlighted Content for The Phillips Loop
Add-PnPPageWebPart -Page "Home" -DefaultWebPartType ContentRollup -Section 1 -Column 1

# Add News
Add-PnPPageWebPart -Page "Home" -DefaultWebPartType News -Section 1 -Column 1

# Add Viva Engage to column 2 (The Hive)
Add-PnPPageWebPart -Page "Home" -DefaultWebPartType YammerEmbed -Section 1 -Column 2

# Publish
Set-PnPPage -Identity "Home" -Publish
```

Then go into the page in edit mode to set the hero image, configure the News source filters to your hub-associated sites, and point Viva Engage at the "Phillips Hive" community.

## Step 7 — Roll out to hubs

Once the home site looks right, register it as a hub site, then apply the same site design when provisioning each business unit / function site so they inherit the branding and navigation:

```powershell
Register-SPOHubSite -Site "https://phillipscorp.sharepoint.com/sites/intranet"

# When provisioning each new site:
New-SPOSite -Url "https://phillipscorp.sharepoint.com/sites/hr" -Template "SITEPAGEPUBLISHING#0" ...
Invoke-SPOSiteDesign -Identity $siteDesignId -WebUrl "https://phillipscorp.sharepoint.com/sites/hr"
Add-SPOHubSiteAssociation -Site "https://phillipscorp.sharepoint.com/sites/hr" -HubSite "https://phillipscorp.sharepoint.com/sites/intranet"
```

## What this does NOT cover (the SPFx gap)

The screenshot has three elements that no OOTB SharePoint configuration can produce. If you want them, scope them as a small SPFx solution after the OOTB rollout is stable:

1. **Personalized "Good Morning, [User]!" hero** with live date/time and time-zone selector — needs an SPFx web part using `@microsoft/sp-page-context` for the user identity and a small React component for the clock.
2. **Custom mega menu** with featured cards and multi-column layout — needs an SPFx Application Customizer.
3. **Branded footer** with social/feedback links — also an SPFx Application Customizer.

These three add up to roughly a 2–3 sprint SPFx project. Until then, the OOTB Hero web part with the sky/cloud image and a static "Welcome to Phillips Portal Exchange" title gets you 80% of the visual impact.

## Validation checklist before user testing

- [ ] Theme appears in **Change the look → Theme** on a test site
- [ ] Site script provisioning completes without errors (check `Get-SPOSiteDesignRun -WebUrl ... | Get-SPOSiteDesignRunStatus`)
- [ ] Top navigation shows: Home, Our Culture, Our Partners, Dashboard, The Hub
- [ ] Quick Links list renders as tiles, not the default grid view
- [ ] Tile icons render in Phillips Red on light pink background
- [ ] Hero web part appears with a custom background image, not the default
- [ ] News web part shows curated stories with audience-targeted filtering
- [ ] Viva Engage web part loads "The Hive" community feed
- [ ] Site loads in under 2s on first paint (test from the India office too)
