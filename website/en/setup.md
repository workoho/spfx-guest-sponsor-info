---
layout: doc
lang: en
title: Setup Guide
permalink: /en/setup/
description: >-
  Step-by-step setup guide for the Guest Sponsor Info web part
  and Guest Sponsor API — SharePoint and Azure setup.
lead: >-
  Initial setup and configuration reference for
  SharePoint and Azure administrators.
github_doc: deployment.md
---

## SharePoint Setup

### Install from Microsoft AppSource

The web part is available in the
[**Microsoft commercial marketplace (AppSource)**](https://appsource.microsoft.com/).
Installing from there deploys it tenant-wide via the Tenant App Catalog — no
Site Collection App Catalog or file upload required.

**Install via SharePoint Admin Center:**

1. Open **SharePoint Admin Center → More features → Apps → Open**.
2. Click **Get apps from marketplace** and search for *Guest Sponsor Info*.
3. Select the app and click **Get it now**.

The solution uses `skipFeatureDeployment: false` — the web part does **not**
become available tenant-wide automatically. After the Tenant App Catalog
installation, a Site Collection Administrator must add the app to each site
explicitly: **Site Contents → Add an app → Guest Sponsor Info**.
This is intentional and prevents accidental installation on unintended sites.

The web part requests **no Microsoft Graph permissions** of its own — the
**API access** queue will remain empty. All Graph calls are made server-side
by the companion Azure Function using its Managed Identity.

### Make the web part accessible to guests

When installed via AppSource or the Tenant App Catalog, the web part JavaScript
bundle is served from the Tenant App Catalog's `ClientSideAssets` library.
B2B guest users cannot access this library before authenticating to the host
tenant, which is not guaranteed before the page load. If guests cannot load
the bundle, the web part silently fails to render.

The web part's built-in **Guest Accessibility** diagnostics panel (property
pane) detects the current scenario and shows the result of each check with a
recommendation.

**Option A — Enable the Office 365 Public CDN (recommended)**

When the Office 365 Public CDN is enabled, SharePoint replicates web part
bundles to Microsoft's edge CDN (`publiccdn.sharepointonline.com`), which is
accessible anonymously — no SharePoint authentication required. This is the
most reliable approach for guest users.

**Required role:** SharePoint Administrator.

```powershell
# SharePoint Online Management Shell (Windows):
Connect-SPOService -Url "https://<tenant>-admin.sharepoint.com"
Set-SPOTenantCdnEnabled -CdnType Public -Enable $true

# Verify the ClientSideAssets origin is included (added by default):
Get-SPOTenantCdnOrigins -CdnType Public
# Expected output includes: */CLIENTSIDEASSETS
```

If `*/CLIENTSIDEASSETS` is missing, add it:

```powershell
Add-SPOTenantCdnOrigin -CdnType Public -OriginUrl "*/CLIENTSIDEASSETS"
```

```powershell
# PnP PowerShell (cross-platform):
Connect-PnPOnline -Url "https://<tenant>-admin.sharepoint.com" `
    -ClientId "<your-pnp-app-client-id>" -Interactive
Set-PnPTenantCdnEnabled -CdnType Public -Enable $true
```

> CDN propagation takes **15-30 minutes**. Once active, the bundle URL changes
> to `publiccdn.sharepointonline.com` automatically — no reconfiguration needed.

**Option B — Grant Everyone read access to the Tenant App Catalog**

If enabling the Public CDN is not possible, grant the built-in **Everyone**
group read access to the Tenant App Catalog site instead.

**Required roles:** SharePoint Administrator and Site Collection Administrator
on the Tenant App Catalog site
(`https://<tenant>.sharepoint.com/sites/appcatalog`).

```powershell
# SharePoint Online Management Shell (Windows):
Connect-SPOService -Url "https://<tenant>-admin.sharepoint.com"
Add-SPOUser -Site "https://<tenant>.sharepoint.com/sites/appcatalog" `
    -LoginName "c:0(.s|true" -Group "App Catalog Visitors"
```

```powershell
# PnP PowerShell (connect to the App Catalog site directly):
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com/sites/appcatalog" `
    -ClientId "<your-pnp-app-client-id>" -Interactive
Add-PnPGroupMember -LoginName "c:0(.s|true" -Group "App Catalog Visitors"
```

> **Limitation:** Only covers guests who have already authenticated to the host
> tenant. The Public CDN (Option A) does not have this limitation.

For an advanced alternative (Site Collection App Catalog, no marketplace), see
the full [setup guide on GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/deployment.md#option-c--use-a-site-collection-app-catalog).

### Verify guest access to the landing page site

Guests need at least **Read** (Visitor) permission on the landing page site.
Rather than a dynamic Entra group — which can take up to 24 hours to reflect
new members — use the built-in **Everyone** group. It covers every
authenticated user including B2B guests who have accepted their invitation,
and takes effect immediately.

The *Everyone* group is controlled by the `ShowEveryoneClaim` tenant setting,
which defaults to `$false` on tenants provisioned after March 2018. If
*Everyone* does not appear in the People Picker, enable it first:

```powershell
# SharePoint Online Management Shell (Windows):
(Get-SPOTenant).ShowEveryoneClaim   # check current value
Set-SPOTenant -ShowEveryoneClaim $true

# PnP PowerShell (cross-platform):
(Get-PnPTenant).ShowEveryoneClaim
Set-PnPTenant -ShowEveryoneClaim $true
```

Then add *Everyone* to the site's Visitors group: **Site Settings → People
and Groups → [Site] Visitors → New → Add Users** → search for *Everyone*
→ **Share**.

> **Pitfall — similar-sounding groups:**
>
> - *Everyone* — includes B2B guests ✓
> - *Everyone except external users* — **excludes** guests ✗

### External sharing

SharePoint's tenant-level sharing setting acts as a **ceiling**: individual
sites cannot be more permissive than the tenant allows.

- **Active sites → [landing page site] → Policies → External sharing** —
  set to at least *Existing guests only*.

If that option is greyed out, raise it under **SharePoint Admin Center →
Policies → Sharing** to at least *Existing guests only*, then configure the
site.

---

## Guest Sponsor API

### Pre-step: create the App Registration

The Azure Function uses
[EasyAuth](https://learn.microsoft.com/azure/app-service/overview-authentication-authorization).
EasyAuth needs an Entra App Registration as its identity provider.

**Option A — run directly from the web** (no clone required,
[PowerShell 7+](https://learn.microsoft.com/powershell/scripting/install/installing-powershell)):

```powershell
& ([scriptblock]::Create((iwr 'https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/setup-app-registration.ps1')))
```

**Option B — from a local clone:**

```powershell
./azure-function/infra/setup-app-registration.ps1
```

<details>
<summary>Option C — download and run manually</summary>

```powershell
Invoke-WebRequest `
  'https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/setup-app-registration.ps1' `
  -OutFile setup-app-registration.ps1

# Review the script content before running it:
Get-Content setup-app-registration.ps1

./setup-app-registration.ps1
```

</details>

<details>
<summary>Option D — manual alternative (Azure Portal)</summary>

1. **Microsoft Entra admin center → App registrations → New registration**.
2. Name: `Guest Sponsor Info - SharePoint Web Part Auth`; Supported
   account types: *Accounts in this organizational directory only*.
3. **Expose an API → Set** Application ID URI:
   `api://guest-sponsor-info-proxy/<clientId>`.
4. Copy the **Client ID** — this is used as `ALLOWED_AUDIENCE`.

</details>

Copy the **Client ID** printed at the end.

### Set up in Azure

Click the button to start:

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2Fworkoho%2Fspfx-guest-sponsor-info%2Fmain%2Fazure-function%2Finfra%2Fazuredeploy.json)

Or from [Azure Cloud Shell](https://shell.azure.com):

```bash
az deployment group create \
  --resource-group <your-resource-group> \
  --template-uri https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/azuredeploy.json \
  --parameters \
      tenantId=<your-tenant-id> \
      tenantName=<your-tenant-name> \
      functionAppName=<globally-unique-name> \
      functionClientId=<client-id-from-pre-step>
```

<details>
<summary>Optional: deploy as a Deployment Stack</summary>

[Deployment Stacks](https://learn.microsoft.com/azure/azure-resource-manager/bicep/deployment-stacks)
track all resources as a managed set. Clean teardown requires a single command.

```bash
az stack group create \
  --name guest-sponsor-info \
  --resource-group <your-resource-group> \
  --template-uri https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/azuredeploy.json \
  --parameters \
      tenantId=<your-tenant-id> \
      tenantName=<your-tenant-name> \
      functionAppName=<globally-unique-name> \
      functionClientId=<client-id-from-pre-step> \
  --action-on-unmanage deleteResources \
  --deny-settings-mode none
```

</details>

### Required parameters

| Parameter | Description |
|---|---|
| `tenantId` | Your Entra tenant ID (GUID) |
| `tenantName` | Tenant name without domain suffix, e.g. `contoso` |
| `functionAppName` | Globally unique name for the Function App |
| `functionClientId` | Client ID from the pre-step |
| `appVersion` | `"latest"` (default) or pinned SemVer without `v` |
| `location` | Azure region |

### Hosting plan options

| | **Consumption** (default) | **Flex Consumption** |
|---|---|---|
| Free tier | 1M exec + 400K GB-s/month | None |
| Cold starts | ~2-5 s after ~20 min idle | Eliminated with `alwaysReadyInstances=1` |
| OS | Windows | Linux only |
| Deploy to Azure button | Supported | Supported |
| Cost guard | `dailyMemoryTimeQuota` | `maximumFlexInstances` |
| Estimated cost | Free (within grant) | ~€2-5/month with 1 warm instance |

Check [aka.ms/flex-region](https://aka.ms/flex-region) for Flex Consumption
regional support. Additional parameters for Flex:

```bash
az deployment group create \
  --resource-group <your-resource-group> \
  --template-uri https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/azuredeploy.json \
  --parameters \
      tenantId=<your-tenant-id> \
      tenantName=<your-tenant-name> \
      functionAppName=<globally-unique-name> \
      functionClientId=<client-id-from-pre-step> \
      hostingPlan=FlexConsumption \
      maximumFlexInstances=10
```

### Setup outputs

After setup, open **Resource Group → Deployments → Outputs**:

| Output | Used for |
|---|---|
| `managedIdentityObjectId` | Required for `setup-graph-permissions.ps1` |
| `functionAppUrl` | Web part property pane → **Guest Sponsor API Base URL** |
| `sponsorApiUrl` | Full endpoint URL (for health checks) |

### Grant Graph permissions

**Option A — run directly from the web:**

```powershell
& ([scriptblock]::Create((iwr 'https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/setup-graph-permissions.ps1')))
```

**Option B — from a local clone:**

```powershell
./azure-function/infra/setup-graph-permissions.ps1
```

This script:

1. **Managed Identity Graph permissions** — assigns
  [`User.Read.All`](https://learn.microsoft.com/en-us/graph/permissions-reference#userreadall),
  [`Presence.Read.All`](https://learn.microsoft.com/en-us/graph/permissions-reference#presencereadall)
  (optional),
  [`MailboxSettings.Read`](https://learn.microsoft.com/en-us/graph/permissions-reference#mailboxsettingsread)
  (optional), and
  [`TeamMember.Read.All`](https://learn.microsoft.com/en-us/graph/permissions-reference#teammemberreadall)
  (optional).
2. **App Registration setup** — exposes a `user_impersonation` scope and
   pre-authorizes *SharePoint Online Web Client Extensibility* so the web
   part can acquire tokens silently.

### Configure the web part

In the property pane (**Guest Sponsor API** group):

- **Guest Sponsor API Base URL** — e.g.
  `https://guest-sponsor-info-xyz.azurewebsites.net`
- **Guest Sponsor API Client ID (App Registration)** — the Client ID from
  the App Registration created in the pre-step

---

## Administration and Operations

For day-2 operations, see the separate
[Operations Guide]({{ '/en/operations/' | relative_url }}):

- [Updating the web part]({{ '/en/operations/' | relative_url }}#updating-the-web-part)
- [Inline address map configuration]({{ '/en/operations/' | relative_url }}#inline-address-map-azure-maps)
- [Updating the function]({{ '/en/operations/' | relative_url }}#updating-the-function)

For security posture and trust assumptions, see the
[security assessment on GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/security-assessment.md).

For telemetry and attribution details, see
[Telemetry]({{ '/en/telemetry/' | relative_url }}).

If something does not work as expected, see the [Support]({{ '/en/support/' | relative_url }}) page.
