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

## Overview

Setting up Guest Sponsor Info involves three phases:

| Phase | Where | Minimum role required |
|---|---|---|
| 1 — SharePoint | SharePoint Admin Center + landing page site | SharePoint Administrator |
| 2 — Guest Sponsor API | Azure Portal / Cloud Shell / PowerShell | Azure Contributor + Entra Admin |
| 3 — Web part | SharePoint landing page (edit mode) | Site Owner |

> **The web part includes a built-in Setup Wizard**
>
> The first time you add the web part to a page, a **Setup Wizard** opens
> automatically. It walks you through choosing between production mode
> (Guest Sponsor API) and demo mode, shows the Azure setup commands
> inline with copy buttons, and lets you enter the API credentials at the
> end. This page is the full reference that the wizard links to — work
> through Phases 1 and 2 before (or alongside) running the wizard, then
> complete Phase 3 inside it.

---

## Phase 1 — SharePoint

### Install from Microsoft AppSource

> **AppSource listing pending review** — The web part has been submitted to the
> Microsoft commercial marketplace and is currently awaiting approval. The
> installation steps below describe the process once the listing is live.

The web part will be available in the
[**Microsoft commercial marketplace (AppSource)**](https://appsource.microsoft.com/).
Installing from there deploys it tenant-wide via the Tenant App Catalog — no
file upload or manual deployment required.

**Install via SharePoint Admin Center:**

1. Open **SharePoint Admin Center → More features → Apps → Open**.
2. Click **Get apps from marketplace** and search for *Guest Sponsor Info*.
3. Select the app and click **Get it now**.

The solution uses `skipFeatureDeployment: false` — the web part does **not**
become available tenant-wide automatically. After the Tenant App Catalog
installation, a Site Collection Administrator must add the app to the landing
page site explicitly: **Site Contents → Add an app → Guest Sponsor Info**.
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

> CDN propagation takes **up to 15 minutes**. Once active, the bundle URL changes
> to `publiccdn.sharepointonline.com` automatically — no reconfiguration needed.

If enabling the Public CDN is not possible in your environment, or if you are
deploying outside of AppSource, see the
[full deployment guide on GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/deployment.md)
for alternative options including direct Tenant App Catalog upload and Site
Collection App Catalog deployments.

### Verify guest access to the landing page site

Guests need at least **Read** (Visitor) permission on the landing page site.
Rather than a dynamic Entra group — which can take up to 24 hours to reflect
new members — use the built-in **Everyone** group. It covers every
authenticated user including B2B guests who have accepted their invitation,
and takes effect immediately.

The *Everyone* group is controlled by the `ShowEveryoneClaim` tenant setting.
Since March 2018, external users no longer receive the Everyone claim by
default — you must explicitly enable the setting. If *Everyone* does not
appear in the People Picker, run:

```powershell
# SharePoint Online Management Shell (Windows):
Set-SPOTenant -ShowEveryoneClaim $true

# PnP PowerShell (cross-platform):
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

## Phase 2 — Guest Sponsor API

The Guest Sponsor API is a companion Azure Function that proxies all Microsoft
Graph calls on behalf of the web part. Guests authenticate against it using
[EasyAuth](https://learn.microsoft.com/azure/app-service/overview-authentication-authorization),
and the function queries Graph using its own Managed Identity — guests never
need directory-level permissions in your tenant.

The Setup Wizard shows these three steps inline with copyable commands. The
sections below are the full reference for each step.

### Step 1: Create the App Registration

EasyAuth requires an Entra App Registration as its identity provider for the
Azure Function.

**Option A — run directly from the web** (no clone required,
[PowerShell 7+](https://learn.microsoft.com/powershell/scripting/install/installing-powershell)):

```powershell
& ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/setup-app-registration.ps1').Content))
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

Copy the **Client ID** printed at the end — you will need it in Step 3 and
when configuring the web part.

### Step 2: Deploy to Azure

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
      webPartClientId=<client-id-from-step-1>
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
      webPartClientId=<client-id-from-step-1> \
  --action-on-unmanage deleteResources \
  --deny-settings-mode none
```

</details>

#### Required parameters

| Parameter | Description |
|---|---|
| `tenantId` | Your Entra tenant ID (GUID) |
| `tenantName` | Tenant name without domain suffix, e.g. `contoso` |
| `functionAppName` | Globally unique name for the Function App |
| `webPartClientId` | Client ID from Step 1 |
| `appVersion` | `"latest"` (default) or pinned SemVer without `v` |
| `location` | Azure region |

#### Hosting plan options

| | **Consumption** (default) | **Flex Consumption** |
|---|---|---|
| Free tier | 1M exec + 400K GB-s/month | 250K exec + 100K GB-s/month (on-demand) |
| Cold starts | ~2-5 s after ~20 min idle | Eliminated with `alwaysReadyInstances=1` |
| OS | Windows | Linux only |
| Deploy to Azure button | Supported | Supported |
| Cost guard | `dailyMemoryTimeQuota` | `maximumFlexInstances` |
| Estimated cost | Free (within grant) | ~€2-5/month with 1 warm instance |

Check the [supported regions list](https://learn.microsoft.com/en-us/azure/azure-functions/flex-consumption-how-to#view-currently-supported-regions)
for Flex Consumption availability. Additional parameters for Flex:

```bash
az deployment group create \
  --resource-group <your-resource-group> \
  --template-uri https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/azuredeploy.json \
  --parameters \
      tenantId=<your-tenant-id> \
      tenantName=<your-tenant-name> \
      functionAppName=<globally-unique-name> \
      webPartClientId=<client-id-from-step-1> \
      hostingPlan=FlexConsumption \
      maximumFlexInstances=10
```

#### Deployment outputs

After deployment, open **Resource Group → Deployments → Outputs**:

| Output | Used for |
|---|---|
| `managedIdentityObjectId` | Required for Step 3 |
| `functionAppUrl` | Web part property pane → **Guest Sponsor API Base URL** |
| `sponsorApiUrl` | Full endpoint URL (for health checks) |

### Step 3: Grant Graph permissions

**Option A — run directly from the web:**

```powershell
& ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/setup-graph-permissions.ps1').Content))
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

---

## Phase 3 — Configure the web part

With Phases 1 and 2 complete, open the SharePoint landing page in edit mode
and add the **Guest Sponsor Info** web part to the page.

The **Setup Wizard** will open automatically (it appears whenever the API
URL has not been configured yet). Select **Guest Sponsor API**, then advance
through the wizard steps to the **Connect** screen and enter:

- **Guest Sponsor API Base URL** — the `functionAppUrl` from the deployment
  outputs (Phase 2, Step 2),
  e.g. `https://guest-sponsor-info-xyz.azurewebsites.net`
- **Guest Sponsor API Client ID** — the Client ID from Phase 2, Step 1

The wizard validates the format of both values before saving. You can also
skip the wizard and configure the web part manually: open the **property pane**
(gear icon in edit mode) and fill in the **Guest Sponsor API** group directly.

> **Guest Accessibility check**
>
> After saving, open the property pane and navigate to the
> **Guest Accessibility** panel. It runs a series of checks (CDN status,
> site permissions, external sharing) and shows the result of each with a
> recommendation. Use this to confirm that the Phase 1 prerequisites are
> working as expected.

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
