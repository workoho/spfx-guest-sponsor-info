---
layout: doc
lang: en
title: Deployment Guide
permalink: /en/setup/
description: >-
  Step-by-step deployment guide for the Guest Sponsor Info web part
  and Guest Sponsor API — SharePoint and Azure setup.
lead: >-
  Initial deployment and first-time configuration reference for
  SharePoint and Azure administrators.
github_doc: deployment.md
---

## SharePoint Deployment

### Enable the Site Collection App Catalog

The web part's bundle is hosted in a
[**Site Collection App Catalog**](https://learn.microsoft.com/sharepoint/dev/general-development/site-collection-app-catalog)
directly on the guest landing page site. Because guest users already need read
access to that site, no CDN configuration or additional permissions on the
global App Catalog are required.

Enable the Site Collection App Catalog once. There is no GUI option for this
step — PowerShell is required. The executing account must satisfy **all three**
conditions below; if any is missing the command may appear to succeed but the
App Catalog will be mis-provisioned and deployments will silently fail.

**Required conditions:**

1. [**SharePoint Administrator**](https://learn.microsoft.com/sharepoint/sharepoint-admin-role)
   role in Microsoft 365. A Global Administrator satisfies this implicitly.
2. **Site Collection Administrator on the tenant-level App Catalog** site
   (typically `https://<tenant>.sharepoint.com/sites/appcatalog`).
   The SharePoint Administrator role does *not* grant this automatically.
   If needed, add your account first using
   [`Set-SPOUser`](https://learn.microsoft.com/powershell/module/sharepoint-online/set-spouser):

   ```powershell
   Set-SPOUser -Site "https://<tenant>.sharepoint.com/sites/appcatalog" `
       -LoginName "<admin@tenant.onmicrosoft.com>" `
       -IsSiteCollectionAdmin $true
   ```

3. **Site Collection Administrator on the landing-page site** itself.

> **Prerequisite — tenant App Catalog must exist first:**
> A tenant App Catalog is *not* provisioned automatically on a fresh Microsoft
> 365 tenant. If it has not been created, open
> **SharePoint Admin Center → More features → Apps → Open** — this triggers
> automatic creation
> ([Manage apps using the Apps site](https://learn.microsoft.com/sharepoint/use-app-catalog)).
> Without it,
> [`Add-SPOSiteCollectionAppCatalog`](https://learn.microsoft.com/powershell/module/sharepoint-online/add-spositecollectionappcatalog)
> fails with a cryptic null-reference error.

On Windows, the [**SharePoint Online Management Shell**](https://learn.microsoft.com/powershell/sharepoint/sharepoint-online/connect-sharepoint-online)
is the simplest option:

```powershell
# Install once:
# Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser

Connect-SPOService -Url "https://<tenant>-admin.sharepoint.com"
Add-SPOSiteCollectionAppCatalog -Site "https://<tenant>.sharepoint.com/sites/<landing-site>"
```

On macOS or Linux, use
[PnP PowerShell](https://pnp.github.io/powershell/). Current versions require
you to register your own Entra app and pass its Client ID:

```powershell
# Install once (PowerShell 7+):
# Install-Module PnP.PowerShell -Scope CurrentUser

Connect-PnPOnline -Url "https://<tenant>-admin.sharepoint.com" `
    -ClientId "<your-pnp-app-client-id>" -Interactive
Add-PnPSiteCollectionAppCatalog -Site "https://<tenant>.sharepoint.com/sites/<landing-site>"
```

### Upload and install

1. Download the latest `guest-sponsor-info.sppkg` from
   [Releases](https://github.com/workoho/spfx-guest-sponsor-info/releases).
2. Open the Site Collection App Catalog and upload the `.sppkg` file.

   > **Navigation tip — use the direct URL:**
   > The Site Collection App Catalog is a document library called
   > **Apps for SharePoint** inside the landing-page site. The most
   > reliable way to get there is the direct URL:
   > `https://<tenant>.sharepoint.com/sites/<landing-site>/AppCatalog/`

3. The web part becomes available on all pages within this site collection
   immediately — no additional "Add App" step is required.

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

The [architecture diagram]({{ '/en/architecture/' | relative_url }})
gives a visual overview of all admin roles and deployment steps involved.

### Pre-step: create the App Registration

The Azure Function uses
[EasyAuth](https://learn.microsoft.com/azure/app-service/overview-authentication-authorization).
EasyAuth needs an Entra App Registration as its identity provider.

**Option A — run directly from the web** (no clone required,
[PowerShell 7+](https://learn.microsoft.com/powershell/scripting/install/installing-powershell)):

```powershell
& ([scriptblock]::Create((iwr 'https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/setup-app-registration.ps1'))) -TenantId "<your-tenant-id>"
```

**Option B — from a local clone:**

```powershell
./azure-function/infra/setup-app-registration.ps1 -TenantId "<your-tenant-id>"
```

<details>
<summary>Option C — download and run manually</summary>

```powershell
Invoke-WebRequest `
  'https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/setup-app-registration.ps1' `
  -OutFile setup-app-registration.ps1

# Review the script content before running it:
Get-Content setup-app-registration.ps1

./setup-app-registration.ps1 -TenantId "<your-tenant-id>"
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

### Deploy to Azure

Click the button to start the deployment:

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fgithub.com%2Fworkoho%2Fspfx-guest-sponsor-info%2Freleases%2Flatest%2Fdownload%2Fazuredeploy.json)

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
| Cold starts | ~2–5 s after ~20 min idle | Eliminated with `alwaysReadyInstances=1` |
| OS | Windows | Linux only |
| Deploy to Azure button | Supported | Supported |
| Cost guard | `dailyMemoryTimeQuota` | `maximumFlexInstances` |
| Estimated cost | Free (within grant) | ~€2–5/month with 1 warm instance |

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

### Deployment outputs

After deployment, open **Resource Group → Deployments → Outputs**:

| Output | Used for |
|---|---|
| `managedIdentityObjectId` | Required for `setup-graph-permissions.ps1` |
| `functionAppUrl` | Web part property pane → **Guest Sponsor API Base URL** |
| `sponsorApiUrl` | Full endpoint URL (for health checks) |

### Grant Graph permissions

**Option A — run directly from the web:**

```powershell
& ([scriptblock]::Create((iwr 'https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/setup-graph-permissions.ps1'))) `
  -ManagedIdentityObjectId "<oid-from-deployment-output>" `
  -TenantId "<your-tenant-id>" `
  -FunctionAppClientId "<client-id-from-pre-step>"
```

**Option B — from a local clone:**

```powershell
./azure-function/infra/setup-graph-permissions.ps1 `
  -ManagedIdentityObjectId "<oid-from-deployment-output>" `
  -TenantId "<your-tenant-id>" `
  -FunctionAppClientId "<client-id-from-pre-step>"
```

This script:

1. **Managed Identity Graph permissions** — assigns `User.Read.All`,
   `Presence.Read.All` (optional), and `MailboxSettings.Read` (optional).
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
