# Deployment Guide

Initial deployment and first-time configuration reference for
SharePoint and Azure administrators.

For a quick-start overview, see the [README](../README.md).
For architecture decisions and internals, see [architecture.md](architecture.md).
For a visual system overview (including a setup checklist diagram), see
[architecture-diagram.md](architecture-diagram.md#setup--two-admin-roles-recommended-path).

For day-2 operations and administration tasks (updates, map configuration,
and runtime maintenance), see [operations.md](operations.md).
For security and telemetry details, see
[security-assessment.md](security-assessment.md) and [telemetry.md](telemetry.md).

## Table of Contents

- [SharePoint Deployment](#sharepoint-deployment)
  - [Install the Webpart](#step-1--Installation)
  - [Make the web part accessible to guest users](#step-2--make-the-web-part-accessible-to-guest-users)
  - [Verify guest access to the landing page site](#verify-guest-access-to-the-landing-page-site)
  - [External sharing](#external-sharing)
- [Guest Sponsor API](#guest-sponsor-api)
- [Administration and Operations](#administration-and-operations)

---

## SharePoint Deployment

### Step 1 - Installation

#### Option A - Install from Microsoft AppSource

The web part is available in the
[**Microsoft commercial marketplace (AppSource)**](https://appsource.microsoft.com/).
Installing from there deploys it tenant-wide via the Tenant App Catalog — no
Site Collection App Catalog or file upload required.

**Install via SharePoint Admin Center:**

1. Open [**SharePoint Admin Center → More features → Apps → Open**](https://admin.microsoft.com/sharepoint).
2. Click **Get apps from marketplace** and search for *Guest Sponsor Info*.
3. Select the app and click **Get it now**.

The solution uses `skipFeatureDeployment: false`. The web part is **not**
automatically available on all pages after installation. A Site Collection
Administrator must first add the app to the desired site:
**Site Contents → Add an app → Guest Sponsor Info**.
This is intentional — it prevents the web part from being accidentally placed
on unintended sites.

The web part requests no Microsoft Graph permissions of its own — the
**SharePoint Admin Center → Advanced → API access** queue will remain empty.
All Graph data is fetched server-side by the companion Azure Function using
its Managed Identity.

---
#### Option B — Use a Tenant App Catalog
**This option does not use the Microsoft commercial marketplace.** Download the
`.sppkg` from
[GitHub Releases](https://github.com/workoho/spfx-guest-sponsor-info/releases)
instead.

> A tenant App Catalog is *not* provisioned automatically on a fresh Microsoft
> 365 tenant. If it has not been created, open
> **SharePoint Admin Center → More features → Apps → Open** — this triggers
> automatic creation
> ([Manage apps using the Apps site](https://learn.microsoft.com/sharepoint/use-app-catalog)).

**Upload and install:**

1. Download the latest `guest-sponsor-info.sppkg` from
   [GitHub Releases](https://github.com/workoho/spfx-guest-sponsor-info/releases).
2. Open the Tenant App Catalog, upload the `.sppkg` file, and click
   **Deploy** in the dialog that appears.
3. Go to your landing site and open **Site Contents → Add an app**, find **Guest Sponsor Info** in the
   list, and click it to install the app on the site. The web part then
   appears in the page editor.

   Because this solution uses `skipFeatureDeployment: false`, the app must be
   explicitly added to each site where the web part is needed — even when
   using a Site Collection App Catalog.

   > **Updating to a new version:** Re-upload the `.sppkg` to
   > **Apps for SharePoint** and click **Deploy**. SharePoint will then show
   > an **Update** banner on the installed app in **Site Contents** — click it
   > to apply the update. If the banner does not appear, remove the app
   > (**Site Contents** → hover the app → **Remove**) and add it again via
   > **Add an app**.

---

#### Option C — Use a Site Collection App Catalog

As an alternative to the AppSource / Tenant App Catalog path, deploy the web
part bundle in a Site Collection App Catalog directly on the guest landing page
site. This bypasses the Tenant App Catalog entirely — the bundle is served from
a library the guest already has access to. No CDN configuration or separate
Everyone grant on the App Catalog is needed.

**This option does not use the Microsoft commercial marketplace.** Download the
`.sppkg` from
[GitHub Releases](https://github.com/workoho/spfx-guest-sponsor-info/releases)
instead.

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
   # SharePoint Online Management Shell:
   Set-SPOUser -Site "https://<tenant>.sharepoint.com/sites/appcatalog" `
       -LoginName "<admin@tenant.onmicrosoft.com>" `
       -IsSiteCollectionAdmin $true
   ```

3. **Site Collection Administrator on the landing-page site** itself.

> **Prerequisite — tenant App Catalog must exist first:**\
> A tenant App Catalog is *not* provisioned automatically on a fresh Microsoft
> 365 tenant. If it has not been created, open
> **SharePoint Admin Center → More features → Apps → Open** — this triggers
> automatic creation
> ([Manage apps using the Apps site](https://learn.microsoft.com/sharepoint/use-app-catalog)).
> Without it,
> [`Add-SPOSiteCollectionAppCatalog`](https://learn.microsoft.com/powershell/module/sharepoint-online/add-spositecollectionappcatalog)
> fails with a cryptic null-reference error.

On Windows, the [**SharePoint Online Management Shell**](https://learn.microsoft.com/powershell/sharepoint/sharepoint-online/connect-sharepoint-online)
is the simplest option — it works with your existing credentials and requires
no additional setup:

```powershell
# Install once:
# Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser

Connect-SPOService -Url "https://<tenant>-admin.sharepoint.com"
Add-SPOSiteCollectionAppCatalog -Site "https://<tenant>.sharepoint.com/sites/<landing-site>"
```

*Cmdlet references:
[`Connect-SPOService`](https://learn.microsoft.com/powershell/module/sharepoint-online/connect-sposervice) ·
[`Add-SPOSiteCollectionAppCatalog`](https://learn.microsoft.com/powershell/module/sharepoint-online/add-spositecollectionappcatalog)*

On macOS or Linux (or if you prefer a cross-platform tool), use
[PnP PowerShell](https://pnp.github.io/powershell/). Note that current
versions of PnP PowerShell require you to register your own Entra app and
pass its Client ID — the built-in interactive login was removed:

```powershell
# Install once (PowerShell 7+):
# Install-Module PnP.PowerShell -Scope CurrentUser

Connect-PnPOnline -Url "https://<tenant>-admin.sharepoint.com" `
    -ClientId "<your-pnp-app-client-id>" -Interactive
Add-PnPSiteCollectionAppCatalog -Site "https://<tenant>.sharepoint.com/sites/<landing-site>"
```

*Cmdlet references:
[`Connect-PnPOnline`](https://pnp.github.io/powershell/cmdlets/Connect-PnPOnline.html) ·
[`Add-PnPSiteCollectionAppCatalog`](https://pnp.github.io/powershell/cmdlets/Add-PnPSiteCollectionAppCatalog.html) ·
[register an Entra app](https://pnp.github.io/powershell/articles/registerapplication.html)*

**Upload and install:**

1. Download the latest `guest-sponsor-info.sppkg` from
   [GitHub Releases](https://github.com/workoho/spfx-guest-sponsor-info/releases).
2. Open the Site Collection App Catalog, upload the `.sppkg` file, and click
   **Deploy** in the dialog that appears.

   > **Navigation tip — use the direct URL:**\
   > The Site Collection App Catalog is a document library called
   > **Apps for SharePoint** inside the landing-page site, but there is no
   > obvious link to it in the SharePoint UI. The most reliable way to get
   > there is the direct URL:
   > `https://<tenant>.sharepoint.com/sites/<landing-site>/AppCatalog/`\
   > Alternatively: **Site Contents** (gear icon → Site contents) → scroll
   > down to find **Apps for SharePoint** → open it.

3. Go to **Site Contents → Add an app**, find **Guest Sponsor Info** in the
   list, and click it to install the app on the site. The web part then
   appears in the page editor.

   Because this solution uses `skipFeatureDeployment: false`, the app must be
   explicitly added to each site where the web part is needed — even when
   using a Site Collection App Catalog.

   > **Updating to a new version:** Re-upload the `.sppkg` to
   > **Apps for SharePoint** and click **Deploy**. SharePoint will then show
   > an **Update** banner on the installed app in **Site Contents** — click it
   > to apply the update. If the banner does not appear, remove the app
   > (**Site Contents** → hover the app → **Remove**) and add it again via
   > **Add an app**.

---

### Step 2 - Make the web part accessible to guest users

When installed via AppSource or the Tenant App Catalog, the web part JavaScript
bundle is served from the Tenant App Catalog's `ClientSideAssets` library by
default. B2B guest users cannot access this library before authenticating to
the host tenant — which is not guaranteed before the page load. If guests
cannot load the bundle, the web part silently fails to render.

The web part's built-in **Guest Accessibility** diagnostics panel (property
pane) detects the current scenario and shows the result of each check with
a recommendation.

Choose **one** of the following options:

#### Option A — Enable the Office 365 Public CDN (recommended)

When the Office 365 Public CDN is enabled, SharePoint automatically replicates
web part bundles to Microsoft's edge CDN network
(`publiccdn.sharepointonline.com`), which is accessible anonymously without
any SharePoint authentication. This is the most reliable approach for guest
users and requires no per-site configuration.

**Required role:** SharePoint Administrator.

```powershell
# SharePoint Online Management Shell (Windows):
# Install once: Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
Connect-SPOService -Url "https://<tenant>-admin.sharepoint.com"

# Enable the Public CDN:
Set-SPOTenantCdnEnabled -CdnType Public -Enable $true

# Verify the ClientSideAssets origin is included (added by default):
Get-SPOTenantCdnOrigins -CdnType Public
# Expected output includes: */CLIENTSIDEASSETS
```

If `*/CLIENTSIDEASSETS` is not listed, add it:

```powershell
Add-SPOTenantCdnOrigin -CdnType Public -OriginUrl "*/CLIENTSIDEASSETS"
```

```powershell
# PnP PowerShell (cross-platform):
# Install once (PowerShell 7+): Install-Module PnP.PowerShell -Scope CurrentUser
Connect-PnPOnline -Url "https://<tenant>-admin.sharepoint.com" `
    -ClientId "<your-pnp-app-client-id>" -Interactive
Set-PnPTenantCdnEnabled -CdnType Public -Enable $true
```

*Cmdlet references:
[`Set-SPOTenantCdnEnabled`](https://learn.microsoft.com/powershell/module/sharepoint-online/set-spotenantcdnenabled) ·
[`Get-SPOTenantCdnOrigins`](https://learn.microsoft.com/powershell/module/sharepoint-online/get-spotenantcdnorigins) ·
[`Add-SPOTenantCdnOrigin`](https://learn.microsoft.com/powershell/module/sharepoint-online/add-spotenantcdnorigin) ·
[Office 365 CDN documentation](https://learn.microsoft.com/sharepoint/dev/general-development/office-365-cdn-with-spo-ps)*

> CDN propagation typically takes **15-30 minutes** after enabling. Once active,
> the bundle URL changes from the Tenant App Catalog to
> `publiccdn.sharepointonline.com` automatically — no reconfiguration is needed.

#### Option B — Grant Everyone read access to the Tenant App Catalog

If enabling the Public CDN is not possible in your environment, grant the
built-in **Everyone** group (`c:0(.s|true`) read access to the Tenant App
Catalog site. This allows B2B guests who have already accepted their invitation
to load the bundle directly from the App Catalog.

**Required roles:** SharePoint Administrator and Site Collection Administrator
on the Tenant App Catalog site (typically
`https://<tenant>.sharepoint.com/sites/appcatalog`).

```powershell
# SharePoint Online Management Shell (Windows):
Connect-SPOService -Url "https://<tenant>-admin.sharepoint.com"
Add-SPOUser -Site "https://<tenant>.sharepoint.com/sites/appcatalog" `
    -LoginName "c:0(.s|true" -Group "App Catalog Visitors"
```

```powershell
# PnP PowerShell (cross-platform; connect to the App Catalog site directly):
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com/sites/appcatalog" `
    -ClientId "<your-pnp-app-client-id>" -Interactive
Add-PnPGroupMember -LoginName "c:0(.s|true" -Group "App Catalog Visitors"
```

*Cmdlet references:
[`Add-SPOUser`](https://learn.microsoft.com/powershell/module/sharepoint-online/add-spouser) ·
[`Add-PnPGroupMember`](https://pnp.github.io/powershell/cmdlets/Add-PnPGroupMember.html)*

> **Limitation:** This covers only guests who have already authenticated to the
> host tenant. Guests who have never visited the tenant (e.g. before accepting a
> full invitation) cannot load the bundle. The Public CDN (Option A) does not
> have this limitation.


### Step 3 - Verify guest access to the landing page site

If your landing page site is already in use, Visitor access for guests is most
likely already configured — but it's worth checking that the approach works
reliably for newly invited users. If you're setting up a fresh landing page,
use a **Communication Site** (not a Team Site): it has a clean Visitor
permission model with no attached Microsoft 365 group, which simplifies guest
access management.

Guests need at least **Read** (Visitor) permission on the landing page site
to view it. Rather than a dynamic Entra group — which can take up to 24 hours
to reflect new members, meaning freshly invited guests may see an access-denied
page until the next sync — use the built-in **Everyone** group. It covers every
authenticated user including B2B guests who have accepted their invitation, and
takes effect immediately.

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

*Cmdlet references:
[`Get-SPOTenant`](https://learn.microsoft.com/powershell/module/sharepoint-online/get-spotenant) ·
[`Set-SPOTenant`](https://learn.microsoft.com/powershell/module/sharepoint-online/set-spotenant) ·
[`Get-PnPTenant`](https://pnp.github.io/powershell/cmdlets/Get-PnPTenant.html) ·
[`Set-PnPTenant`](https://pnp.github.io/powershell/cmdlets/Set-PnPTenant.html)*

Then add *Everyone* to the site's Visitors group. The easiest path is the GUI:
**Site Settings → People and Groups → [Site] Visitors → New → Add Users**
→ search for *Everyone* → **Share**. Alternatively via PowerShell:

```powershell
# SharePoint Online Management Shell (Windows):
# Connect-SPOService must already be active (see Enable step above)
Add-SPOUser -Site "https://<tenant>.sharepoint.com/sites/<landing-site>" `
    -LoginName "c:0(.s|true" -Group "<SiteName> Visitors"

# PnP PowerShell (cross-platform; connect to the site, not the admin URL):
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com/sites/<landing-site>" `
    -ClientId "<your-pnp-app-client-id>" -Interactive
Add-PnPGroupMember -LoginName "c:0(.s|true" -Group "<SiteName> Visitors"
```

*Cmdlet references:
[`Add-SPOUser`](https://learn.microsoft.com/powershell/module/sharepoint-online/add-spouser) ·
[`Add-PnPGroupMember`](https://pnp.github.io/powershell/cmdlets/Add-PnPGroupMember.html)*

> **Pitfall — similar-sounding groups:**
>
> - *Everyone* — includes B2B guests ✓
> - *Everyone except external users* — **excludes** guests ✗

#### Alternative: static Entra security group

If your organisation uses an **automated guest invitation workflow** — rather than
implicit invitations triggered by sharing content in Teams or SharePoint — a
dedicated static Entra security group populated at invitation time is a viable
alternative. The group membership is immediately correct in Entra ID, so SharePoint
Online can resolve it on the next access check, typically within seconds to a few
minutes. This contrasts with dynamic Entra groups, where Entra itself must first
re-evaluate the dynamic membership rule — a process that can take up to 24 hours —
before SharePoint can reflect the change.

The *Everyone* group remains the **recommended approach**, however, because the
`c:0(.s|true` claim is evaluated entirely within SharePoint Online's own
authentication layer: SharePoint simply checks whether the incoming request carries
a valid, authenticated identity — with no need to query Entra ID for group
membership at all. A static group still requires SharePoint to resolve the user's
current Entra group memberships before granting access, which introduces a
dependency on that resolution path even though it is normally fast.

> **Tip:** [EasyLife 365 Collaboration](https://easylife365.cloud/products/collaboration/) is purpose-built
> for automated Microsoft 365 guest lifecycle management and can ensure that a static
> site-access group is populated for every guest invitation — including guests who
> would otherwise be invited implicitly through Teams or SharePoint.
> [Workoho](https://workoho.com/?utm_source=gsiw&utm_medium=docs&utm_campaign=repo&utm_content=deployment),
> the author of this web part, is a Platinum
> implementation partner of EasyLife 365.

### Step 4 - Ensure External sharing is enabled

SharePoint's tenant-level sharing setting acts as a **ceiling**: individual
sites cannot be more permissive than the tenant allows, but they can be
more restrictive. What matters here is the setting on the landing page site
itself:

- **Active sites → [landing page site] → Policies → External sharing** —
  set to at least *Existing guests only*.

If that option is greyed out or missing, the tenant-level ceiling is too
restrictive. Raise it under **SharePoint Admin Center → Policies → Sharing**
to at least *Existing guests only*, then configure the site.

For background on how SharePoint sharing levels interact, see
[External sharing overview](https://learn.microsoft.com/sharepoint/external-sharing-overview)
in the Microsoft documentation.

---

## Guest Sponsor API

> The [Setup diagram](architecture-diagram.md#setup--two-admin-roles-recommended-path)
> gives a visual overview of all admin roles and deployment steps involved.

### Pre-step: create the App Registration

The Azure Function uses [EasyAuth](https://learn.microsoft.com/azure/app-service/overview-authentication-authorization)
(Azure App Service Authentication). EasyAuth needs an Entra App Registration
as its identity provider.

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
# Download the script first, then review and execute it.
Invoke-WebRequest `
  'https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/setup-app-registration.ps1' `
  -OutFile setup-app-registration.ps1

# Review the script content before running it:
Get-Content setup-app-registration.ps1

# Run it (prompts interactively for any required values):
./setup-app-registration.ps1
```

</details>

Copy the **Client ID** printed at the end.

<details>
<summary>Option D — manual alternative (Azure Portal)</summary>

1. **Microsoft Entra admin center → App registrations → New registration**.
2. Name: `Guest Sponsor Info - SharePoint Web Part Auth`; Supported
   account types: *Accounts in this organizational directory only*.
3. **Expose an API → Set** Application ID URI:
   `api://guest-sponsor-info-proxy/<clientId>`.
4. Copy the **Client ID** — this is used as `ALLOWED_AUDIENCE`.

</details>

### Step 1 - Deploy to Azure

Click the button to start the deployment:

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2Fworkoho%2Fspfx-guest-sponsor-info%2Fmain%2Fazure-function%2Finfra%2Fazuredeploy.json)

Or from [Azure Cloud Shell](https://shell.azure.com) (no local tooling
required; also works for updates — ARM deployments are idempotent):

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
<summary>Optional: deploy as a Deployment Stack (recommended for long-term
CLI management)</summary>

[Azure Deployment Stacks](https://learn.microsoft.com/azure/azure-resource-manager/bicep/deployment-stacks)
track all resources as a managed set. Removed resources are automatically
deleted; clean teardown requires a single command.

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

Use the same `az stack group create` command for updates and initial
deployment.

</details>

#### Required parameters

| Parameter | Description |
|---|---|
| `tenantId` | Your Entra tenant ID (GUID) |
| `tenantName` | Tenant name without domain suffix, e.g. `contoso` |
| `functionAppName` | Globally unique name for the Function App |
| `functionClientId` | Client ID from the pre-step |
| `appVersion` | `"latest"` (default) or pinned SemVer without `v`, e.g. `"1.4.2"` |
| `location` | Azure region |

#### Optional hosting plan parameters

| Parameter | Description |
|---|---|
| `hostingPlan` | `Consumption` (default) or `FlexConsumption`. See below. |
| `alwaysReadyInstances` | Pre-warmed instances (Flex only). `1` eliminates cold starts. Default: `1`. |
| `maximumFlexInstances` | **Required for Flex.** Hard upper bound on scale-out (cost ceiling). 1-1000. |
| `instanceMemoryMB` | `512` or `2048` (Flex only). Default: `2048`. |
| `dailyMemoryTimeQuotaGBs` | Daily GB-s budget (Consumption only). Default: `10000`. |

#### Flex Consumption plan

Check [aka.ms/flex-region](https://aka.ms/flex-region) first to confirm
regional support. Add the extra parameters:

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

The initial ZIP upload is automated via a short-lived Azure CLI container
(~2-5 min). The provisioning script resource is retained for 2 hours for
troubleshooting, then removed automatically.

| | **Consumption** (default) | **Flex Consumption** |
|---|---|---|
| Free tier | 1M exec + 400K GB-s/month | None |
| Cold starts | ~2-5 s after ~20 min idle | Eliminated with `alwaysReadyInstances=1` |
| OS | Windows | Linux only |
| Deploy to Azure button | Supported | Supported |
| Cost guard | `dailyMemoryTimeQuota` | `maximumFlexInstances` |
| Estimated cost | Free (within grant) | ~€2-5/month with 1 warm instance |

### Step 3 - Note Deployment outputs

After deployment, open **Resource Group → Deployments → select deployment template →
Outputs**:

| Output | Used for |
|---|---|
| `managedIdentityObjectId` | Required for `setup-graph-permissions.ps1` (next step) |
| `functionAppUrl` | Web part property pane → **Guest Sponsor API Base URL** |
| `sponsorApiUrl` | Full endpoint URL (for curl/Postman health checks) |

### Step 4 - Grant Graph permissions and configure the App Registration

**Option A — run directly from the web** (no clone required,
[PowerShell 7+](https://learn.microsoft.com/powershell/scripting/install/installing-powershell)):

```powershell
& ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/setup-graph-permissions.ps1').Content))
```

**Option B — from a local clone:**

```powershell
./azure-function/infra/setup-graph-permissions.ps1
```

This script:

1. **Managed Identity Graph permissions** — assigns `User.Read.All`,
   `Presence.Read.All` (optional; requires Microsoft Teams), and
   `MailboxSettings.Read` (optional; filters shared/room/equipment mailboxes).
2. **App Registration setup for silent token acquisition** — exposes a
   `user_impersonation` scope and
   [pre-authorizes](https://learn.microsoft.com/entra/identity-platform/permissions-consent-overview#preauthorization)
   *SharePoint Online Web Client Extensibility* so the web part can acquire
   tokens silently without consent prompts or page reloads. See
   [Silent Token Acquisition and Pre-Authorization](architecture.md#silent-token-acquisition-and-pre-authorization)
   in the architecture guide for the full explanation.

### Step 5 - Configure the web part

In the property pane (**Guest Sponsor API** group):

- **Guest Sponsor API Base URL** — e.g.
  `https://guest-sponsor-info-xyz.azurewebsites.net`
- **Guest Sponsor API Client ID (App Registration)** — the Client ID from the
  App Registration named **"Guest Sponsor Info - SharePoint Web Part Auth"**
  in your Entra tenant (created in the pre-step)

---

## Administration and Operations

Deployment and day-2 operations are split into separate guides:

- [operations.md](operations.md) for ongoing administration, including:
  web part updates, inline map configuration, and function update playbooks.
- [security-assessment.md](security-assessment.md) for security posture,
  threat boundaries, and deployment trust assumptions.
- [telemetry.md](telemetry.md) for Customer Usage Attribution, opt-out,
  and verification steps.
