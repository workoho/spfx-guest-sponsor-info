# Deployment Guide

Initial deployment and first-time configuration reference for
SharePoint and Azure administrators.

For a quick-start overview, see the [README](../README.md).
For architecture decisions and internals, see [architecture.md](architecture.md).
For a visual system overview (including a setup checklist diagram), see
[architecture-diagram.md](architecture-diagram.md#setup--admin-roles-and-automation-recommended-path).

For day-2 operations and administration tasks (updates, map configuration,
and runtime maintenance), see [operations.md](operations.md).
For security and telemetry details, see
[security-assessment.md](security-assessment.md) and [telemetry.md](telemetry.md).

## Deployment Phases

This guide keeps the detailed instructions grouped by system, but the overall
operator flow still has three phases:

| Phase | What you finish | Typical role |
|---|---|---|
| 1 - SharePoint | Install the web part package, make the bundle reachable for guests, and confirm landing-page access | SharePoint Administrator |
| 2 - Guest Sponsor API | Deploy the Azure Function, App Registration, and Graph permissions | Azure `Contributor` + `Owner` (or `User Access Administrator`) + required Entra roles |
| 3 - Web part configuration | Paste the API URL and Client ID into the landing-page web part and verify the live connection | Site Owner |

If you only need one specific area, jump to the matching section below. If you
are doing a first-time rollout end to end, complete the phases in that order.

## Table of Contents

- [SharePoint Deployment](#sharepoint-deployment)
  - [Step 1 - Install the web part](#step-1---install-the-web-part)
  - [Step 2 - Make the web part accessible to guest users](#step-2---make-the-web-part-accessible-to-guest-users)
  - [Step 3 - Verify guest access to the landing page site](#step-3---verify-guest-access-to-the-landing-page-site)
  - [Step 4 - Ensure external sharing is enabled](#step-4---ensure-external-sharing-is-enabled)
- [Guest Sponsor API](#guest-sponsor-api)
  - [Step 1 - Confirm prerequisites and required roles](#step-1---confirm-prerequisites-and-required-roles)
  - [Step 2 - Run the deployment wizard](#step-2---run-the-deployment-wizard)
  - [Step 2 (Alternative) - Run from a local infra ZIP](#step-2-alternative---run-from-a-local-infra-zip)
  - [Step 3 (Optional) - Choose hosting plan parameters](#step-3-optional---choose-hosting-plan-parameters)
  - [Step 4 - Record the deployment outputs](#step-4---record-the-deployment-outputs)
  - [Step 5 - Configure the web part](#step-5---configure-the-web-part)
  - [Advanced scenario - Deploy from a Privileged Access Workstation (PAW)](#advanced-scenario---deploy-from-a-privileged-access-workstation-paw)
- [Administration and Operations](#administration-and-operations)

---

## SharePoint Deployment

### Step 1 - Install the web part

#### Option A - Install from Microsoft AppSource

The web part is available in the
[**Microsoft commercial marketplace (AppSource)**](https://appsource.microsoft.com/).
Installing from there places the package in the Tenant App Catalog — no
manual file upload or Site Collection App Catalog setup required.

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

Use this option when you want a tenant-wide package source without relying on
AppSource.

<details>
<summary>Expand Tenant App Catalog instructions</summary>

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

</details>

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

Use this option when the landing-page site must stay isolated from the tenant
App Catalog and you are prepared to manage a site-scoped deployment path.

<details>
<summary>Expand Site Collection App Catalog instructions</summary>

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

</details>

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
[Microsoft 365 CDN documentation](https://learn.microsoft.com/sharepoint/dev/spfx/enable-microsoft-365-content-delivery-network)*

> CDN propagation typically takes **15-30 minutes** after enabling. Once active,
> the bundle URL changes from the Tenant App Catalog to
> `publiccdn.sharepointonline.com` automatically — no reconfiguration is needed.

#### Option B — Grant Everyone read access to the Tenant App Catalog

If enabling the Public CDN is not possible in your environment, grant the
built-in **Everyone** group (`c:0(.s|true`) read access to the Tenant App
Catalog site. This allows B2B guests who have already accepted their invitation
to load the bundle directly from the App Catalog.

Use this fallback only when the Public CDN cannot be enabled in your tenant.

<details>
<summary>Expand Tenant App Catalog permission instructions</summary>

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

</details>

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
authenticated user, and can also cover B2B guests once the tenant is configured
to grant the **Everyone** claim to external users.

Do not assume that external users already receive the **Everyone** claim.
Before relying on this pattern for B2B guests, explicitly set
`ShowEveryoneClaim` to `$true`:

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

This setting controls whether external users receive the **Everyone** claim.
It does not grant site access by itself.

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

Use this alternative only when your guest invitation flow already populates a
dedicated access group predictably.

<details>
<summary>Expand static Entra security group guidance</summary>

The *Everyone* group remains the default recommendation because the
`c:0(.s|true` claim is evaluated directly by SharePoint Online's own
authentication layer. A static Entra group still requires SharePoint to
resolve the user's current Entra group memberships before access is granted,
which adds a dependency on that resolution path even when it is usually fast.

</details>

### Step 4 - Ensure external sharing is enabled

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

> The [Setup diagram](architecture-diagram.md#setup--admin-roles-and-automation-recommended-path)
> gives a visual overview of all admin roles and deployment steps involved.

This section is ordered by execution sequence. Items marked **Alternative** or
**Optional** are only needed when your environment or operating model requires them.

### Step 1 - Confirm prerequisites and required roles

**Required tooling:**

- [PowerShell 7+](https://learn.microsoft.com/powershell/scripting/install/installing-powershell)
  when using the PowerShell command directly.
- On macOS or Linux, the `install.sh` bootstrapper can install PowerShell
  first and then hand off to the same PowerShell installer.

**Required roles:**

| Scope | Required role |
|---|---|
| Resource group | **Contributor** |
| Resource group | **Owner** (or User Access Administrator) — for Managed Identity role assignments |
| Entra ID | **Cloud Application Administrator** — to create and configure the App Registration |
| Entra ID | **Privileged Role Administrator** — to assign Graph app roles to the Managed Identity |

> **Global Administrator** replaces both Entra requirements with a single role.
> The Azure **Contributor** and **Owner** roles on the resource group are still
> required separately and cannot be substituted by any Entra directory role.

If your organisation uses
[PIM](https://learn.microsoft.com/entra/id-governance/privileged-identity-management/pim-configure),
activate the Entra roles before running the script. The pre-provision hook
checks your active directory roles and warns if any are missing.

### Step 2 - Run the deployment wizard

The `install.ps1` script is the recommended entry point. It downloads the
latest infra bundle from GitHub Releases, runs the interactive deployment
wizard (`deploy-azure.ps1`), and cleans up afterwards. No local repository
clone is required.

Run this command in PowerShell 7+:

```powershell
& ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.ps1').Content))
```

On macOS or Linux, you can start from a plain shell instead. The shell
bootstrapper installs PowerShell when needed, downloads `install.ps1`, and then
hands off to the same deployment wizard:

```bash
curl -fsSL https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.sh | bash
```

To pass PowerShell installer parameters through the shell bootstrapper, use
`bash -s --`:

```bash
curl -fsSL https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.sh | bash -s -- -Version v1.2.0
```

If your Azure account is a guest in other tenants and `az login` reports
Conditional Access warnings for those tenants, pass the target Azure/Entra
tenant ID explicitly:

```powershell
& ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.ps1').Content)) -AzureTenantId <tenant-id>
```

```bash
curl -fsSL https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.sh | bash -s -- -AzureTenantId <tenant-id>
```

The selected Azure subscription must belong to the same Entra tenant as the
SharePoint tenant where the web part is installed. The wizard validates the
selected Azure tenant against the SharePoint tenant name and requires explicit
confirmation if they do not appear to match, for example after a SharePoint
domain rename.

To check local tools, sign-in, subscription selection, and visible Azure/Entra
roles without deploying resources, add `-PreflightOnly`:

```powershell
& ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.ps1').Content)) -PreflightOnly
```

```bash
curl -fsSL https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.sh | bash -s -- -PreflightOnly
```

For a full parameter walkthrough and real `azd provision --preview` without
creating Azure resources, use PowerShell `-WhatIf`:

```powershell
& ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.ps1').Content)) -WhatIf
```

[Azure Developer CLI (azd)](https://aka.ms/azd) is installed automatically if
not already present. The wizard walks through selecting a subscription and
resource group, creates the Entra App Registration, deploys all Azure
infrastructure, and assigns Microsoft Graph permissions.

---

### Step 2 (Alternative) - Run from a local infra ZIP

Use this path instead of Step 2 when you need to inspect or customise the
deployment files before running them. Download the
`guest-sponsor-info-infra.zip` from
[GitHub Releases](https://github.com/workoho/spfx-guest-sponsor-info/releases),
extract it to a local directory, and run `deploy-azure.ps1` directly:

<details>
<summary>Expand local infra ZIP instructions</summary>

```powershell
# Extract the ZIP (example):
Expand-Archive -Path guest-sponsor-info-infra.zip -DestinationPath ./infra

# Run the wizard from the extracted directory:
Set-Location ./infra
./deploy-azure.ps1
```

The wizard behaviour is identical to the `install.ps1` path — the ZIP simply
provides the files locally so you can review them first.

</details>

---

### Step 3 (Optional) - Choose hosting plan parameters

If the defaults are acceptable, skip this step. The deployment wizard uses the
Consumption plan by default.

<details>
<summary>Expand hosting plan comparison and optional parameters</summary>

**Hosting plan comparison:**

| | **Consumption** (default) | **Flex Consumption** |
|---|---|---|
| Free tier | 1M exec + 400K GB-s/month | None |
| Cold starts | ~2-5 s after ~20 min idle | Eliminated with `alwaysReadyInstances=1` |
| OS | Windows | Linux only |
| Cost guard | `dailyMemoryTimeQuota` | `maximumFlexInstances` |
| Estimated cost | Free (within grant) | ~€2-5/month with 1 warm instance |

#### Optional hosting plan parameters

| Parameter | Description |
|---|---|
| `hostingPlan` | `Consumption` (default) or `FlexConsumption`. |
| `alwaysReadyInstances` | Pre-warmed instances (Flex only). `1` eliminates cold starts. Default: `1`. |
| `maximumFlexInstances` | **Required for Flex.** Hard upper bound on scale-out (cost ceiling). 1-1000. |
| `instanceMemoryMB` | `512` or `2048` (Flex only). Default: `2048`. |
| `dailyMemoryTimeQuotaGBs` | Daily GB-s budget (Consumption only). Default: `10000`. |

Check [aka.ms/flex-region](https://aka.ms/flex-region) for Flex Consumption
regional availability.

</details>

### Step 4 - Record the deployment outputs

At the end of the run, `deploy-azure.ps1` prints:

| Value | Used for |
|---|---|
| **Guest Sponsor API Base URL** (`functionAppUrl`) | Web part property pane → **Guest Sponsor API Base URL** |
| **Web Part Client ID** (`webPartClientId`) | Web part property pane → **Guest Sponsor API Client ID** |

You can also retrieve them later:

```powershell
azd env get-values
```

### Step 5 - Configure the web part

With Phase 1 and Phase 2 complete, finish the connection on the landing page
itself:

1. Open the landing page in edit mode and add **Guest Sponsor Info** if the app
  was installed on the site but the web part is not yet on the page.
2. Open the web part property pane. If the built-in Setup Wizard opens
  automatically, use that instead of entering the same values manually.
3. Paste the values from Step 4 into **Guest Sponsor API Base URL** and
  **Guest Sponsor API Client ID**.
4. Save or publish the page.
5. Verify the connection. If the page is opened by a real guest, the web part
  should load sponsor data from the API. If you are still configuring the page
  as an admin, use the web part's diagnostics and setup UI to confirm that
  token acquisition and API calls reach the expected tenant resources.

---

### Advanced scenario - Deploy from a Privileged Access Workstation (PAW)

On PAW environments where **Privileged Role Administrator** cannot be
activated (required to assign Graph app roles to Managed Identities), the
standard single-step `deploy-azure.ps1` path can still create the App
Registration via Bicep — you only need to defer the Graph role assignments.

Use this advanced path only when Graph app-role assignment must happen from a
separate privileged workstation.

<details>
<summary>Expand PAW deployment steps</summary>

> **Note:** The App Registration is always created by Bicep. Only the Graph
> app role assignment step (which requires Privileged Role Administrator) can
> be deferred. Cloud Application Administrator is always required for the
> Bicep deployment to succeed.

**Prerequisites (standard machine):**

- PowerShell 7+
- Azure Contributor + Owner on the resource group
- Entra **Cloud Application Administrator** (for Bicep to create the App Registration)

#### Step A — Run the deployment with role assignments deferred

```powershell
& ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.ps1').Content)) -SkipGraphRoleAssignments $true
```

The pre-provision hook will display a confirmation that role assignments are
deferred. The NEXT STEPS box at the end shows the Managed Identity Object ID
and the command to run in Step B.

**Required Azure roles:** Contributor + Owner on the resource group
**Required Entra roles:** Cloud Application Administrator (Bicep creates App Registration)

#### Step B — On the PAW: assign Graph permissions

Download `setup-graph-permissions.ps1` from
[GitHub Releases](https://github.com/workoho/spfx-guest-sponsor-info/releases)
(or copy it from an extracted infra ZIP), transfer it to the PAW, and run it
with the values from Step A:

```powershell
./setup-graph-permissions.ps1 `
    -ManagedIdentityObjectId <object-id-from-step-a> `
    -TenantId <your-tenant-id>
```

This script assigns Microsoft Graph application permissions to the Managed
Identity: `User.Read.All`, `Presence.Read.All` (optional),
`MailboxSettings.Read` (optional), `TeamMember.Read.All` (optional).

If the PAW cannot run `setup-graph-permissions.ps1` (for example because local
script execution is blocked), you can perform the same assignment manually with
only the `Microsoft.Graph.Authentication` module. This fallback does not need
any `Az.*` modules as long as you already have the `managedIdentityObjectId`
from Step A:

> **PAW policy note:** This fallback still depends on a working
> `Microsoft.Graph.Authentication` module on the PAW. The published module is
> not a single self-contained script; it loads PowerShell files plus multiple
> DLL dependencies (including native `msalruntime*.dll` files). AppLocker can
> explicitly control scripts and DLLs, and App Control for Business (formerly
> WDAC) can also force PowerShell into restricted language modes for untrusted
> modules. In other words: even without any `.exe` files, a hardened PAW can
> still block module installation, import, or runtime loading. If that happens,
> the practical fallback is to ask the PAW / Tier-0 administrators to provide
> an approved, working `Microsoft.Graph.Authentication` module on that device.
> How they package or pre-stage it is outside the scope of this guide.

```powershell
if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
}
Import-Module Microsoft.Graph.Authentication

$tenantId = '<your-tenant-id>'
$managedIdentityObjectId = '<object-id-from-step-a>'
$permissions = @(
  'User.Read.All'
  'Presence.Read.All'      # optional
  'MailboxSettings.Read'   # optional
  'TeamMember.Read.All'    # optional
)

Connect-MgGraph -TenantId $tenantId -Scopes 'Application.Read.All', 'AppRoleAssignment.ReadWrite.All'

$graphSp = (Invoke-MgGraphRequest -Method GET `
  -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '00000003-0000-0000-c000-000000000000'&`$select=id,appRoles" `
  -OutputType PSObject).value[0]

$existingAssignments = Invoke-MgGraphRequest -Method GET `
  -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$managedIdentityObjectId/appRoleAssignments?`$select=appRoleId" `
  -OutputType PSObject

$existingRoleIds = @{}
foreach ($assignment in $existingAssignments.value) {
  $existingRoleIds[$assignment.appRoleId] = $true
}

foreach ($permission in $permissions) {
  $appRole = $graphSp.appRoles | Where-Object {
    $_.value -eq $permission -and $_.allowedMemberTypes -contains 'Application'
  }

  if (-not $appRole) {
    Write-Warning "$permission is not available in this tenant. Skipping."
    continue
  }

  if ($existingRoleIds.ContainsKey($appRole.id)) {
    Write-Host "$permission already assigned."
    continue
  }

  Invoke-MgGraphRequest -Method POST `
    -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$managedIdentityObjectId/appRoleAssignments" `
    -Body @{
      principalId = $managedIdentityObjectId
      resourceId  = $graphSp.id
      appRoleId   = $appRole.id
    }
}
```

If the PAW cannot install or load any PowerShell modules at all, you can do the
same assignment in [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer).
This uses the same Microsoft Graph API, but sends the requests manually from
the browser:

1. Sign in to Graph Explorer as a user with **Privileged Role Administrator**.
2. Consent to the delegated permissions **Application.Read.All** and
   **AppRoleAssignment.ReadWrite.All** for Graph Explorer.
3. Retrieve the Microsoft Graph service principal and its app roles:

   ```http
   GET https://graph.microsoft.com/v1.0/servicePrincipals?$filter=displayName eq 'Microsoft Graph'&$select=id,appRoles
   ```

   Record the Graph service principal `id` and the `appRoles[].id` values for
   the permissions you want to grant (`User.Read.All`, optionally
   `Presence.Read.All`, `MailboxSettings.Read`, `TeamMember.Read.All`).

4. Retrieve existing assignments on the Managed Identity so you do not create
   duplicates:

   ```http
   GET https://graph.microsoft.com/v1.0/servicePrincipals/<managed-identity-object-id>/appRoleAssignments?$select=appRoleId
   ```

5. For each missing permission, create the assignment:

   ```http
   POST https://graph.microsoft.com/v1.0/servicePrincipals/<managed-identity-object-id>/appRoleAssignments
   Content-Type: application/json

   {
     "principalId": "<managed-identity-object-id>",
     "resourceId": "<microsoft-graph-service-principal-id>",
     "appRoleId": "<permission-app-role-id>"
   }
   ```

6. Repeat the POST only for the permissions you actually need.

This Graph Explorer path is intentionally low-level. Changes take effect
immediately and are not wrapped in any helper logic, so double-check the three
IDs before sending the request.

Use the Azure portal only for lookup and verification here:

1. **Azure Portal -> Function App -> Identity** to copy the Managed Identity
   Object ID.
2. **Microsoft Entra admin center -> Enterprise applications -> Microsoft
   Graph** if you need to inspect the Graph service principal itself.
3. **Microsoft Entra admin center -> Enterprise applications -> [managed
   identity service principal] -> Permissions** to review the resulting grants.

This guide does not document a portal-only grant path for managed identities.
Microsoft's documented assignment path for this scenario is Microsoft Graph
PowerShell or Microsoft Graph API, and the portal is best treated as the place
to find object IDs and confirm the result afterwards.

**Required Entra roles on the PAW account:**
`Privileged Role Administrator` (to assign Graph app roles)

> **Do not use enterprise application ownership as a shortcut here.** Future
> Microsoft Graph app-role changes still require `Privileged Role Administrator`
> (or a higher Entra role), because the assignment operation itself remains a
> high-privilege tenant action. In this guide, keep Graph permission changes in
> the PAW / privileged-admin workflow and prefer PIM-activated roles over
> permanent ownership grants on the managed identity's service principal.

</details>

---

## Administration and Operations

Deployment and day-2 operations are split into separate guides:

- [operations.md](operations.md) for ongoing administration, including:
  web part updates, inline map configuration, and function update playbooks.
- [security-assessment.md](security-assessment.md) for security posture,
  threat boundaries, and deployment trust assumptions.
- [telemetry.md](telemetry.md) for Customer Usage Attribution, opt-out,
  and verification steps.
