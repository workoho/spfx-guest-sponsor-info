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
| 2 — Guest Sponsor API | PowerShell (`deploy-azure.ps1`) | Azure Contributor + Owner + Entra roles via PIM |
| 3 — Web part | SharePoint landing page (edit mode) | Site Owner |

> **The web part includes a built-in Setup Wizard**
>
> The first time you add the web part to a page, a **Setup Wizard** opens
> automatically. It walks you through choosing between production mode
> (Guest Sponsor API) and demo mode, shows the deploy command with a
> copy button, and lets you enter the API credentials at the end.
> This page is the full reference that the wizard links to — work
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

The `deploy-azure.ps1` script handles the full deployment in one step:
creating the Entra App Registration, deploying all Azure infrastructure, and
assigning the required Microsoft Graph permissions — powered by the
[Microsoft Graph Bicep extension](https://learn.microsoft.com/azure/templates/microsoft.graph/applications).

For restricted environments (Privileged Access Workstations) where the Entra
directory roles required by the Bicep Graph extension cannot be activated, see
[Deploying from a PAW](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/deployment.md#deploying-from-a-privileged-access-workstation-paw)
in the deployment guide.

### Deploy with deploy-azure.ps1

From a local clone of the repository, run:

```powershell
./deploy-azure.ps1
```

[Azure Developer CLI (azd)](https://aka.ms/azd) is installed automatically
if it is not already present. The script walks through selecting a
subscription and resource group, runs a pre-provision check, executes the
Bicep deployment, and prints the web part configuration values at the end.

#### What the script does

- **Creates the Entra App Registration** —
  `Guest Sponsor Info - SharePoint Web Part Auth`
  (via the [Microsoft Graph Bicep extension](https://learn.microsoft.com/azure/templates/microsoft.graph/applications))
- **Deploys Azure infrastructure** — Function App, Storage Account, App Service Plan
- **Assigns Microsoft Graph permissions** to the Managed Identity:
  `User.Read.All`, `Presence.Read.All` (optional), `MailboxSettings.Read`
  (optional), `TeamMember.Read.All` (optional)
- **Configures EasyAuth** on the Function App with the App Registration
- **Prints the web part configuration values** at the end

#### Required Azure and Entra roles

| Scope | Required role |
|---|---|
| Resource group | **Contributor** |
| Resource group | **Owner** (or User Access Administrator) — for Managed Identity role assignments |
| Entra ID | **Cloud Application Administrator** — to create and configure the App Registration |
| Entra ID | **Privileged Role Administrator** — to assign Graph app roles to the Managed Identity |

> **PIM tip:** If your organisation uses
> [Privileged Identity Management (PIM)](https://learn.microsoft.com/entra/id-governance/privileged-identity-management/pim-configure),
> activate the required Entra roles before running the script. The
> pre-provision hook checks your active directory roles and warns if any
> are missing.
>
> *Alternatively,* **Global Administrator** satisfies all Entra requirements
> with a single role.

#### Hosting plan options

| | **Consumption** (default) | **Flex Consumption** |
|---|---|---|
| Free tier | 1M exec + 400K GB-s/month | 250K exec + 100K GB-s/month (on-demand) |
| Cold starts | ~2-5 s after ~20 min idle | Eliminated with `alwaysReadyInstances=1` |
| OS | Windows | Linux only |
| Cost guard | `dailyMemoryTimeQuota` | `maximumFlexInstances` |
| Estimated cost | Free (within grant) | ~€2-5/month with 1 warm instance |

Check the [supported regions list](https://learn.microsoft.com/en-us/azure/azure-functions/flex-consumption-how-to#view-currently-supported-regions)
for Flex Consumption availability.

#### Deployment outputs

At the end of the run, `deploy-azure.ps1` prints:

| Value | Used for |
|---|---|
| **Guest Sponsor API Base URL** | Web part property pane → **Guest Sponsor API Base URL** |
| **Web Part Client ID** | Web part property pane → **Guest Sponsor API Client ID** |

You can also retrieve them later with `azd env get-values`.

---

## Phase 3 — Configure the web part

With Phases 1 and 2 complete, open the SharePoint landing page in edit mode
and add the **Guest Sponsor Info** web part to the page.

The **Setup Wizard** will open automatically (it appears whenever the API
URL has not been configured yet). Select **Guest Sponsor API**, then advance
through the wizard steps to the **Connect** screen and enter:

- **Guest Sponsor API Base URL** — the Base URL printed at the end of
  `deploy-azure.ps1` (or from `azd env get-values`),
  e.g. `https://guest-sponsor-info-xyz.azurewebsites.net`
- **Guest Sponsor API Client ID** — the Web Part Client ID printed at the
  end of `deploy-azure.ps1` (or from `azd env get-values`)

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

For security posture and trust assumptions, see the
[security assessment on GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/security-assessment.md).

For telemetry and attribution details, see
[Telemetry]({{ '/en/telemetry/' | relative_url }}).

If something does not work as expected, see the [Support]({{ '/en/support/' | relative_url }}) page.
