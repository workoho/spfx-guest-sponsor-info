---
layout: doc
lang: en
title: Setup Guide
permalink: /en/setup/
description: >-
  Step-by-step setup guide for a SharePoint guest landing page with Guest
  Sponsor Info and the Guest Sponsor API — Azure setup, SharePoint guest
  access, and sponsor visibility.
lead: >-
  Implementation guide for SharePoint and Azure administrators who want
  cleaner guest onboarding, reliable SharePoint guest access, and visible
  sponsors on the landing page.
github_doc: deployment.md
---

## Overview

Guest Sponsor Info setup has three phases:

| Phase | Where | Minimum role required |
|---|---|---|
| 1 — SharePoint | SharePoint Admin Center + landing page site | SharePoint Administrator |
| 2 — Guest Sponsor API | PowerShell (`install.ps1` via `iwr`) | Azure Contributor + Owner + Entra roles via PIM |
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

### Before you begin

This guide assumes a dedicated **SharePoint landing page** as the first
reliable destination for guest users. If your invitation
process or governance tooling supports a custom redirect URL, point it to that
page instead of a generic My Apps destination. My Apps is designed for app
launch, not for sponsor visibility, and a tenant-scoped Teams deep link only
helps after the guest has already been added to at least one team in your
tenant.

It also helps to align your wording early: the inviter and the sponsor are not
always the same person in guest onboarding workflows. Some tools also label the
sponsor as the "owner" of the guest relationship. If your landing page,
emails, or admin instructions mix those roles, guests may still contact the
wrong person.

[Read the sponsor vs inviter explanation]({{ '/en/sponsor-vs-inviter/' | relative_url }}).

For Microsoft Graph permissions and runtime data handling, see the
[Privacy Policy](/en/privacy/). For Azure deployment attribution and opt-out,
see [Telemetry](/en/telemetry/). If you need hands-on help instead of a
self-service rollout, see [Support](/en/support/).

## Phase 1 — SharePoint

### Decide what the guest should open first

Before you install anything, choose the SharePoint page that should serve as
the guest landing page. This is the page you should reference in onboarding
emails, governance workflows, and invitation redirects. It should become the
first reliable SharePoint destination after invitation redemption.

- Use a dedicated landing page, not a generic collaboration site home page.
- Put the web part high on the page so sponsor, backup sponsor, and contact
  context are visible immediately.
- Treat Teams links as a follow-up step from that page, not as the only first
  destination.

### Decide where the landing page should live

If you are creating a new landing page anyway, also consider whether it should
eventually live at the tenant's **root site** (`/`). Microsoft describes the
SharePoint home site as a major organizational entry point, and in newer
tenants the root site is often still flexible enough to make that decision
early. If you use `/`, the address is also easier for guests to remember
without an extra shortlink service.

That does not mean your employee portal has to live on the same page. In many
organizations, internal employee content already lives elsewhere, and the
shared landing page simply links to it. SharePoint audience targeting can also
help you show different navigation, news, and web-part content to employees
and guests on the same landing page.

Even if the root site is already occupied, this can still be a sensible
long-term direction. You can start with a communication site such as
`/sites/entrance`, establish it as the shared landing page first, and later use
Microsoft's supported root-site swap approach to move that experience to `/`
when the timing is right. If you plan for that, keep the landing page as a
modern communication site and review root-site prerequisites, permissions, and
sharing settings early.

See also:

- [Landing Page Ideas]({{ '/en/landing-page-ideas/' | relative_url }})
- [Modernize your root site](https://learn.microsoft.com/sharepoint/modern-root-site)
- [Plan, build, and launch a SharePoint home site](https://learn.microsoft.com/viva/connections/home-site-plan)

### Install from Microsoft AppSource

> **AppSource listing pending review** — The web part has been submitted to the
> Microsoft commercial marketplace and is currently awaiting approval. The
> installation steps below describe the process once the listing is live. If
> you need to deploy before approval, use the
> [deployment guide on GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/deployment.md)
> for the non-AppSource path.

The web part is available in the
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

Enable the Office 365 Public CDN.

When the Office 365 Public CDN is enabled, SharePoint replicates web part
bundles to Microsoft's edge CDN (`publiccdn.sharepointonline.com`), which is
accessible anonymously — no SharePoint authentication required. This is the
most reliable approach for guest users.

**Required role:** SharePoint Administrator.

Choose one of the following equivalent admin shells:

<details markdown="1">
<summary>Windows: SharePoint Online Management Shell</summary>

```powershell
Connect-SPOService -Url "https://<tenant>-admin.sharepoint.com"
Set-SPOTenantCdnEnabled -CdnType Public -Enable $true

# Verify the ClientSideAssets origin is included (added by default):
Get-SPOTenantCdnOrigins -CdnType Public
# Expected output includes: */CLIENTSIDEASSETS

# If the origin is missing, add it:
Add-SPOTenantCdnOrigin -CdnType Public -OriginUrl "*/CLIENTSIDEASSETS"
```

</details>

<details markdown="1">
<summary>Cross-platform: PowerShell 7 with PnP PowerShell (also works on Windows)</summary>

```powershell
Connect-PnPOnline -Url "https://<tenant>-admin.sharepoint.com" `
  -ClientId "<your-pnp-app-client-id>" -Interactive
Set-PnPTenantCdnEnabled -CdnType Public -Enable $true

# Verify the ClientSideAssets origin is included (added by default):
Get-PnPTenantCdnOrigin -CdnType Public
# Expected output includes: */CLIENTSIDEASSETS

# If the origin is missing, add it:
Add-PnPTenantCdnOrigin -CdnType Public -OriginUrl "*/CLIENTSIDEASSETS"
```

</details>

> CDN propagation takes **up to 15 minutes**. Once active, the bundle URL changes
> to `publiccdn.sharepointonline.com` automatically — no reconfiguration needed.

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

Choose one of the following equivalent admin shells:

<details markdown="1">
<summary>Windows: SharePoint Online Management Shell</summary>

```powershell
Set-SPOTenant -ShowEveryoneClaim $true
```

</details>

<details markdown="1">
<summary>Cross-platform: PowerShell 7 with PnP PowerShell (also works on Windows)</summary>

```powershell
Set-PnPTenant -ShowEveryoneClaim $true
```

</details>

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

## Phase 2 — Guest Sponsor API

The Guest Sponsor API is a companion Azure Function that proxies Microsoft
Graph calls for the web part. Guests authenticate against it using
[EasyAuth](https://learn.microsoft.com/azure/app-service/overview-authentication-authorization),
and the function queries Graph using its own Managed Identity — guests never
need directory-level permissions in your tenant.

Use `install.ps1` as the default entry point. It downloads the infra package,
runs the deployment wizard, creates the Entra App Registration, deploys the
Azure infrastructure, and assigns the required Microsoft Graph permissions —
powered by the
[Microsoft Graph Bicep extension](https://learn.microsoft.com/azure/templates/microsoft.graph/applications).

### Run the installer

<details markdown="1">
<summary>Optional: review the scripts before you run them</summary>

If you want to review the scripts before executing anything, inspect the
[install.ps1 source](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/azure-function/infra/install.ps1)
and the
[deploy-azure.ps1 source](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/azure-function/infra/deploy-azure.ps1)
on GitHub first.

`install.ps1` is a small bootstrap wrapper: it downloads the current infra
package to a temporary folder, extracts it, forwards your parameters, and then
starts `deploy-azure.ps1`.

`deploy-azure.ps1` is the actual deployment wizard: it collects or accepts the
Azure settings, ensures the required CLIs are available, runs the `azd`/Bicep
deployment, configures the app registration flow, and prints the values the
web part needs afterwards.

In short: `install.ps1` is the recommended entry point for a clean start,
while `deploy-azure.ps1` does the real deployment work once the infra package
is available locally.

</details>

Run this command in PowerShell 7+:

```powershell
& ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.ps1').Content))
```

[Azure Developer CLI (azd)](https://aka.ms/azd) is installed automatically
if it is not already present. The installer downloads the infra package,
walks through selecting a subscription and resource group, runs a
pre-provision check, executes the Bicep deployment, and prints the web part
configuration values at the end.

### Deployment wizard prompts

<details markdown="1">
<summary>Optional: show the prompt-by-prompt reference</summary>

Before the parameter prompts, the wizard may first ask you to pick the Azure
subscription that should own the deployment.

| Wizard question | What it means | Typical answer |
|---|---|---|
| **azd Environment Name** | Local `azd` environment name used to store deployment settings and outputs. It is mainly an operator-facing label, not a public Azure resource name. | Keep the default or use a short identifier such as `contoso-gsi`. |
| **Resource Group** | Target Azure resource group for all deployed resources. The wizard creates it if it does not exist. | Keep `rg-<environment>` unless your organisation already has a naming scheme. |
| **Azure Location** | Azure region for the Function App, Storage Account, monitoring resources, and related infrastructure. | Choose the region closest to your tenant or data residency needs, for example `westeurope`. |
| **Environment Tag** | Optional governance tag written to the resource group and resources. Helps with inventory, policy, and cost views. | `prod` is a sensible default; use `-` only if you intentionally do not use tags. |
| **Criticality Tag** | Optional business-criticality tag for the workload. Useful if your Azure policies or FinOps reporting use these tags. | `low` is the recommended default for most guest landing page deployments. |
| **SharePoint Tenant Name** | Short tenant name before `.sharepoint.com`. The deployment uses it to configure SharePoint-related app settings correctly. | Enter `contoso` for `contoso.sharepoint.com`. |
| **Function App Name** | Public Azure Function App host name. Must be globally unique if you set it yourself. | Leave it blank unless you need a fixed name for governance or DNS reasons. |
| **Hosting Plan** | Azure Functions runtime plan. This controls scaling model, cold-start behaviour, Linux/Windows support, and cost profile. | Start with `Consumption` unless you specifically want Flex Consumption features. |
| **Azure Maps** | Whether the deployment should create an Azure Maps resource for embedded address map rendering in the web part. | `true` if you want in-page map rendering; `false` if an external map link is enough. |
| **Function Package Version** | Release tag of the Azure Function package to deploy. | Keep `latest` unless you intentionally pin a tested release. |
| **Enable Monitoring** | Deploys Application Insights, Log Analytics, and alert resources. | `true` for production; only disable it for very minimal test setups. |
| **Graph permissions** | Whether Graph app-role assignment should happen during deployment or later with `setup-graph-permissions.ps1`. | Use the default immediate assignment if you have the Entra role; defer only when role separation or a PAW process requires it. |

Some follow-up questions only appear in specific cases:

| Shown when | Wizard question | What it means | Typical answer |
|---|---|---|---|
| Monitoring is enabled | **Enable Failure Anomalies alert** | Turns on the Application Insights smart-detector email alert for unusual failure spikes. | `false` for quieter setups, `true` if you want proactive alerting from day one. |
| Hosting plan is `FlexConsumption` | **Always-Ready Instances** | Number of pre-warmed instances. `0` allows full scale-to-zero, `1` reduces cold starts noticeably. | `1` for most production-like setups. |
| Hosting plan is `FlexConsumption` | **Maximum Flex Instances** | Hard upper scale limit for concurrent instances. This is mainly a cost and burst-control setting. | Keep the default `10` unless you know you need a tighter or higher cap. |
| Hosting plan is `FlexConsumption` | **Instance memory in MB** | Memory size per Flex instance. Higher memory increases headroom and cost. | Keep `2048` unless you explicitly optimise for minimum cost. |

</details>

### Installer workflow

<details markdown="1">
<summary>Optional: show the installer workflow</summary>

- **Downloads the infra package** and launches the deployment wizard
- **Creates the Entra App Registration** —
  `Guest Sponsor Info - SharePoint Web Part Auth`
  (via the [Microsoft Graph Bicep extension](https://learn.microsoft.com/azure/templates/microsoft.graph/applications))
- **Deploys Azure infrastructure** — Function App, Storage Account, App Service Plan
- **Assigns Microsoft Graph permissions** to the Managed Identity:
  `User.Read.All`, `Presence.Read.All` (optional), `MailboxSettings.Read`
  (optional), `TeamMember.Read.All` (optional)
- **Configures EasyAuth** on the Function App with the App Registration
- **Prints the web part configuration values** at the end

</details>

### Required Azure and Entra roles

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
> **Global Administrator** also satisfies the Entra requirements with a single
> role.

### Hosting plan options

<details markdown="1">
<summary>Optional: compare Consumption and Flex Consumption</summary>

| | **Consumption** (default) | **Flex Consumption** |
|---|---|---|
| Free tier | 1M exec + 400K GB-s/month | 250K exec + 100K GB-s/month (on-demand) |
| Cold starts | ~2-5 s after ~20 min idle | Eliminated with `alwaysReadyInstances=1` |
| OS | Windows | Linux only |
| Cost guard | `dailyMemoryTimeQuota` | `maximumFlexInstances` |
| Estimated cost | Free (within grant) | ~€2-5/month with 1 warm instance |

Check the [supported regions list](https://learn.microsoft.com/en-us/azure/azure-functions/flex-consumption-how-to#view-currently-supported-regions)
for Flex Consumption availability.

</details>

### Deployment outputs

At the end, the installer prints:

| Value | Used for |
|---|---|
| **Guest Sponsor API Base URL** | Web part property pane → **Guest Sponsor API Base URL** |
| **Web Part Client ID** | Web part property pane → **Guest Sponsor API Client ID** |

You can also retrieve them later with `azd env get-values`.

## Phase 3 — Configure the web part

### Add the web part to the landing page

After Phases 1 and 2, open the SharePoint landing page in edit mode and add
the **Guest Sponsor Info** web part to the page.

Place it near the top of the page, where guests see it before long text blocks
or downstream links. The landing page works best when it first answers the two
questions MyApps and Teams usually do not answer on their own: who the guest's
sponsors are, and how they can reach them right now.

### Connect the web part to the API

If the **Setup Wizard** is still pending, it opens automatically in edit mode.
Otherwise, open the **property pane** manually (gear icon in edit mode). Then
select **Guest Sponsor API** in the wizard or enter the values directly in the
**Guest Sponsor API** property group:

- **Guest Sponsor API Base URL** — the Base URL printed at the end of
  the `install.ps1` run (or from `azd env get-values`),
  e.g. `https://guest-sponsor-info-xyz.azurewebsites.net`
- **Guest Sponsor API Client ID** — the Web Part Client ID printed at the
  end of the `install.ps1` run (or from `azd env get-values`)

The wizard validates the format of both values before saving. If the wizard no
longer opens automatically, fill in the same values in the **Guest Sponsor API**
group of the property pane.

### Run the Guest Accessibility check

> **Guest Accessibility check**
>
> After saving, open the property pane and navigate to the
> **Guest Accessibility** panel. It runs a series of checks (CDN status,
> site permissions, external sharing) and shows the result of each with a
> recommendation. Use this to confirm that the Phase 1 prerequisites are
> working as expected.

## Further reading

For security posture and trust assumptions, see the
[security assessment on GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/security-assessment.md).

For telemetry and attribution details, see
[Telemetry]({{ '/en/telemetry/' | relative_url }}).

If something does not work as expected, see the [Support]({{ '/en/support/' | relative_url }}) page.
