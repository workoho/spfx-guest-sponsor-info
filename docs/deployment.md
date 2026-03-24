# Deployment and Administration Guide

Full deployment, configuration, and operational reference for
SharePoint and Azure administrators.

For a quick-start overview, see the [README](../README.md).
For architecture decisions and internals, see [architecture.md](architecture.md).
For a visual system overview (including a setup checklist diagram), see
[architecture-diagram.md](architecture-diagram.md#setup--two-admin-roles-recommended-path).

## Table of Contents

- [SharePoint Deployment](#sharepoint-deployment)
- [Guest Sponsor API](#guest-sponsor-api)
- [Inline Address Map (Azure Maps)](#inline-address-map-azure-maps)
- [Updating the Function](#updating-the-function)
- [Security Assessment](#security-assessment)
- [Legacy Options (no Guest Sponsor API)](#legacy-options-no-guest-sponsor-api)

---

## SharePoint Deployment

### Enable the Site Collection App Catalog

The web part's bundle is hosted in a **Site Collection App Catalog** directly
on the guest landing page site. Because guest users already need read access
to that site, no CDN configuration or additional permissions on the global
App Catalog are required.

Enable the Site Collection App Catalog once as a **SharePoint Administrator**.
There is no GUI option for this step — PowerShell is required.

On Windows, the **SharePoint Online Management Shell** is the simplest option
— it works with your existing credentials and requires no additional setup:

```powershell
# Install once:
# Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser

Connect-SPOService -Url "https://<tenant>-admin.sharepoint.com"
Add-SPOSiteCollectionAppCatalog -Site "https://<tenant>.sharepoint.com/sites/<landing-site>"
```

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

### Upload and install

1. Download the latest `guest-sponsor-info.sppkg` from
   [Releases](https://github.com/workoho/spfx-guest-sponsor-info/releases).
2. Navigate to the landing page site → **Site Contents** →
   **Apps for SharePoint** and upload the `.sppkg` file.
   *(The Site Collection App Catalog library is also directly accessible at
   `https://<tenant>.sharepoint.com/sites/<landing-site>/AppCatalog/`.)*
3. The web part becomes available on all pages within this site collection
   immediately — no additional “Add App” step is required.

### Verify guest access to the landing page site

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

> **Tip:** [EasyLife 365 Collaboration](https://easylife365.cloud/) is purpose-built
> for automated Microsoft 365 guest lifecycle management and can ensure that a static
> site-access group is populated for every guest invitation — including guests who
> would otherwise be invited implicitly through Teams or SharePoint.
> [Workoho](https://www.workoho.com/), the author of this web part, is a Platinum
> implementation partner of EasyLife 365.

### External sharing

SharePoint's tenant-level sharing setting acts as a **ceiling**: individual
sites cannot be more permissive than the tenant allows, but they can be
more restrictive. What matters here is the setting on the landing page site
itself:

- **Active sites → [landing page site] → Policies → External sharing** —
  set to at least *Existing guests only*.

If that option is greyed out or missing, the tenant-level ceiling is too
restrictive. Raise it under **SharePoint Admin Center → Policies → Sharing**
to at least *Existing guests only*, then configure the site.

### Required Graph permissions

All three permissions below are **pre-authorized by Microsoft** for the
*SharePoint Online Client Extensibility Web Application Principal*. No manual
consent in **SharePoint Admin Center → Advanced → API access** is needed —
the queue will simply be empty.

| Scope | Resource | Reason |
|---|---|---|
| `User.Read` | Microsoft Graph | Read the signed-in user's own profile and sponsor list |
| `User.ReadBasic.All` | Microsoft Graph | Fetch sponsor name, mail, job title, department, phone |
| `Presence.Read.All` | Microsoft Graph | **Optional.** Show online presence status of sponsors |

> **Why not `User.Read.All`?**
> `User.ReadBasic.All` is sufficient for sponsor profiles and does not expose
> sensitive account data such as `accountEnabled` or `onPremisesSyncEnabled`.

---

## Guest Sponsor API

> The [Setup diagram](architecture-diagram.md#setup--two-admin-roles-recommended-path)
> gives a visual overview of all admin roles and deployment steps involved.

### Pre-step: create the App Registration

The Azure Function uses [EasyAuth](https://learn.microsoft.com/azure/app-service/overview-authentication-authorization)
(Azure App Service Authentication). EasyAuth needs an Entra App Registration
as its identity provider.

Run the included script (requires `Microsoft.Graph` PowerShell module):

```powershell
./azure-function/infra/setup-app-registration.ps1 -TenantId "<your-tenant-id>"
```

Copy the **Client ID** printed at the end.

<details>
<summary>Manual alternative (Azure Portal)</summary>

1. **Microsoft Entra admin center → App registrations → New registration**.
2. Name: `Guest Sponsor Info Proxy`; Supported account types: *Accounts in
   this organizational directory only*.
3. **Expose an API → Set** Application ID URI:
   `api://guest-sponsor-info-proxy/<clientId>`.
4. Copy the **Client ID** — this is used as `ALLOWED_AUDIENCE`.

</details>

### Deploy to Azure

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2Fjpawlowski%2Fspfx-guest-sponsor-info%2Fmain%2Fazure-function%2Finfra%2Fazuredeploy.json)

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

### Required parameters

| Parameter | Description |
|---|---|
| `tenantId` | Your Entra tenant ID (GUID) |
| `tenantName` | Tenant name without domain suffix, e.g. `contoso` |
| `functionAppName` | Globally unique name for the Function App |
| `functionClientId` | Client ID from the pre-step |
| `appVersion` | `"latest"` (default) or pinned SemVer without `v`, e.g. `"1.4.2"` |
| `location` | Azure region |

### Optional hosting plan parameters

| Parameter | Description |
|---|---|
| `hostingPlan` | `Consumption` (default) or `FlexConsumption`. See below. |
| `alwaysReadyInstances` | Pre-warmed instances (Flex only). `1` eliminates cold starts. Default: `1`. |
| `maximumFlexInstances` | **Required for Flex.** Hard upper bound on scale-out (cost ceiling). 1–1000. |
| `instanceMemoryMB` | `512` or `2048` (Flex only). Default: `2048`. |
| `dailyMemoryTimeQuotaGBs` | Daily GB-s budget (Consumption only). Default: `10000`. |

### Flex Consumption plan

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
(~2–5 min). The provisioning script resource is retained for 2 hours for
troubleshooting, then removed automatically.

| | **Consumption** (default) | **Flex Consumption** |
|---|---|---|
| Free tier | 1M exec + 400K GB-s/month | None |
| Cold starts | ~2–5 s after ~20 min idle | Eliminated with `alwaysReadyInstances=1` |
| OS | Windows | Linux only |
| Deploy to Azure button | Supported | Supported |
| Cost guard | `dailyMemoryTimeQuota` | `maximumFlexInstances` |
| Estimated cost | Free (within grant) | ~€2–5/month with 1 warm instance |

### Deployment outputs

After deployment, open **Resource Group → Deployments → select deployment →
Outputs**:

| Output | Used for |
|---|---|
| `managedIdentityObjectId` | Required for `setup-graph-permissions.ps1` (next step) |
| `functionAppUrl` | Web part property pane → **Guest Sponsor API Base URL** |
| `sponsorApiUrl` | Full endpoint URL (for curl/Postman health checks) |

### Grant Graph permissions and configure the App Registration

```powershell
./azure-function/infra/setup-graph-permissions.ps1 `
  -ManagedIdentityObjectId "<oid-from-deployment-output>" `
  -TenantId "<your-tenant-id>" `
  -FunctionAppClientId "<client-id-from-pre-step>"
```

This script:

1. **Managed Identity Graph permissions** — assigns `User.Read.All`,
   `Presence.Read.All` (optional; requires Microsoft Teams), and
   `MailboxSettings.Read` (optional; filters shared/room/equipment mailboxes).
2. **App Registration setup for silent token acquisition** — exposes a
   `user_impersonation` scope and pre-authorizes *SharePoint Online Web Client
   Extensibility*. Without this, the web part would trigger full page reloads.

### Configure the web part

In the property pane (**Guest Sponsor API** group):

- **Guest Sponsor API Base URL** — e.g.
  `https://guest-sponsor-info-xyz.azurewebsites.net`
- **Guest Sponsor API Client ID** — the Client ID from the pre-step

---

## Inline Address Map (Azure Maps)

The ARM template deploys an Azure Maps account by default
(`deployAzureMaps=true`).

### Enable map rendering

1. Get the key:

   ```bash
   az maps account keys list \
     -g <resource-group> \
     -n <azure-maps-account-name> \
     --query primaryKey -o tsv
   ```

2. In the web part property pane:
   - Enable **Show address map preview**
   - Paste the key into **Azure Maps subscription key**
   - Choose fallback provider (`Bing`, `Google`, `Apple`,
     `OpenStreetMap`, `HERE`)

Without an Azure Maps key (or when geocoding fails), the card shows an
external map link fallback.

### CSP-restricted environments

Allow at least:

- `https://atlas.microsoft.com` (geocoding + static map image)
- The selected external map provider domain for fallback links

### Quick decision guide

1. Keep `deployAzureMaps=true` — deploying Azure Maps costs nothing initially.
2. Enter the key in the web part only when you want inline maps.
3. No key → external provider link fallback is shown automatically.

Billing: Azure Maps pricing is request-based with a free monthly quota (S0).
No key configured in the web part → no Azure Maps requests issued.

---

## Updating the Function

### Consumption plan

The Function App uses `WEBSITE_RUN_FROM_PACKAGE` pointing to the latest
GitHub Release ZIP. A restart pulls the current ZIP:

```bash
az functionapp restart \
  --resource-group <your-resource-group> \
  --name <your-function-app-name>
```

Or from the Azure Portal: **Function App → Overview → Restart**.

### Flex Consumption plan

Re-deploy the ARM template with a pinned `appVersion`:

```bash
az deployment group create \
  --resource-group <your-resource-group> \
  --template-uri https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/azuredeploy.json \
  --parameters \
      tenantId=<your-tenant-id> \
      tenantName=<your-tenant-name> \
      functionAppName=<your-function-app-name> \
      functionClientId=<your-client-id> \
      hostingPlan=FlexConsumption \
      maximumFlexInstances=10 \
      appVersion=1.x.y
```

<details>
<summary>Manual upload via Azure Portal or CLI</summary>

**Via the Azure Portal:**

1. Open Storage Account → **Containers** → `app-package`.
2. **Upload** → select the ZIP from the
   [Releases page](https://github.com/workoho/spfx-guest-sponsor-info/releases).
3. **Advanced** → Blob name: `function.zip` → enable **Overwrite** → Upload.

**Via Azure CLI ([Cloud Shell](https://shell.azure.com)):**

```bash
curl -sSfL -o function.zip \
  https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/guest-sponsor-info-function.zip

az storage blob upload \
  --account-name <storage-account-name> \
  --container-name app-package \
  --name function.zip \
  --file function.zip \
  --auth-mode login \
  --overwrite
```

</details>

<details>
<summary>Infrastructure changed? Re-run the full deployment</summary>

If a release states that Azure infrastructure was updated, re-run the ARM
deployment (idempotent):

```bash
az deployment group create \
  --resource-group <your-resource-group> \
  --template-uri https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/azuredeploy.json \
  --parameters \
      tenantId=<your-tenant-id> \
      tenantName=<your-tenant-name> \
      functionAppName=<your-function-app-name> \
      functionClientId=<your-client-id>
```

For Deployment Stacks, use `az stack group create` with the same parameters.

To remove all deployed resources:

```bash
az stack group delete \
  --name guest-sponsor-info \
  --resource-group <your-resource-group> \
  --action-on-unmanage deleteResources \
  --yes
```

</details>

---

## Security Assessment

### Guest Sponsor API approach (recommended)

- **Managed Identity** — no secrets stored anywhere.
- `User.Read.All` is an **application permission**: the guest user never holds
  it. The function returns only the calling user's own sponsors (OID from the
  EasyAuth-validated token).
- **EasyAuth** rejects unauthenticated requests before function code runs.
- **CORS** restricted to the tenant's SharePoint origin.
- Caller OID redacted in function logs; structured reason codes for failures.

**Overall risk level: Low.** Recommended for production.

### Site Collection App Catalog

The web part bundle is served from the guest landing page site itself.
Guests cannot list or modify apps in the Site Collection App Catalog — they
can only download the compiled bundle via their normal site read access.
The bundle contains no credentials, user data, or secrets; environment-specific
values are public Microsoft URLs and tenant-specific IDs obtained at runtime
from `pageContext`.

---

## Legacy Options (no Guest Sponsor API)

If you cannot deploy the Guest Sponsor API, guests need an Entra directory role
to call `/me/sponsors` directly. The Guest Sponsor API approach is strongly
preferred — see the
[security assessment in architecture.md](architecture.md#security) for why.

<details>
<summary>Legacy Option A – Custom role (requires Entra ID P1 or P2)</summary>

A custom role scoped to
`microsoft.directory/users/sponsors/read` is the least-privilege legacy
approach.

1. **Microsoft Entra admin center → Roles and admins → New custom role**.
2. Name: `Sponsor Viewer`.
3. Permissions → search `sponsors` → add
   `microsoft.directory/users/sponsors/read`.
4. Open the role → **Add assignments** → select the security group containing
   your guests.

**Caveats:**

- Role-assignable groups require `isAssignableToRole = true` (cannot be set on
  existing groups). Dynamic membership is not supported.
- This permission is **not self-scoped**: guests can read other guests' sponsor
  relationships.

**Risk level: Low.**

</details>

<details>
<summary>Legacy Option B – Directory Readers built-in role (no P1/P2
required)</summary>

1. **Microsoft Entra admin center → Roles and admins → Directory Readers**.
2. **Add assignments** → select the security group containing your guests.

**Warning:** Directory Readers grants much broader directory read access than
just sponsors. Only use as a last resort.

**Risk level: Low–Medium** (broader directory exposure).

</details>
