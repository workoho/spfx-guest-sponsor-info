# Deployment and Administration Guide

Full deployment, configuration, and operational reference for
SharePoint and Azure administrators.

For a quick-start overview, see the [README](../README.md).
For architecture decisions and internals, see [architecture.md](architecture.md).
For a visual system overview (including a setup checklist diagram), see
[architecture-diagram.md](architecture-diagram.md#setup--two-admin-roles-recommended-path).

## Table of Contents

- [SharePoint Deployment](#sharepoint-deployment)
- [Guest Access Requirements](#guest-access-requirements)
- [Guest Sponsor API](#guest-sponsor-api)
- [Inline Address Map (Azure Maps)](#inline-address-map-azure-maps)
- [Updating the Function](#updating-the-function)
- [Security Assessment](#security-assessment)
- [Legacy Options (no Guest Sponsor API)](#legacy-options-no-guest-sponsor-api)

---

## SharePoint Deployment

### Upload and install

1. Download the latest `guest-sponsor-info.sppkg` from
   [Releases](https://github.com/workoho/spfx-guest-sponsor-info/releases).
2. Upload it to your SharePoint **App Catalog**.
3. Check **SharePoint Admin Center → Advanced → API access** for pending
   permission requests and approve them if present.
   In many tenants the required Graph permissions (`User.Read`,
   `User.ReadBasic.All`, `Presence.Read.All`) are already pre-consented by the
   *SharePoint Online Client Extensibility Web Application Principal* — the
   queue will simply be empty and no action is needed.
4. Add the *Guest Sponsor Info* web part to a modern page.

### Tenant-wide vs. per-site deployment

The upload dialog offers *"Make this solution available to all sites in the
organization"*. **For a minimal-footprint deployment, leave this unchecked**
and add the app manually only to target sites (**Site Contents → New → App →
Guest Sponsor Info**).

| | Per site (recommended) | Tenant-wide |
|---|---|---|
| Web part in toolbox | Only sites where admin added the app | Every modern site |
| Admin effort | Once per target site | Once, at upload |
| Guest data exposure | None | None — renders `null` for member users |
| New sites | Manual per site | Automatic |

**Per-site is the recommended default** — the web part is only available where
explicitly needed. Guest landing sites are typically a known, stable set.

If you expect guest-accessible sites to grow over time and prefer to avoid
per-site admin steps, the tenant-wide option works equally well.

### Required Graph permissions

| Scope | Resource | Reason |
|---|---|---|
| `User.Read` | Microsoft Graph | Read the signed-in user's own profile and sponsor list |
| `User.ReadBasic.All` | Microsoft Graph | Fetch sponsor name, mail, job title, department, phone |
| `Presence.Read.All` | Microsoft Graph | **Optional.** Show online presence status of sponsors |

> **Why not `User.Read.All`?**
> `User.ReadBasic.All` is sufficient for sponsor profiles and does not expose
> sensitive account data such as `accountEnabled` or `onPremisesSyncEnabled`.

---

## Guest Access Requirements

The web part's JavaScript and CSS bundle is packaged with
`includeClientSideAssets: true` and re-hosted by SharePoint. By default, guest
users cannot reach those assets, causing the web part to fail silently or show
"Something went wrong" errors for guests only.

**Step 1** has two options — choose whichever fits your environment. Steps 2
and 3 are always required.

### Step 1 – Make web part assets accessible to guests

#### Option A – SharePoint Public CDN (recommended)

When the **Public CDN** is enabled with the `*/CLIENTSIDEASSETS` origin,
SharePoint rewrites asset URLs to `https://publiccdn.sharepointonline.com/…`
— a publicly accessible edge cache. Guest users (and even anonymous users on
public-access sites) can download the bundle without further configuration.

This is the **simplest and most performant** option. No claim setting changes
or App Catalog permissions are needed.

```powershell
# Install PnP PowerShell once (PowerShell 7+, cross-platform):
# Install-Module PnP.PowerShell -Scope CurrentUser

Connect-PnPOnline -Url "https://<tenant>-admin.sharepoint.com" -Interactive

# Enable the Public CDN (idempotent).
Set-PnPTenantCdnEnabled -CdnType Public -Enable $true

# Add the SPFx asset library as a public origin (idempotent).
Add-PnPTenantCdnOrigin -CdnType Public -OriginUrl "*/CLIENTSIDEASSETS"

# Verify (propagation takes ≈ 15–30 min).
Get-PnPTenantCdnOrigin -CdnType Public
```

Wait for `*/CLIENTSIDEASSETS` to appear with status `OK`. Once propagated,
asset URLs are rewritten automatically — no `.sppkg` redeployment required.

> **If you use Option A, skip Option B entirely.** Steps 2 and 3 still apply.

#### Option B – App Catalog permissions (alternative)

Use this only if the Public CDN cannot be enabled in your tenant.

##### Step 1b-i – Enable the Everyone claim

`ShowEveryoneClaim` controls the **"Everyone"** group in SharePoint's People
Picker. It covers all authenticated users in the tenant's Entra ID, **including
B2B guests who accepted their invitation**.

On tenants provisioned after March 2018 this defaults to `$false`. Check and
enable:

```powershell
Connect-PnPOnline -Url "https://<tenant>-admin.sharepoint.com" -Interactive
(Get-PnPTenant).ShowEveryoneClaim   # should return: True

# If False:
Set-PnPTenant -ShowEveryoneClaim $true
```

> **Why not `ShowAllUsersClaim`?**
> That setting controls the legacy *All Users (membership)* / *All Users
> (windows)* groups — Windows/NTLM-era claims that do **not** include B2B
> guests. `ShowEveryoneClaim` and the *Everyone* group are the modern
> equivalent.
>
> A third setting, `ShowEveryoneExceptExternalUsersClaim` (default `$true`),
> controls *Everyone except external users* — it explicitly **excludes** B2B
> guests.

##### Step 1b-ii – Enable external sharing on the App Catalog site

1. **SharePoint Admin Center → Sites → Active sites**.
2. Open the App Catalog site (typically `appcatalog`).
3. **Policies → External sharing** → set to at least **Existing guests**.

##### Step 1b-iii – Grant Read permission

1. Navigate to your App Catalog site.
2. **Site Settings → People and Groups → App Catalog Visitors**.
3. **New → Add Users** → add one of:
   - **Everyone** — broadest, simplest. Requires
     `ShowEveryoneClaim = $true`.
   - A **specific security group** containing the guest accounts —
     more targeted, no claim change needed.
4. Permission level: **Read**.

> **Pitfall — similar-sounding groups:**
>
> - *Everyone* — includes B2B guests ✓
> - *Everyone except external users* — **excludes** guests ✗
> - *All Users (membership/windows)* — members only, **excludes** guests ✗

### Step 2 – Verify external sharing on the landing page site

External sharing must be enabled at both the tenant level and at each site:

- **SharePoint Admin Center → Policies → Sharing** — at least
  *Existing guests only*.
- Confirm each site where the web part is placed has external sharing enabled.

### Step 3 – Deploy the Guest Sponsor API

The Microsoft Graph `/me/sponsors` API requires the calling user to hold a
directory role — impractical for guest accounts at scale
(see [architecture.md](architecture.md#guest-sponsor-api-recommended) for the full
analysis). The recommended solution is the **Guest Sponsor API** — a custom
Azure Function that calls Graph with application permissions on behalf of the user.

See the [Guest Sponsor API section](#guest-sponsor-api) below for full
deployment instructions.

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

### Public CDN (Option A)

The compiled JavaScript and CSS bundle at
`publiccdn.sharepointonline.com` contains no credentials, user data, or
secrets. Environment-specific values are public Microsoft URLs; tenant-specific
IDs are obtained at runtime from `pageContext`.

Enabling the Public CDN is a tenant-wide change affecting all SPFx solutions
using `includeClientSideAssets: true`. Review other solutions before enabling.

### ShowEveryoneClaim (Option B only)

Enabling is a tenant-wide change: *Everyone* becomes selectable as a
permission target. The setting itself does not grant access — it only makes
the claim group available in the People Picker. Blast radius is purely
administrative.

### App Catalog Read for Everyone (Option B only)

Authenticated guests can download the compiled bundle (the goal) and see
the list of deployed SPFx solutions (low-sensitivity metadata). They cannot
install, retract, or modify any apps.

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
