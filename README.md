# Guest Sponsor Info

A SharePoint Online web part for landing pages in Microsoft Entra **resource tenants** that
shows the sponsors of the currently signed-in **guest user**.

The layout matches the SharePoint People web part:
each sponsor is shown as a card with a live photo (or initials fallback),
name, and job title.
Hovering or focusing a card reveals contact details.

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [SharePoint Online](https://www.microsoft.com/microsoft-365)
- [Microsoft Entra ID – External Identities (B2B)](https://learn.microsoft.com/azure/active-directory/external-identities/)

## Prerequisites

| Requirement | Detail |
|---|---|
| SharePoint Online | Modern team or communication site |
| Microsoft Entra | Guest accounts with one or more sponsors assigned |
| Microsoft Graph permissions | `User.Read` · `User.ReadBasic.All` |

The two Graph permissions must be approved by a tenant administrator in the
**SharePoint Admin Center → Advanced → API access** page after the solution is deployed.

## Solution

| Solution | Author(s) |
|---|---|
| `guest-sponsor-info.sppkg` | [Julian Pawlowski](https://github.com/jpawlowski) |

## Features

- **Sponsor cards** – photo (or initials + deterministic colour) · name · job title
- **Contact overlay** – email · business phone · mobile · office location · department on hover/focus
- **Guest-only in view mode** – renders `null` for member users; they see nothing
- **Edit-mode placeholder** – always visible to page authors regardless of guest status,
  so the web part can be positioned and configured on the page
- **Unavailable-sponsor handling** – sponsors whose accounts are deleted, soft-deleted, or
  disabled are not rendered; a friendly message is shown when all assigned sponsors are gone
- **Multilingual** – English · German · French · Spanish · Italian
- **Theme-aware** – `supportsThemeVariants: true` honours the site theme
- **Least-privilege Graph permissions** – `User.ReadBasic.All` instead of `User.Read.All`

## Required Permissions

| Scope | Resource | Reason |
|---|---|---|
| `User.Read` | Microsoft Graph | Read the signed-in user's own profile and sponsor list |
| `User.ReadBasic.All` | Microsoft Graph | Fetch sponsor name, mail, job title, department, and phone |
| `Presence.Read.All` | Microsoft Graph | Show the online presence status of sponsors |

> **Why not `User.Read.All`?**
> Sponsor profiles are publicly visible within the organisation.
> `User.ReadBasic.All` is sufficient and does not expose sensitive account data such as
> `accountEnabled` or `onPremisesSyncEnabled`.

## Minimal Path to Awesome

> "Minimal Path to Awesome" is a [PnP community convention](https://aka.ms/m365pnp) for SPFx
> web part README files — it means the shortest way to get the web part running.

### Deploy a pre-built release

1. Download the latest `guest-sponsor-info.sppkg` from [Releases](../../releases).
2. Upload it to your SharePoint **App Catalog**.
   Leave the *"Make this solution available to all sites in the organization"* checkbox
   **unchecked** for a minimal-footprint deployment.
   Then add the app manually to each target site: **Site Contents → New → App →
   Guest Sponsor Info**.
3. Check **SharePoint Admin Center → Advanced → API access** for any pending permission
   requests and approve them if present.
   In many tenants the required Graph permissions (`User.Read`, `User.ReadBasic.All`,
   `Presence.Read.All`) are already pre-consented by the *SharePoint Online Client
   Extensibility Web Application Principal* — the queue will simply be empty and no
   action is needed.
4. Follow the [Guest Access Requirements](#guest-access-requirements) steps below.
5. Add the *Guest Sponsor Info* web part to a modern page.

### Build from source

```bash
npm install        # install dependencies
npm run build      # compile, test, bundle, and package
```

The packaged solution is written to `sharepoint/solution/guest-sponsor-info.sppkg`.

### Local development

```bash
cp .env.example .env          # fill in SPFX_TENANT=<your-tenant>.sharepoint.com
./scripts/start.sh            # starts dev server with hot-reload
```

The dev server bundles your code locally and serves it to the **hosted workbench** on your
SharePoint Online tenant. Accept the certificate warning at `https://localhost:4321` once
per browser session, then open the hosted workbench URL printed on startup.

See [docs/architecture.md](docs/architecture.md) for the different testing scenarios
(hosted workbench as member vs. guest vs. full integration test).

## Guest Access Requirements

The web part's JavaScript and CSS bundle is packaged with `includeClientSideAssets: true`
and re-hosted by SharePoint. By default, guest users cannot reach those assets, which causes
the web part to fail silently or show "Something went wrong" errors for guests only.

**Step 1** has two options — choose the one that fits your environment; the remaining steps
are always required regardless of which option you choose.

### Step 1 – Make web part assets accessible to guests

#### Option A – SharePoint Public CDN (recommended)

When the **Public CDN** is enabled with the `*/CLIENTSIDEASSETS` origin, SharePoint
automatically rewrites asset URLs to `https://publiccdn.sharepointonline.com/…` — a
publicly accessible edge cache — instead of serving them from the App Catalog site
collection. Guest users (including those with no App Catalog permissions, and even
anonymous users on sites that allow public access) can download the bundle without any
further configuration.

This is the **simplest and most performant** option. No claim setting changes are needed.
No App Catalog permissions need to be assigned.

```powershell
# Install PnP PowerShell once (PowerShell 7+, cross-platform):
# Install-Module PnP.PowerShell -Scope CurrentUser

Connect-PnPOnline -Url "https://<tenant>-admin.sharepoint.com" -Interactive

# Enable the Public CDN (idempotent — safe to run even if already enabled).
Set-PnPTenantCdnEnabled -CdnType Public -Enable $true

# Add the SPFx asset library as a public origin.
# Microsoft may have added this automatically when CDN was first enabled; the command is idempotent.
Add-PnPTenantCdnOrigin -CdnType Public -OriginUrl "*/CLIENTSIDEASSETS"

# Verify the origin is registered (propagation takes ≈ 15–30 min).
Get-PnPTenantCdnOrigin -CdnType Public
```

Wait for `*/CLIENTSIDEASSETS` to appear in the output with status `OK`.
Once propagated, SharePoint rewrites all asset URLs automatically on the next page load —
no redeployment of the `.sppkg` is required.

> **Skip Option B entirely** if you use Option A. Steps 2 and 3 below still apply.

#### Option B – App Catalog permissions (alternative)

Use this option only if the Public CDN cannot be enabled in your tenant (for example,
your organisation has a policy against public CDN origins for SharePoint).

##### Step 1b-i – Enable the Everyone claim

`ShowEveryoneClaim` controls the **"Everyone"** group in SharePoint's People Picker.
When enabled, this group covers all authenticated users in the tenant's Microsoft Entra ID,
**including B2B guests who have accepted their invitation**.

On tenants provisioned after March 2018 this setting defaults to `$false`.
Check the current value:

```powershell
# Install PnP PowerShell once (PowerShell 7+, cross-platform):
# Install-Module PnP.PowerShell -Scope CurrentUser

Connect-PnPOnline -Url "https://<tenant>-admin.sharepoint.com" -Interactive
(Get-PnPTenant).ShowEveryoneClaim   # should return: True
```

If it returns `False`, enable it:

```powershell
Set-PnPTenant -ShowEveryoneClaim $true
```

This makes the **"Everyone"** SharePoint claim available for permission assignments.

> **Why not `ShowAllUsersClaim`?**
> Many older SharePoint guides recommend checking or enabling `ShowAllUsersClaim` for
> guest access. That setting controls the legacy *All Users (membership)* and
> *All Users (windows)* groups — Windows/NTLM-era claims from the on-premises SharePoint
> era. Those groups cover only organisation members authenticated via those older methods
> and do **not** include B2B guests.
> `ShowEveryoneClaim` and the *Everyone* group are the modern equivalent, and the only
> built-in claim that includes B2B guests.
> A third setting, `ShowEveryoneExceptExternalUsersClaim` (default `$true`), controls the
> *Everyone except external users* group that Microsoft 365 services such as Teams and
> M365 Groups use internally — as the name says, it explicitly excludes B2B guests.

##### Step 1b-ii – Enable external sharing on the App Catalog site

The App Catalog obeys SharePoint's site-level sharing settings. If external sharing is
disabled on it, guests receive HTTP 403 even after the permission below is granted.

1. Go to **SharePoint Admin Center → Sites → Active sites**.
2. Open the App Catalog site (typically named `appcatalog`).
3. Click **Policies** → **External sharing** and set it to at least **Existing guests**.

##### Step 1b-iii – Grant Read permission

1. Navigate to your App Catalog site
   (typically `https://<tenant>.sharepoint.com/sites/appcatalog`).
2. **Site Settings → People and Groups → App Catalog Visitors**.
3. **New → Add Users** and add one of the following:
   - **Everyone** – covers every authenticated user including B2B guests who have accepted
     their invitation. This is the broadest option and the simplest to configure.
     Requires `ShowEveryoneClaim = $true` (see Step 1b-i above).
   - A **specific Microsoft 365 or security group** that contains the guest accounts
     who need access. This is the more targeted, least-privilege alternative and does
     not require any claim setting change.
4. Set the permission level to **Read**.

> **Pitfall – confusingly similar group names**
> SharePoint shows several built-in groups with similar-sounding names:
>
> - *Everyone* — includes B2B guests who have accepted their invitation. Use this one. ✓
> - *Everyone except external users* — explicitly **excludes** B2B guests. ✗
> - *All Users (membership)* / *All Users (windows)* — cover organisation members only;
>   B2B guests are **not** included. ✗

### Step 2 – Verify external sharing on the landing page site

External sharing must be enabled at both the tenant level and at each relevant site:

- **SharePoint Admin Center → Policies → Sharing** – set to at least *Existing guests only*.
- Open each site where the web part is placed and confirm external sharing is enabled
  there too (same setting, per-site).
- The **App Catalog site** is covered by Step 2a above.

### Step 3 – Deploy the Sponsor API

The Microsoft Graph `/me/sponsors` API requires the calling user to hold a directory role
in addition to the `User.Read` delegated permission — a requirement that is impractical to
meet at scale for guest accounts (see [docs/architecture.md](docs/architecture.md#azure-function-proxy)
for the full analysis). The recommended solution is an **Azure Function proxy** that calls
Graph with application permissions on behalf of the user.

#### Pre-step: create the App Registration

The Azure Function uses EasyAuth (Azure App Service Authentication) to validate the caller.
EasyAuth needs an Entra App Registration as its identity provider.

Run the included script (requires `Microsoft.Graph` PowerShell module):

```powershell
./azure-function/infra/setup-app-registration.ps1 -TenantId "<your-tenant-id>"
```

Copy the **Client ID** printed at the end — you will need it in the next step.

Alternatively, create the App Registration manually in the Azure Portal:

1. **Microsoft Entra admin center → App registrations → New registration**.
2. Name: `Guest Sponsor Info Proxy`; Supported account types: *Accounts in this organizational
   directory only*.
3. After creation, open **Expose an API** → **Set** Application ID URI:
   `api://guest-sponsor-info-proxy/<clientId>`.
4. Copy the **Client ID**.

#### Deploy to Azure

Click the button below to open the Azure Portal with the ARM template pre-loaded:

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fgithub.com%2Fjpawlowski%2Fspfx-guest-sponsor-info%2Freleases%2Flatest%2Fdownload%2Fazuredeploy.json)

Alternatively, deploy from [Azure Cloud Shell](https://shell.azure.com) without any local
tooling — this also works for updates (Bicep deployments are idempotent):

```bash
az deployment group create \
  --resource-group <your-resource-group> \
  --template-uri https://raw.githubusercontent.com/jpawlowski/spfx-guest-sponsor-info/main/azure-function/infra/main.bicep \
  --parameters \
      tenantId=<your-tenant-id> \
      tenantName=<your-tenant-name> \
      functionAppName=<globally-unique-name> \
      functionClientId=<client-id-from-pre-step>
```

Fill in the parameters:

| Parameter | Description |
|---|---|
| `tenantId` | Your Entra tenant ID (GUID) |
| `tenantName` | Your tenant name without domain suffix, e.g. `contoso` |
| `functionAppName` | Globally unique name for the Function App |
| `functionClientId` | Client ID from the pre-step above |
| `packageUrl` | Leave as default (points to the latest GitHub Release ZIP) |
| `location` | Azure region |

After deployment, note the **Managed Identity Object ID** shown in the output.

#### Grant Graph permissions to the Managed Identity

```powershell
./azure-function/infra/setup-graph-permissions.ps1 \
  -ManagedIdentityObjectId "<oid-from-deployment-output>" \
  -TenantId "<your-tenant-id>"
```

This grants `User.Read.All` and `Presence.Read.All` application permissions to the
Function App's system-assigned Managed Identity.

#### Configure the web part

In the property pane of the web part (edit the page → edit the web part → **Azure Function** group):

- **Sponsor API URL**: paste the function endpoint, e.g.
  `https://guest-sponsor-info-xyz.azurewebsites.net/api/getGuestSponsors`
- **Sponsor API Client ID**: paste the Client ID from the pre-step

#### Updating the function

The simplest way to update is to re-run the deployment — it is idempotent and only
updates what has changed. From [Azure Cloud Shell](https://shell.azure.com):

```bash
az deployment group create \
  --resource-group <your-resource-group> \
  --template-uri https://raw.githubusercontent.com/jpawlowski/spfx-guest-sponsor-info/main/azure-function/infra/main.bicep \
  --parameters \
      tenantId=<your-tenant-id> \
      tenantName=<your-tenant-name> \
      functionAppName=<your-function-app-name> \
      functionClientId=<your-client-id>
```

Alternatively, update only the function package via the Azure Portal:

1. Open the Function App in the Azure Portal.
2. **Configuration → Application settings → `WEBSITE_RUN_FROM_PACKAGE`**.
3. Replace the URL with the new release asset URL (e.g.
   `https://github.com/jpawlowski/spfx-guest-sponsor-info/releases/download/vX.Y.Z/guest-sponsor-info-function.zip`).
4. **Save**. The Function App picks up the new package on next cold start.

#### Security assessment of the Azure Function approach

- The **Managed Identity** never exposes credentials — no secrets are stored anywhere.
- `User.Read.All` is an **application permission**: unlike the delegated approach, the guest
  user never holds this permission themselves. The function enforces that it only ever returns
  the calling user's own sponsors (OID is taken from the EasyAuth-validated token).
- **EasyAuth** ensures unauthenticated requests are rejected by Azure before the function code
  runs. There is no custom JWT validation code to maintain.
- Because guests do not need any Entra directory role assignment, none of the role-assignable
  group limitations, dynamic membership restrictions, or third-party SaaS trust issues described
  in Option A/B below apply.

**Overall risk level: Low.** This is the recommended approach for production deployments.

#### Legacy option (no Azure Function): assign a directory role to guests

If you cannot deploy the Azure Function, guests need a directory role to call `/me/sponsors`
directly. See the legacy instructions below.

<details>
<summary>Legacy Option A – Custom role (requires Entra ID P1 or P2)</summary>

A custom role scoped to exactly `microsoft.directory/users/sponsors/read` is the
least-privilege approach among the legacy options.

1. **Microsoft Entra admin center → Roles and admins → Roles → New custom role**.
2. Name it e.g. `Sponsor Viewer`.
3. On the *Permissions* step, search for `sponsors` and add
   `microsoft.directory/users/sponsors/read`.
4. Save the role.
5. Open the new role → **Add assignments** → select the security group containing your guests.

**Note:** Role-assignable groups require `isAssignableToRole = true` (cannot be set on an
existing group). Dynamic membership is not supported. Every new guest must be added manually
or via automation with `RoleManagement.ReadWrite.Directory` permissions.

**Privacy note:** This permission is not self-scoped — a guest with this role can also read
the sponsor relationships of other guest accounts. See the security assessment below.

</details>

<details>
<summary>Legacy Option B – Directory Readers built-in role (no P1/P2 required)</summary>

If your tenant does not have Entra ID P1 or P2:

1. **Microsoft Entra admin center → Roles and admins → Directory Readers**.
2. **Add assignments** → select the security group containing your guests.

**Warning:** Directory Readers grants much broader directory read access than just sponsors.
Only use this as a last resort and document it in your risk register.

</details>

### Security assessment of these settings

**Public CDN (`*/CLIENTSIDEASSETS` origin)** *(Option A)*
Enabling the Public CDN makes the compiled JavaScript and CSS bundle available at
`publiccdn.sharepointonline.com` without authentication. This is intentional and safe:
the bundle contains only compiled code — no credentials, no user data, no secrets.
Environment-specific values such as Graph endpoint URLs are public Microsoft URLs;
tenant-specific IDs are obtained at runtime from `pageContext` and are never embedded
in the bundle.
Enabling the Public CDN is a tenant-wide change that affects all SPFx solutions using
`includeClientSideAssets: true`. If your organisation publishes SPFx solutions that embed
sensitive configuration data directly in the bundle, review those solutions before enabling
this; this web part itself is safe to serve publicly.

**`ShowEveryoneClaim`** *(Option B only)*
This setting controls whether the *Everyone* claim group is visible in SharePoint's People
Picker. The *Everyone* group covers all authenticated users in the tenant's Microsoft Entra
ID, **including B2B guests who have accepted their invitation** — which is exactly why it is
needed here. It defaults to `$false` on tenants provisioned after March 2018.
Enabling it is a tenant-wide change: *Everyone* becomes selectable as a permission target
anywhere in the tenant. The setting itself does not grant access — it only makes the claim
group available for permission assignments.
The blast radius is purely administrative.

Do not confuse this with the two other claim settings:

- `ShowAllUsersClaim` (`$true` by default) — controls the legacy *All Users (membership)*
  and *All Users (windows)* groups, left over from on-premises SharePoint's Windows/NTLM
  authentication. They cover only organisation members authenticated via those methods and
  do **not** include B2B guests. Many older guides recommend this setting for "everyone"
  access — for modern B2B guest access it is the wrong setting.
- `ShowEveryoneExceptExternalUsersClaim` (`$true` by default) — controls *Everyone except
  external users*, which Microsoft 365 services (Teams, M365 Groups, Planner) rely on
  internally. Despite being always on, it explicitly **excludes** B2B guests — the opposite
  of what is needed here.

**Read permission for Everyone on the App Catalog** *(Option B only)*
"Everyone" covers every authenticated identity accepted by the tenant:
full members, licensed guests, and B2B guests who have accepted their invitation.
Anonymous (unauthenticated) users are explicitly excluded.

With Read permission on the App Catalog, authenticated guests can:

- Download the compiled JavaScript/CSS bundle — the explicit goal of this step.
- Browse the App Catalog site and see the list of deployed SPFx solutions (names, versions,
  deployment dates). This is low-sensitivity metadata: it reveals which tools the tenant uses,
  but contains no business data or secrets.

Guests cannot:

- Install, retract, or manage any apps (that requires Site Collection Admin or higher).
- Access any content on any other site collection.
- Modify App Catalog contents.

The compiled bundle itself contains no secrets. Environment-specific values such as Graph
endpoints are public Microsoft URLs; tenant-specific IDs (used for Teams deep links) are
obtained at runtime from `pageContext` and never hard-coded in the bundle.

**`Sponsor Viewer` custom role / Directory Readers (legacy Step 3 options)**
The `microsoft.directory/users/sponsors/read` role permission is not scoped to "only my own
sponsors". A guest who holds this role — whether via the custom `Sponsor Viewer` role or the
broader `Directory Readers` role — can also call `/users/{other-id}/sponsors` and read the
sponsor relationships of *other* guest accounts in the tenant, e.g. via Graph Explorer.

This privacy concern is the primary motivation for the recommended **Azure Function proxy**
approach (Step 3 above): with the proxy, guests never hold any Entra directory role, and
the function enforces server-side that only the caller's own sponsors are returned.

If you are using the legacy options:

What this exposes:

- Which internal employee(s) are responsible for a given guest — an internal accountability
  relationship that guests would not otherwise be able to enumerate.
- No additional profile data beyond what `User.ReadBasic.All` already allows (name, mail,
  job title). Sensitive fields such as `accountEnabled` or group memberships remain inaccessible.

`Directory Readers` (Legacy Option B) makes this significantly worse: it grants read access
to practically all directory objects and their relationships, not just sponsor assignments.
The custom `Sponsor Viewer` role (Legacy Option A) limits the extra exposure to sponsor
relationships only — it is the better choice from a data-minimisation perspective.

**Overall risk level (legacy options): Low** for Legacy Option A (custom role),
**Low–Medium** for Legacy Option B (Directory Readers, due to broader directory read access).
Organisations with strict data-minimisation requirements should use the Azure Function proxy
instead and document the rationale in their data processing register.

### Tenant-wide deployment vs. per-site-collection

The deployment dialog offers a checkbox
*"Make this solution available to all sites in the organization"*.
**For a minimal-footprint deployment, leave this unchecked** and add the app manually only
to the sites where guests actually land.

Here is the practical trade-off:

| | Per site collection (recommended) | Tenant-wide |
|---|---|---|
| Web part appears in toolbox | Only sites where admin added the app | Every modern site |
| Admin effort | Once per target site (Site Contents → Add an app) | Once, at upload |
| App Catalog Read requirement | Same — required either way | Same — required either way |
| Guest data exposure | None | None — web part renders nothing for member users |
| Adding after new sites are created | Manual per site | Automatic |

The App Catalog Read permission (Step 2) applies regardless of the deployment model —
guests must be able to fetch the bundle from the App Catalog no matter which sites the
solution is active on.

**Per-site is the recommended default** because:

- It follows the least-privilege principle: the web part is only available where it is
  explicitly needed.
- Guest landing sites are typically a known, stable set (intranet home, welcome page).

If you expect guest-accessible sites to grow over time and prefer to avoid per-site admin
steps, the tenant-wide option works equally well — the web part renders `null` for
non-guest users so there is no data exposure on member-only sites. Check the
*"Make this solution available to all sites in the organization"* checkbox during upload.

| Command | Description |
|---|---|
| `npm run build` | Full production build + unit tests + packaging |
| `npm test` | Compile and run unit tests |
| `npm start` | Start dev server (hot-reload, hosted workbench) |
| `npm run clean` | Delete all build output |
| `npm run lint` | Run all linters (TypeScript · SCSS · Markdown) |

Wrapper scripts in `scripts/` provide additional convenience (see below).

## Unit Tests

Tests are written with **Jest 29** and `react-dom/test-utils` (no additional test library needed).

```bash
npm test
```

Coverage output is written to `jest-output/coverage/`.

## Publishing a Release

Releases are created by pushing a SemVer tag. The recommended workflow:

```bash
./scripts/set-version.sh v1.2.3 --commit   # stamp version, commit, and create tag
git push && git push --tags                 # triggers the release GitHub Actions workflow
```

The workflow automatically:

1. Generates release notes from the Conventional Commit history (via [git-cliff](https://git-cliff.org)).
2. Builds the production `.sppkg`.
3. Creates a GitHub Release with the notes and the `.sppkg` attached.

For the **first release**, this works even with a single commit and no prior tags.
Preview what the release notes will look like before tagging:

```bash
./scripts/release-notes.sh
```

## Continuous Integration

| Workflow | Trigger | What it does |
|---|---|---|
| `ci.yml` | Push / PR to `main` | Build · test · upload coverage |
| `release.yml` | `v*` SemVer tag | Bump version · build · create GitHub Release + `.sppkg` asset |

## References

- [SharePoint Framework documentation](https://aka.ms/spfx)
- [Microsoft Graph – List sponsors](https://learn.microsoft.com/graph/api/user-list-sponsors)
- [Microsoft Entra B2B sponsors](https://learn.microsoft.com/azure/active-directory/external-identities/b2b-sponsors)
- [Use Microsoft Graph in your SPFx solution](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp)

## License

MIT — see [LICENSE](LICENSE) for details.

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED,
INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY,
OR NON-INFRINGEMENT.**
