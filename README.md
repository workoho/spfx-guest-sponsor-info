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
- **Deleted-sponsor handling** – existence is verified without `User.Read.All`; accounts
  that have been deleted are counted and a friendly message is shown when all sponsors are gone
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

Because the web part's JavaScript bundle is served from the **App Catalog site collection**,
guest users must be able to reach that site.
By default they cannot, which causes the web part to fail silently or show
"Something went wrong" errors only for guests.

### Step 1 – Verify the All Users claim (modern tenants: already enabled)

`ShowAllUsersClaim` defaults to `$true` in all modern SharePoint Online tenants
(provisioned after ~2018). No action is needed unless your tenant is unusually old or the
setting was explicitly disabled.

To verify the current value:

```powershell
Connect-SPOService -Url "https://<tenant>-admin.sharepoint.com"
(Get-SPOTenant).ShowAllUsersClaim   # should return: True
```

If it returns `False`, re-enable it:

```powershell
Set-SPOTenant -ShowAllUsersClaim $true
```

This makes the **"All Users (membership)"** SharePoint claim available for use in permission
assignments. That claim covers every authenticated user including B2B guests.

### Step 2 – Grant guests Read access to the App Catalog

Two things must be done: enable external sharing on the App Catalog site itself, then
grant the Read permission.

#### Step 2a – Enable external sharing on the App Catalog site

The App Catalog is a regular site collection and obeys SharePoint's site-level sharing
settings. If external sharing is disabled on it, guests receive HTTP 403 even after the
permission in Step 2b has been granted.

1. Go to **SharePoint Admin Center → Sites → Active sites**.
2. Open the App Catalog site (typically named `appcatalog`).
3. Click **Policies** → **External sharing** and set it to at least **Existing guests**.

#### Step 2b – Grant Read permission

1. Navigate to your App Catalog site
   (typically `https://<tenant>.sharepoint.com/sites/appcatalog`).
2. **Site Settings → People and Groups → App Catalog Visitors**.
3. **New → Add Users** and add one of the following:
   - **All Users (membership)** – covers every authenticated user including B2B guests.
     This is the broadest option and the simplest to configure.
   - A **specific Microsoft 365 or security group** that contains the guest accounts
     who need access. This is the more targeted, least-privilege alternative.
4. Set the permission level to **Read**.

> **Pitfall – "Everyone except external users"**
> SharePoint shows two similarly named built-in groups. *Everyone except external users*
> explicitly **excludes** B2B guests — despite the name sounding inclusive.
> Only *All Users (membership)* (or a named group that contains your guests) will work.

This allows authenticated guest users to download the web part bundle from the App Catalog.

> **Why is this needed?**
> `includeClientSideAssets: true` packages all JavaScript and CSS directly into the `.sppkg`.
> SharePoint then re-hosts those assets on the App Catalog site collection.
> A guest user loading the web part will make a request to the App Catalog to fetch that bundle;
> without Read access the request returns 403 and the web part never loads.

### Step 3 – Verify external sharing on the landing page site

External sharing must be enabled at both the tenant level and at each relevant site:

- **SharePoint Admin Center → Policies → Sharing** – set to at least *New and existing guests*.
- Open each site where the web part is placed and confirm external sharing is enabled
  there too (same setting, per-site).
- The **App Catalog site** is covered by Step 2a above.

### Security assessment of these settings

**`ShowAllUsersClaim`**
This setting only controls whether the *All Users (membership)* claim is selectable in
SharePoint permission UIs — it does not grant anyone any new access by itself.
It is `True` by default on all modern tenants, so enabling it does not change the
security posture for the vast majority of organizations. Without it, SharePoint admins
cannot assign permissions to a group that includes guests; they would have to invite every
guest account individually to the App Catalog.
The blast radius is purely administrative: the claim becomes usable for permission
assignments across the entire tenant.

**Read permission for All Users (membership) on the App Catalog**
"All Users (membership)" covers every authenticated identity accepted by the tenant:
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

**Overall risk level: Low.**
The configuration widens read access to non-sensitive deployment metadata for authenticated
identities only. Regular security reviews of the App Catalog Visitors group are recommended
to ensure no overly broad permissions accumulate over time.

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
