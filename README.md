# Guest Sponsor Info for Microsoft Entra B2B

<p align="center">
  <img src="images/icon-rounded.svg" width="96" height="96" alt="Guest Sponsor Info icon" />
</p>

A SharePoint Online web part for landing pages in Microsoft Entra **resource
tenants** that shows the sponsors of the currently signed-in **guest user**.

Each sponsor is rendered as a card with a live profile photo (or initials
fallback), name, and job title. Hovering or focusing a card reveals contact
details. The layout matches the built-in SharePoint People web part.

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [SharePoint Online](https://www.microsoft.com/microsoft-365)
- [Microsoft Entra ID – External Identities (B2B)](https://learn.microsoft.com/azure/active-directory/external-identities/)
- [Azure Functions](https://learn.microsoft.com/azure/azure-functions/functions-overview)
- [Microsoft Graph](https://learn.microsoft.com/graph/overview)

## Solution

| Solution | Author(s) |
|---|---|
| `guest-sponsor-info.sppkg` | [Workoho GmbH](https://github.com/workoho) ([Julian Pawlowski](https://github.com/jpawlowski)) |

## Prerequisites

| Requirement | Detail |
|---|---|
| SharePoint Online | Modern team or communication site |
| Microsoft Entra | Guest accounts with one or more sponsors assigned |
| Azure subscription | Required to host the included Guest Sponsor API (Azure Function) |
| Azure Function permissions for Microsoft Graph | `User.Read.All` granted to the Function's Managed Identity; optional: `Presence.Read.All` · `MailboxSettings.Read` · `TeamMember.Read.All` |

## Features

Works out of the box with any standard Microsoft 365 environment — no
third-party tools or paid add-ons required. Every feature below relies solely
on Microsoft Graph, SharePoint Framework, and the included Guest Sponsor API.
[EasyLife 365 Collaboration](https://easylife365.cloud/products/collaboration/) pairs
naturally as a companion: it automates sponsor assignments and the full guest
lifecycle so the right information stays accurate over time — solid on its
own, stronger together.
[Workoho](https://workoho.com/?utm_source=gsiw&utm_medium=docs&utm_campaign=repo&utm_content=readme), the author of
this web part, is a Platinum EasyLife 365 partner and happy to advise.

- **"Who is my contact here?"** — guests see their sponsor's photo, name, and
  job title directly on the SharePoint landing page — no searching, no asking
  around
- **Familiar look and feel** — the contact card is modelled after the Microsoft
  Teams profile card and the SharePoint People web part, so it feels native
  rather than like a custom add-on
- **Everything needed to reach out** — email, phone, and Teams Chat / Call
  buttons in one click; office address with map preview (Azure Maps) or a link
  to the map provider of your choice (Google, Apple, Bing, OpenStreetMap)
- **Fully configurable contact details** — page editors choose exactly which
  fields are shown: phone numbers, full address broken down by street, city,
  state, ZIP, and country, map preview, manager section, presence status,
  photos — all individually toggleable in the property pane
- **Preview mode for page editors** — editors don't need to be a guest to see
  how the web part will look; a demo mode shows realistic sample cards so the
  page can be designed and reviewed without any real guest account
- **"Can I already use Teams here?"** — Microsoft requires guests to be
  [added to at least one Team](https://learn.microsoft.com/microsoftteams/guest-access)
  before they can use any Teams features in the host tenant; if that hasn't
  happened yet, the web part shows a friendly notice explaining exactly what is
  happening and what to do next — instead of a confusing error
- **Sponsor's manager visible too** — guests can see who the sponsor reports to,
  giving them a clearer picture of their contact's role in the organisation
- **Live availability** — the sponsor's current Teams status (available, busy,
  in a meeting, out of office …) is shown in real time so guests know whether
  to chat, call, or send an email
- **Only active people, no stale entries** — the web part filters out disabled
  accounts and shared/room mailboxes, so guests always see real, reachable
  colleagues — not former employees or system accounts that are still lingering
  in the sponsor list; if all assigned sponsors have since left the organisation,
  the guest receives a clear notice instead of an empty page
- **Automatic sponsor delegation** — when sponsors are stored in priority order
  (primary, secondary, tertiary … as tools like
  [EasyLife 365 Collaboration](https://easylife365.cloud/products/collaboration/) do), the web part honours that
  order: if a higher-priority sponsor is unavailable, the next active one steps
  in automatically — no configuration change needed; unavailable sponsors are
  still shown as read-only tiles so the guest sees the full picture
- **Only shown to guests** — member users see nothing; the web part is invisible
  unless the visitor is actually a guest account
- **Works without giving guests extra permissions** — if you've ever tried to
  build something like this, you'll know that guests can't read their own sponsor
  list with default permissions, and there's no good way to grant that right
  granularly. The included **Guest Sponsor API** acts as a secure proxy so that
  problem never reaches your guests (powered by a custom Azure Function)
- **17 languages** — including an informal salutation mode (`du`/`tu`) for
  German, French, Spanish, Italian, and Dutch

> **Tip:** Want to automate who gets assigned as a sponsor — and keep those
> assignments current over time? [EasyLife 365 Collaboration](https://easylife365.cloud/)
> handles the full lifecycle of Microsoft 365 collaboration workspaces — Teams, SharePoint
> team sites, Viva Engage communities, and more — including guest onboarding and sponsor
> management. This web part then takes care of the guest-facing experience.
> [Workoho](https://workoho.com/?utm_source=gsiw&utm_medium=docs&utm_campaign=repo&utm_content=readme),
> the author of this web part, is a Platinum sales and
> implementation partner of EasyLife 365 and happy to help.
>
> Full feature descriptions, design decisions, and the problem this solves:
> **[docs/features.md](docs/features.md)**

## Minimal Path to Awesome

> For a visual overview of all setup steps and required admin roles, see the
> [Setup diagram](docs/architecture-diagram.md#setup--two-admin-roles-recommended-path).

### 1. Install the web part

The web part is available in the
[**Microsoft commercial marketplace (AppSource)**](https://appsource.microsoft.com/).

**Install via SharePoint Admin Center:**

1. Open **SharePoint Admin Center → More features → Apps → Open**.
2. Click **Get apps from marketplace** and search for *Guest Sponsor Info*.
3. Select the app and click **Get it now** — it deploys to the Tenant App
   Catalog automatically. A Site Collection Administrator must then add the
   app to the desired site via **Site Contents → Add an app → Guest Sponsor
   Info** before the web part appears in the page editor.

The web part requests no Microsoft Graph permissions of its own — the
**API access** queue will remain empty. All Graph data is fetched
server-side by the companion Azure Function using its Managed Identity.

**Enable the Office 365 Public CDN** so that guest users can load the web part
bundle. By default, the Tenant App Catalog asset library is not accessible to
B2B guests before authentication. The Public CDN serves the bundle anonymously
from Microsoft's edge network:

PowerShell prerequisites for the commands below:

- **Windows / SharePoint Online Management Shell:** install
  `Microsoft.Online.SharePoint.PowerShell` once.
- Prefer `Install-PSResource` for module installation. On Windows PowerShell
  5.1, first update
  [PowerShellGet / PSResourceGet](https://learn.microsoft.com/powershell/gallery/powershellget/install-powershellget?view=powershellget-3.x)
  because `Install-PSResource` isn't available out of the box.
- **PnP path:** use **PowerShell 7+** even on Windows, install
  [PnP PowerShell](https://pnp.github.io/powershell/) once, and
  [register your own app in Microsoft Entra](https://pnp.github.io/powershell/articles/registerapplication.html)
  because `Connect-PnPOnline -Interactive` requires a client ID.

```powershell
# Install once: Install-PSResource Microsoft.Online.SharePoint.PowerShell -Repository PSGallery -Scope CurrentUser
Connect-SPOService -Url "https://<tenant>-admin.sharepoint.com"
Set-SPOTenantCdnEnabled -CdnType Public -Enable $true
```

> **Alternative deployment options** (Everyone on Tenant App Catalog, or Site
> Collection App Catalog from GitHub Releases):
> **[docs/deployment.md — SharePoint Deployment](docs/deployment.md#make-the-web-part-accessible-to-guest-users)**

### 2. Verify external sharing

What matters is the sharing setting on the landing page site itself:
**SharePoint Admin Center → Active sites → [site] → Policies → External
sharing** — set to at least *Existing guests only*. If that option is greyed
out, the tenant-level ceiling (**Policies → Sharing**) needs to be raised
first.

### 3. Verify guest access to the landing page site

If your landing page site is already serving guests, Visitor access is most
likely in place — but it's worth checking that it's configured in a way that
works reliably for newly invited users too.

> **New to the landing page?** Use a **Communication Site** (not a Team
> Site) — it has a clean Visitor permission model with no attached Microsoft
> 365 group. The instructions in this step then apply from scratch.

Guests need at least **Read** (Visitor) permission. The built-in **Everyone**
group is the most reliable option: it takes effect immediately and covers all
B2B guests who have accepted their invitation — no backend group sync needed.

If *Everyone* is not visible in the People Picker, enable the claim first:

Prefer `Install-PSResource` for module installation. On Windows PowerShell 5.1,
first update
[PowerShellGet / PSResourceGet](https://learn.microsoft.com/powershell/gallery/powershellget/install-powershellget?view=powershellget-3.x)
because `Install-PSResource` isn't available out of the box.

```powershell
# SharePoint Online Management Shell:
# Install once: Install-PSResource Microsoft.Online.SharePoint.PowerShell -Repository PSGallery -Scope CurrentUser
Set-SPOTenant -ShowEveryoneClaim $true

# PnP PowerShell (cross-platform; PowerShell 7+, also on Windows):
# Install once: Install-PSResource PnP.PowerShell -Repository PSGallery -Scope CurrentUser
# Register once: https://pnp.github.io/powershell/articles/registerapplication.html
Connect-PnPOnline -Url "https://<tenant>-admin.sharepoint.com" `
  -ClientId "<your-pnp-app-client-id>" -Interactive
Set-PnPTenant -ShowEveryoneClaim $true
```

Then add *Everyone* to the Visitors group via **Site Settings → People and
Groups → [Site] Visitors → New → Add Users → Everyone**.

> **Pitfall:** The similarly named *Everyone except external users* group
> explicitly **excludes** B2B guests — do not use it here.

**Alternative — static Microsoft Entra security group:** If your organisation uses an
automated guest invitation workflow (not implicit Teams/SharePoint invitations),
a static security group populated at invitation time is a viable alternative.
Microsoft Entra ID immediately reflects the new membership; SharePoint then resolves it
within seconds to a few minutes. Dynamic groups are slower because Microsoft Entra ID must
first re-evaluate its membership rule (up to 24 hours) before SharePoint sees
the change. The *Everyone* group remains preferred because its
`c:0(.s|true` claim is evaluated entirely within SharePoint's own
authentication layer — no Microsoft Entra group membership resolution required at all.
See the [deployment guide](docs/deployment.md#verify-guest-access-to-the-landing-page-site)
for full details, including an EasyLife 365 Collaboration tip for automated
guest lifecycle scenarios.

### 4. Deploy the Guest Sponsor API

The Microsoft Graph `/me/sponsors` API requires a directory role — impractical for
guests at scale. The included **Guest Sponsor API** calls Microsoft Graph with application
permissions instead (powered by a custom Azure Function).

Run the following command in PowerShell 7+ to deploy everything in one step:

```powershell
& ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.ps1').Content))
```

In Azure Cloud Shell, prefer this PowerShell entry point instead of
`install.sh`. If `azd` is missing, the wizard can install it into your Cloud
Shell user profile. With persisted Cloud Shell storage, that is usually a
one-time setup.

The wizard signs in with the Azure CLI via **device code** there — the Cloud
Shell identity cannot create the Entra app registration. Two prerequisites
follow from that:

- Device code sign-in must be permitted for your account. Some tenants block
  the device code flow via Conditional Access; ask your Entra administrator to
  allow it for this deployment, or run the installer outside Cloud Shell.
- If your admin account may only sign in from a Privileged Access Workstation,
  confirm the device code **on that PAW**. Any other device will be rejected at
  sign-in.

Your existing Cloud Shell login stays untouched; the run keeps its own
configuration directories.

On macOS or Linux, you can also start from a plain shell. This bootstraps
PowerShell when needed, then runs the same installer:

```bash
curl -fsSL https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.sh | bash
```

Without `-Version`, the installer deploys the newest published infra release.
Use `-Version vX.Y.Z` to pin a specific infra release, or `-Version main` only
when you explicitly want the mutable main-branch snapshot. When `-AppVersion`
is omitted, release-based `-Version` values also pin the Azure Function package
to the same release. Treat `-AppVersion` as an expert override, or use it
alongside `-Version main`.

`auto` selects device code in Azure Cloud Shell and on remote or headless
terminals, and a browser sign-in on local consoles. Use `-AzureLoginMode
browser` or `-AzureLoginMode device-code` to override that detection. Device
code sign-in must be permitted for your account, and PAW-restricted admin
accounts have to confirm the code on their PAW.

The wizard prepares the Microsoft Entra app registration, deploys the Azure-only hosting
stack, and then assigns Microsoft Graph permissions automatically. No local
repository clone is required.

In the web part property pane, open the **Guest Sponsor API** section and enter
the **Guest Sponsor API Base URL** and the **Guest Sponsor API Client ID
(App Registration)**. The Client ID comes from the App Registration named
**"Guest Sponsor Info - SharePoint Web Part Auth"** in your Microsoft Entra tenant.

> Full deployment details (Flex Consumption, Azure Maps,
> updating, security assessment, legacy options without the Guest Sponsor API):
> **[docs/deployment.md](docs/deployment.md)**

### 5. Add the web part to a page

Edit a modern page → add the *Guest Sponsor Info* web part.

## Security At A Glance

- Microsoft Graph application permissions stay server-side on the Azure
  Function's Managed Identity. The web part itself has no Microsoft Graph permissions.
- Azure App Service EasyAuth blocks unauthenticated requests before function
  code runs. In production, the function also validates tenant, audience, and
  the expected SharePoint client application.
- Follow-up presence and photo requests are limited to sponsor or manager IDs
  authorized for the current caller.
- The client bundle contains no secrets. Treat optional client-side
  configuration such as the Azure Maps subscription key as visible to page
  viewers.
- The main residual risk is the Azure Function's server-side Graph permission
  scope plus any admin configuration drift around EasyAuth, SharePoint access,
  CORS, and Azure RBAC.

> Full posture, residual risk, and hardening guidance:
> **[docs/security-assessment.md](docs/security-assessment.md)**
>
> Data handling and privacy scope:
> **[docs/privacy-policy.md](docs/privacy-policy.md)**

## Development

```bash
./scripts/bootstrap.sh         # install deps + create .env (then set SPFX_SERVE_TENANT_DOMAIN in .env)
./scripts/dev-webpart.sh       # SPFx dev server with hot-reload
```

To develop the Azure Function locally:

```bash
az login                       # authenticate for Microsoft Graph API access
./scripts/dev-function.sh      # build + start on http://localhost:7071
```

```bash
./scripts/build.sh             # CI-style clean build → .sppkg
./scripts/test.sh              # unit tests (Jest 29)
./scripts/lint.sh              # TypeScript · SCSS · Markdown
```

> Full development guide (scripts, testing scenarios, release workflow, CI,
> code conventions): **[docs/development.md](docs/development.md)**

## Further Documentation

| Document | Audience | Content |
|---|---|---|
| [docs/architecture-diagram.md](docs/architecture-diagram.md) | Everyone | Visual Mermaid diagram of the full system architecture |
| [docs/features.md](docs/features.md) | Everyone | Detailed feature descriptions and the problems they solve |
| [docs/deployment.md](docs/deployment.md) | Admins / Ops | Full deployment, guest access, Guest Sponsor API, security |
| [docs/security-assessment.md](docs/security-assessment.md) | Admins / Ops | Security posture, residual risk, and hardening guidance |
| [docs/development.md](docs/development.md) | Developers | Local setup, build, test, release, code conventions |
| [docs/architecture.md](docs/architecture.md) | Developers | Design decisions, data paths, known limitations |

## References

- [SharePoint Framework documentation](https://aka.ms/spfx)
- [Microsoft Graph – List sponsors](https://learn.microsoft.com/graph/api/user-list-sponsors)
- [Microsoft Entra B2B sponsors](https://learn.microsoft.com/azure/active-directory/external-identities/b2b-sponsors)
- [Use Microsoft Graph in your SPFx solution](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp)

## License

PolyForm Shield License 1.0.0 — see [LICENSE](LICENSE.md) for details.

Copyright © 2026 [Workoho GmbH](https://workoho.com/?utm_source=gsiw&utm_medium=docs&utm_campaign=repo&utm_content=readme)

Author: [Julian Pawlowski](https://github.com/jpawlowski)

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS
OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR
PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
