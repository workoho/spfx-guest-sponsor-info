# Guest Sponsor Info

A SharePoint Online web part for landing pages in Microsoft Entra **resource
tenants** that shows the sponsors of the currently signed-in **guest user**.

Each sponsor is rendered as a card with a live profile photo (or initials
fallback), name, and job title. Hovering or focusing a card reveals contact
details. The layout matches the built-in SharePoint People web part.

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [SharePoint Online](https://www.microsoft.com/microsoft-365)
- [Microsoft Entra ID – External Identities (B2B)](https://learn.microsoft.com/azure/active-directory/external-identities/)

## Solution

| Solution | Author(s) |
|---|---|
| `guest-sponsor-info.sppkg` | [Workoho GmbH](https://github.com/workoho) ([Julian Pawlowski](https://github.com/jpawlowski)) |

## Prerequisites

| Requirement | Detail |
|---|---|
| SharePoint Online | Modern team or communication site |
| Microsoft Entra | Guest accounts with one or more sponsors assigned |
| Microsoft Graph permissions | `User.Read` · `User.ReadBasic.All` (· `Presence.Read.All` optional) |

## Features

Works out of the box with any standard Microsoft 365 environment — no
third-party tools or paid add-ons required. Every feature below relies solely
on Microsoft Graph, SharePoint Framework, and the included optional Azure
Function. [EasyLife 365 Collaboration](https://easylife365.cloud/) pairs
naturally as a companion: it automates sponsor assignments and the full guest
lifecycle so the right information stays accurate over time — solid on its
own, stronger together. [Workoho](https://www.workoho.com/), the author of
this web part, is a Platinum EasyLife 365 partner and happy to advise.

- **"Who is my contact here?"** — guests see their sponsor's photo, name, and
  job title directly on the SharePoint landing page — no searching, no asking
  around
- **Familiar look and feel** — the contact card is modelled after the Microsoft
  Teams profile card and the SharePoint People web part, so it feels native
  rather than like a custom add-on
- **Everything needed to reach out** — email, phone, and Teams Chat / Call
  buttons in one click; office address with map preview (Azure Maps) or a link
  to the map provider of your choice (Google, Apple, Bing, OpenStreetMap, HERE)
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
  [EasyLife 365](https://easylife365.cloud/) do), the web part honours that
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
- **14 languages** — including an informal salutation mode (`du`/`tu`) for
  German, French, Spanish, Italian, and Dutch

> **Tip:** Want to automate who gets assigned as a sponsor — and keep those
> assignments current over time? [EasyLife 365 Collaboration](https://easylife365.cloud/)
> handles the full lifecycle of Microsoft 365 collaboration workspaces — Teams, SharePoint
> team sites, Viva Engage communities, and more — including guest onboarding and sponsor
> management. This web part then takes care of the guest-facing experience.
> [Workoho](https://www.workoho.com/), the author of this web part, is a Platinum sales and
> implementation partner of EasyLife 365 and happy to help.
>
> Full feature descriptions, design decisions, and the problem this solves:
> **[docs/features.md](docs/features.md)**

## Minimal Path to Awesome

> For a visual overview of all setup steps and required admin roles, see the
> [Setup diagram](docs/architecture-diagram.md#setup--two-admin-roles-recommended-path).

### 1. Deploy the web part

The web part's bundle is hosted in a **Site Collection App Catalog** directly
on the guest landing page site itself. Because guest users already need read
access to that site, no CDN configuration or extra permissions on the global
App Catalog are required.

**Enable the Site Collection App Catalog** (once, as **SharePoint
Administrator**; no GUI available — PowerShell required).

On Windows, use the **SharePoint Online Management Shell** — no additional
setup needed:

```powershell
Connect-SPOService -Url "https://<tenant>-admin.sharepoint.com"
Add-SPOSiteCollectionAppCatalog -Site "https://<tenant>.sharepoint.com/sites/<landing-site>"
```

On macOS/Linux, use [PnP PowerShell](https://pnp.github.io/powershell/)
(requires your own Entra app registration; pass its Client ID via `-ClientId`):

```powershell
Connect-PnPOnline -Url "https://<tenant>-admin.sharepoint.com" `
    -ClientId "<your-pnp-app-client-id>" -Interactive
Add-PnPSiteCollectionAppCatalog -Site "https://<tenant>.sharepoint.com/sites/<landing-site>"
```

**Upload and install the package:**

1. Download the latest `guest-sponsor-info.sppkg` from
   [Releases](../../releases).
2. Navigate to the landing page site → **Site Contents** →
   **Apps for SharePoint** and upload the `.sppkg` file.
   *(The Site Collection App Catalog is also accessible directly at
   `https://<tenant>.sharepoint.com/sites/<landing-site>/AppCatalog/`.)*
3. The web part becomes available on all pages in this site collection
   immediately — no additional "Add App" step is required.
4. The required Microsoft Graph permissions (`User.Read`, `User.ReadBasic.All`,
   `Presence.Read.All`) are pre-authorized by Microsoft for SharePoint Online —
   the **SharePoint Admin Center → Advanced → API access** queue will simply
   be empty and no manual consent is needed.

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

```powershell
# SharePoint Online Management Shell:
Set-SPOTenant -ShowEveryoneClaim $true

# PnP PowerShell (cross-platform):
Set-PnPTenant -ShowEveryoneClaim $true
```

Then add *Everyone* to the Visitors group via **Site Settings → People and
Groups → [Site] Visitors → New → Add Users → Everyone**.

> **Pitfall:** The similarly named *Everyone except external users* group
> explicitly **excludes** B2B guests — do not use it here.

**Alternative — static Entra security group:** If your organisation uses an
automated guest invitation workflow (not implicit Teams/SharePoint invitations),
a static security group populated at invitation time is a viable alternative.
Entra ID immediately reflects the new membership; SharePoint then resolves it
within seconds to a few minutes. Dynamic groups are slower because Entra must
first re-evaluate its membership rule (up to 24 hours) before SharePoint sees
the change. The *Everyone* group remains preferred because its
`c:0(.s|true` claim is evaluated entirely within SharePoint's own
authentication layer — no Entra group membership resolution required at all.
See the [deployment guide](docs/deployment.md#verify-guest-access-to-the-landing-page-site)
for full details, including an EasyLife 365 Collaboration tip for automated
guest lifecycle scenarios.

### 4. Deploy the Guest Sponsor API

The Graph `/me/sponsors` API requires a directory role — impractical for
guests at scale. The included **Guest Sponsor API** calls Graph with application
permissions instead (powered by a custom Azure Function).

```powershell
# Create the App Registration
./azure-function/infra/setup-app-registration.ps1 -TenantId "<tenant-id>"
```

Then deploy to Azure:

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2Fjpawlowski%2Fspfx-guest-sponsor-info%2Fmain%2Fazure-function%2Finfra%2Fazuredeploy.json)

After deployment, grant Graph permissions and configure the web part:

```powershell
./azure-function/infra/setup-graph-permissions.ps1 `
  -ManagedIdentityObjectId "<oid-from-deployment-output>" `
  -TenantId "<tenant-id>" `
  -FunctionAppClientId "<client-id>"
```

In the web part property pane, enter the **Azure Function Base URL** and the
**Sponsor API Client ID**.

> Full deployment details (Flex Consumption, Deployment Stacks, Azure Maps,
> updating, security assessment, legacy options without the Guest Sponsor API):
> **[docs/deployment.md](docs/deployment.md)**

### 5. Add the web part to a page

Edit a modern page → add the *Guest Sponsor Info* web part.

## Development

```bash
./scripts/bootstrap.sh         # install deps + create .env (then set SPFX_SERVE_TENANT_DOMAIN in .env)
./scripts/dev-webpart.sh       # SPFx dev server with hot-reload
```

To develop the Azure Function locally:

```bash
az login                       # authenticate for Graph API access
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
| [docs/development.md](docs/development.md) | Developers | Local setup, build, test, release, code conventions |
| [docs/architecture.md](docs/architecture.md) | Developers | Design decisions, data paths, known limitations |

## References

- [SharePoint Framework documentation](https://aka.ms/spfx)
- [Microsoft Graph – List sponsors](https://learn.microsoft.com/graph/api/user-list-sponsors)
- [Microsoft Entra B2B sponsors](https://learn.microsoft.com/azure/active-directory/external-identities/b2b-sponsors)
- [Use Microsoft Graph in your SPFx solution](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp)

## License

AGPL-3.0-only — see [LICENSE](LICENSE) for details.

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS
OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR
PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
