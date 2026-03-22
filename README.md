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
| `guest-sponsor-info.sppkg` | [Julian Pawlowski](https://github.com/jpawlowski) |

## Prerequisites

| Requirement | Detail |
|---|---|
| SharePoint Online | Modern team or communication site |
| Microsoft Entra | Guest accounts with one or more sponsors assigned |
| Microsoft Graph permissions | `User.Read` · `User.ReadBasic.All` (· `Presence.Read.All` optional) |

## Features

- **Sponsor cards** — photo (or initials + deterministic colour) · name · job
  title
- **Contact overlay** — email · phone · office location · department on
  hover/focus
- **Guest-only rendering** — renders nothing for member users; always shows a
  placeholder in edit mode
- **Unavailable-sponsor handling** — deleted or disabled sponsors are hidden; a
  message is shown when all sponsors are gone
- **Multilingual** — English · German · French · Spanish · Italian + 9 more
- **Theme-aware** — honours the site theme
- **Least-privilege** — `User.ReadBasic.All` instead of `User.Read.All`

## Minimal Path to Awesome

> [PnP community convention](https://aka.ms/m365pnp) — shortest way to get
> the web part running.

### 1. Deploy the web part

1. Download the latest `guest-sponsor-info.sppkg` from
   [Releases](../../releases).
2. Upload it to your SharePoint **App Catalog** and add the app to each target
   site (**Site Contents → New → App → Guest Sponsor Info**).
3. Approve any pending Graph permissions in **SharePoint Admin Center →
   Advanced → API access** (often already pre-consented — the queue will
   simply be empty).

### 2. Enable the SharePoint Public CDN

Guest users cannot reach assets hosted in the App Catalog by default. The
simplest fix is to enable the Public CDN:

```powershell
Connect-PnPOnline -Url "https://<tenant>-admin.sharepoint.com" -Interactive
Set-PnPTenantCdnEnabled -CdnType Public -Enable $true
Add-PnPTenantCdnOrigin -CdnType Public -OriginUrl "*/CLIENTSIDEASSETS"
```

Propagation takes ≈ 15–30 min. After that, asset URLs are rewritten
automatically — no redeployment needed.

> Cannot use the Public CDN? See
> [docs/deployment.md](docs/deployment.md#option-b--app-catalog-permissions-alternative)
> for the App Catalog permissions alternative.

### 3. Verify external sharing

External sharing must be enabled at the tenant level and on each site where
the web part is placed: **SharePoint Admin Center → Policies → Sharing** → at
least *Existing guests only*.

### 4. Deploy the Sponsor API (Azure Function)

The Graph `/me/sponsors` API requires a directory role — impractical for
guests at scale. The included Azure Function proxy calls Graph with
application permissions instead.

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
> updating, security assessment, legacy options without Azure Function):
> **[docs/deployment.md](docs/deployment.md)**

### 5. Add the web part to a page

Edit a modern page → add the *Guest Sponsor Info* web part.

## Development

```bash
npm install
cp .env.example .env           # set SPFX_TENANT=<tenant>.sharepoint.com
./scripts/dev-webpart.sh       # SPFx dev server with hot-reload
```

To develop the Azure Function locally:

```bash
az login                       # authenticate for Graph API access
./scripts/dev-function.sh      # build + start on http://localhost:7071
```

```bash
npm run build                  # full production build → .sppkg
npm test                       # unit tests (Jest 29)
npm run lint                   # TypeScript · SCSS · Markdown
```

> Full development guide (scripts, testing scenarios, release workflow, CI,
> code conventions): **[docs/development.md](docs/development.md)**

## Further Documentation

| Document | Audience | Content |
|---|---|---|
| [docs/deployment.md](docs/deployment.md) | Admins / Ops | Full deployment, guest access, Azure Function, security |
| [docs/development.md](docs/development.md) | Developers | Local setup, build, test, release, code conventions |
| [docs/architecture.md](docs/architecture.md) | Developers | Design decisions, data paths, known limitations |

## References

- [SharePoint Framework documentation](https://aka.ms/spfx)
- [Microsoft Graph – List sponsors](https://learn.microsoft.com/graph/api/user-list-sponsors)
- [Microsoft Entra B2B sponsors](https://learn.microsoft.com/azure/active-directory/external-identities/b2b-sponsors)
- [Use Microsoft Graph in your SPFx solution](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp)

## License

MIT — see [LICENSE](LICENSE) for details.

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS
OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR
PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
