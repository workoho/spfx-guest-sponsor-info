# Development Guide

Local development setup, build commands, testing, and release workflow.

For deployment and administration, see [deployment.md](deployment.md).
For architecture decisions and internals, see [architecture.md](architecture.md).
For a visual system overview, see [architecture-diagram.md](architecture-diagram.md).

## Prerequisites

| Requirement | Version |
|---|---|
| Node.js | `>=22.14.0 <23.0.0` (see `engines` in `package.json`) |
| npm | Bundled with Node.js |

## Quick Start

```bash
./scripts/bootstrap.sh         # install deps + create .env (then set SPFX_SERVE_TENANT_DOMAIN in .env)
./scripts/dev-webpart.sh       # starts SPFx dev server with hot-reload
```

The dev server serves bundled code to the **hosted workbench** on your
SharePoint Online tenant. Accept the certificate warning at
`https://localhost:4321` once per browser session, then open the hosted
workbench URL printed on startup.

## Scripts

| Script | Purpose |
|---|---|
| `./scripts/bootstrap.sh` | Install deps + create `.env` (run once after cloning) |
| `./scripts/reset.sh` | Wipe build outputs + node_modules, then re-bootstrap |
| `./scripts/dev-webpart.sh` | Start SPFx web part dev server |
| `./scripts/dev-function.sh` | Start Azure Function locally |
| `./scripts/test.sh` | Run tests |
| `./scripts/lint.sh` | Run all linters |
| `./scripts/lint-fix.sh` | Auto-fix lint issues locally |
| `./scripts/build.sh` | CI-style clean build → `.sppkg` |
| `./scripts/release-notes.sh` | Preview release notes locally |
| `./scripts/set-version.sh v1.x.y` | Stamp a release version |
| `./scripts/set-version.sh v1.x.y --commit` | Stamp, commit, and tag |
| `./scripts/upgrade-spfx.sh 1.x.y` | Guided SPFx upgrade |

### npm commands

| Command | Description |
|---|---|
| `npm run build` | Full production build + unit tests + packaging |
| `npm test` | Compile and run unit tests |
| `npm start` | Start dev server (hot-reload, hosted workbench) |
| `npm run clean` | Delete all build output |
| `npm run lint` | Run all linters (TypeScript · SCSS · Markdown) |
| `npm run fix` | Auto-fix formatting issues |

## Testing

Tests are written with **Jest 29** and `react-dom/test-utils`.

```bash
npm test
```

Coverage output is written to `jest-output/coverage/`.

### Testing scenarios

- **Hosted workbench as member:** `SPFX_SERVE_TENANT_DOMAIN` in `.env` (or set on host OS) +
  `./scripts/dev-webpart.sh`. Verifies the non-guest path.
- **Hosted workbench as guest:** Requires a second M365 tenant where your
  account is a guest with sponsors assigned, API permissions consented, and
  `.sppkg` deployed or localhost script loading enabled.
- **Demo mode:** Property pane toggle shows two fictitious sponsors without
  Graph calls. Development and visual review only — disable before production.

See [architecture.md](architecture.md#development-testing) for more details.

## Azure Function Development

The Azure Function (`azure-function/`) acts as a proxy between the web part
and Microsoft Graph. It can be developed and tested locally.

### Quick start

```bash
az login                       # authenticate for Graph API access
./scripts/dev-function.sh      # build + start on http://localhost:7071
```

On first run the script copies `local.settings.json.example` to
`local.settings.json` (git-ignored). Edit it with your tenant values:

| Variable | Value |
|---|---|
| `TENANT_ID` | Your Entra tenant ID (GUID) |
| `ALLOWED_AUDIENCE` | Client ID (GUID) of the Function's App Registration |
| `CORS_ALLOWED_ORIGIN` | `https://<tenant>.sharepoint.com` |

### Testing the function standalone

[EasyAuth](https://learn.microsoft.com/azure/app-service/overview-authentication-authorization) is not
active locally. Pass the caller OID via the dev header:

```bash
curl http://localhost:7071/api/getGuestSponsors \
  -H "X-Dev-User-OID: <guest-user-oid>"
```

The `X-Dev-User-OID` header is only accepted when `NODE_ENV` is not
`production` (the dev script sets it to `development`).

The function uses `DefaultAzureCredential` which resolves to:

- **In Azure** — Managed Identity (zero-config).
- **Locally** — Azure CLI (`az login`), or a service principal when
  `AZURE_CLIENT_ID` / `AZURE_TENANT_ID` / `AZURE_CLIENT_SECRET` are set.

### Mock mode (no Azure credentials needed)

Set `MOCK_MODE=true` in `azure-function/local.settings.json` to return
demo sponsor and presence data without any Graph API calls. This is useful
for:

- Testing HTTP routing, CORS, rate limiting, and error handling
- UI development against a predictable response
- CI environments or dev containers without Azure credentials

Mock mode is blocked in production (`NODE_ENV=production`).

### Connecting the web part to the local function

To test the full web part → function → Graph pipeline locally:

1. Start the function: `./scripts/dev-function.sh`
2. Start the web part: `./scripts/dev-webpart.sh`
3. In the web part property pane, set **Function URL** to `localhost:7071`.

The web part constructs `https://localhost:7071/api/getGuestSponsors`
(always HTTPS). To make this work:

- **Dev container / Codespace** — VS Code port forwarding automatically
  provides an HTTPS endpoint for port 7071. No extra setup needed.
- **Local machine** — start the function with HTTPS:
  `./scripts/dev-function.sh --useHttps`
  Accept the self-signed certificate in your browser at
  `https://localhost:7071` once.

### Watch mode

For automatic TypeScript recompilation during development:

```bash
cd azure-function && npm run watch   # in one terminal
cd azure-function && func start      # in another terminal
```

The VS Code task "npm watch (functions)" also provides this.

## Build from Source

```bash
./scripts/build.sh    # clean install + compile + test + bundle + package
```

The packaged solution is written to
`sharepoint/solution/guest-sponsor-info.sppkg`.

## Publishing a Release

Releases are created by pushing a SemVer tag:

```bash
./scripts/set-version.sh v1.2.3 --commit   # stamp version, commit, tag
git push && git push --tags                 # triggers release workflow
```

The workflow automatically:

1. Generates release notes from Conventional Commit history
   (via [git-cliff](https://git-cliff.org)).
2. Builds the production `.sppkg`.
3. Creates a GitHub Release with the notes and `.sppkg` attached.

For the **first release**, this works even with a single commit and no prior
tags. Preview release notes before tagging:

```bash
./scripts/release-notes.sh
```

## Continuous Integration

| Workflow | Trigger | What it does |
|---|---|---|
| `ci.yml` | Push / PR to `main` | Build · test · upload coverage |
| `release.yml` | `v*` SemVer tag | Bump version · build · create GitHub Release + `.sppkg` |

## Code Conventions

- All code comments and documentation in English.
- Styles in `GuestSponsorInfo.module.scss` (CSS Modules, camelCase class
  names).
- Locale strings follow the SPFx AMD `define()` pattern in `loc/*.js`; add new
  keys to **all** locale files.
- No bundled placeholder images — use live profile photos from Graph; fall back
  to initials.
- Graph permissions: `User.Read` and `User.ReadBasic.All` only. Do not
  introduce `User.Read.All` or broader scopes on the web part side.

## Stack Constraints

- **SPFx** — do not upgrade unless explicitly asked. Use
  `scripts/upgrade-spfx.sh` when needed.
- **React** — pinned; do not change.
- **Build tool** — Heft (no Gulp). Use `npm` scripts, never raw `npx heft`.
- **Never run `npm audit fix --force`** — it would downgrade SPFx build-rig
  packages and break the build.
- **Never run `npm update`** on `@microsoft/sp-*`, `@rushstack/*`, `react`, or
  `@fluentui/react` — these are managed as a coordinated set via
  `scripts/upgrade-spfx.sh`.

## Key Files

| File | Purpose |
|---|---|
| `src/webparts/guestSponsorInfo/services/SponsorService.ts` | All Graph logic |
| `src/webparts/guestSponsorInfo/components/GuestSponsorInfo.tsx` | Main component |
| `src/webparts/guestSponsorInfo/components/SponsorCard.tsx` | Individual card |
| `docs/architecture.md` | Design decisions and limitations |
