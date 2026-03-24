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

## Development Environment

A complete inventory of the toolchain — what is strictly required, what is
convenient to install locally, and what the devcontainer adds on top.

### Required (local setup without devcontainer)

Install these manually when not using the devcontainer.

| Tool | Purpose |
|---|---|
| **Node.js 22.x** | JavaScript runtime for the SPFx build, unit tests, and the Azure Function. Stay within the version range in `engines` in `package.json`. |
| **npm** | Package manager — bundled with Node.js, no separate installation needed. |
| **Git** | Version control. The pre-commit hooks (`husky`) enforce formatting and linting automatically before every commit. |

### Recommended for local setup (optional)

Only needed for specific workflows.

| Tool | Install | Purpose |
|---|---|---|
| **Azure CLI** (`az`) | [docs.microsoft.com/cli/azure](https://learn.microsoft.com/cli/azure/install-azure-cli) | Authenticate against Azure for local Function development and infra deployments. |
| **PowerShell** (`pwsh`) | [github.com/PowerShell](https://github.com/PowerShell/PowerShell#get-powershell) | Run the infra setup scripts (`setup-app-registration.ps1`, `setup-graph-permissions.ps1`) on macOS / Linux. |
| **GitHub CLI** (`gh`) | [cli.github.com](https://cli.github.com) | Manage PRs, issues, and CI runs from the terminal; used by `./scripts/release-notes.sh`. |

---

### Devcontainer — CLI tools

Everything below is pre-installed in the devcontainer and ready on the `PATH`
after a container build. No manual step required.

#### Search and navigation

| Command | Tool | Purpose |
|---|---|---|
| `rg` | **ripgrep** | Fast full-text search — faster than `grep`, respects `.gitignore`. Great for finding all usages of a symbol across the codebase. |
| `fd` | **fd-find** | Fast, readable alternative to `find` for locating files by name or pattern. |
| `bat` | **bat** | `cat` with syntax highlighting and line numbers — handy for quick file inspection in the terminal. |
| `fzf` | **fzf** | Interactive fuzzy finder for files, shell history, and any piped list. |

#### Data processing

| Command | Tool | Purpose |
|---|---|---|
| `jq` | **jq** | Parse and query JSON in the terminal — useful for inspecting Graph API responses or `package.json`. |
| `yq` | **yq** | Same as `jq` but also understands YAML — useful for config files. |

#### Shell script quality

| Command | Tool | Purpose |
|---|---|---|
| `shellcheck` | **ShellCheck** | Static analyzer: catches bugs and portability issues in shell scripts before they run. |
| `shfmt` | **shfmt** | Auto-formatter for shell scripts. Applied automatically at commit time via `lint-staged`. |

#### Git workflow

| Command | Tool | Purpose |
|---|---|---|
| `git diff` | **delta** | Renders `git diff` / `git show` with syntax highlighting and line numbers. Configured automatically as the default git pager. |

#### Azure and cloud

| Command | Tool | Purpose |
|---|---|---|
| `az` | **Azure CLI** | Manage Azure resources and authenticate for local Function development. |
| `func` | **Azure Functions Core Tools** | Start the Azure Function locally for development and debugging. Used by `./scripts/dev-function.sh`. |
| `gh` | **GitHub CLI** | Query PRs, issues, Actions runs, and releases; used by scripts and AI agents. |
| `pwsh` | **PowerShell** | Run `.ps1` infra setup scripts without needing a Windows machine. |

---

### Devcontainer — VS Code extensions

All extensions install automatically when the devcontainer starts.

#### AI assistants

| Extension | Purpose |
|---|---|
| **GitHub Copilot Chat** | Inline code completions and AI chat — integrated directly into the editor and terminal. |
| **Claude** (Anthropic) | Alternative AI coding assistant. |
| **ChatGPT** (OpenAI) | Alternative AI assistant. |

#### Code quality — linting and formatting

| Extension | Purpose |
|---|---|
| **ESLint** | Highlights TypeScript / JavaScript issues inline. Applies auto-fixes on save. |
| **Stylelint** | Highlights SCSS issues inline. |
| **Markdownlint** | Enforces Markdown style rules (heading levels, line length, blank lines). |
| **Prettier — Code Formatter** | Auto-formats JSON and JSONC files on save. |
| **shell-format** | Auto-formats shell scripts on save using `shfmt`. |
| **Error Lens** | Renders ESLint and TypeScript diagnostics directly on the affected line — no need to hover over the red underline. |

#### Git and GitHub

| Extension | Purpose |
|---|---|
| **GitLens** | Inline git blame, per-line authorship, and full file history without leaving the editor. |
| **GitHub Pull Requests** | Review, comment on, and merge PRs without switching to the browser. |
| **GitHub Actions** | Monitor CI workflow runs and see job logs in the sidebar. |
| **Conventional Commits** | GUI helper for composing commit messages that pass `commitlint` — avoids type / format errors. |

#### Azure and infrastructure

| Extension | Purpose |
|---|---|
| **Azure Developer CLI** | Deploy and manage the full app stack with `azd up` / `azd down`. |
| **Azure Functions** | Browse, deploy, and live-debug Azure Functions from VS Code. |
| **Azure Resources** | Browse subscriptions, resource groups, and resources in the sidebar. |
| **Azure Storage** | Browse blob containers and table storage directly in the editor. |
| **Bicep** | Syntax highlighting, validation (linting), and IntelliSense for `.bicep` infrastructure files. |
| **Azure CLI Tools** | Syntax highlighting and basic IntelliSense for `.azcli` script files. |
| **Docker** | Manage container images and Compose files from the sidebar. |
| **PowerShell** | IntelliSense and integrated debugging for `.ps1` scripts. |

#### SharePoint and general

| Extension | Purpose |
|---|---|
| **Deploy to SharePoint Online** | Upload `.sppkg` to the SharePoint App Catalog directly from VS Code — no CLI needed. |
| **EditorConfig** | Reads `.editorconfig` and enforces consistent line endings and indentation across every file type. |

---

### Devcontainer — GitHub MCP server

The devcontainer ships a pre-configured
[GitHub MCP server](https://github.com/modelcontextprotocol/servers/tree/main/src/github).
It gives AI coding agents (Copilot, Claude, etc.) structured access to the
GitHub API — issues, PRs, Actions runs, and releases — as first-class tools,
without the agent having to parse `gh` CLI output. Requires `GITHUB_TOKEN` to
be set (see `containerEnv` in `.devcontainer/devcontainer.json`).

---

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
| `npm run lint` | Run all linters (TypeScript · SCSS · Markdown · JSON · shell) |
| `npm run fix` | Auto-fix all formatting issues (ESLint · Stylelint · Prettier · shfmt · Markdownlint) |

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
