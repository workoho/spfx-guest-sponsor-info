# GitHub Copilot Instructions

This repository contains a SharePoint Online web part built with SharePoint Framework.
Authoritative versions are in `package.json` (`engines`, `dependencies`, `devDependencies`).

> **For AI coding agents:** See [AGENTS.md](../AGENTS.md) for recommended tools and workflows.

## After every code change

Validate changes before considering a task done, but use the right level of
validation for the situation. Running the full lint suite after every small edit
wastes time — use targeted checks during development and the full suite only
before committing.

### Targeted lint (after each edit — fast, 2–5 s)

Run **only** the linter that matches the files you changed:

| Changed files | Command |
| --- | --- |
| `src/**/*.{ts,tsx}` | `npm run lint:ts` |
| `azure-function/src/**/*.ts` | `npm run lint:ts:func` |
| `**/*.md` | `npm run lint:md` |
| `.github/**/*.yml`, `azure.yaml`, `website/**/*.yml` | `npm run lint:yml` |
| `azure-function/infra/**/*.bicep` | `npm run lint:bicep` |
| `scripts/*.sh` | `npm run lint:sh` |
| `src/**/loc/*.js` | `npm run lint:loc` |

If you changed files spanning multiple types, run the relevant subset — not the
full suite.

### Full validation (before committing)

1. **Fix** — `npm run fix` (auto-correct formatting: ESLint, Prettier, shfmt, Markdownlint)
2. **Lint** — `npm run lint` (full suite — catches anything `fix` could not resolve)
3. **Test** — `npm test` only when you changed logic, components, or services.
   Skip for pure documentation, config, locale, or style-only changes.
4. **Build** — `npm run build` only when the packaging artifact (`.sppkg`) is relevant,
   e.g. before a release. Not needed for regular development changes.

The pre-commit hook (`lint-staged`) already runs targeted fix + lint on staged
files automatically (ESLint, Prettier, Markdownlint, shfmt — not the full
lint suite). If you followed the targeted-lint workflow above, the commit hook
will catch any remaining issues in staged files.

For interactive development use `npm start` (hosted workbench with hot-reload; requires `SPFX_SERVE_TENANT_DOMAIN` —
set in `.env` or as a host OS env var, see `.devcontainer/devcontainer.json`).
For a CI-style clean build from scratch use `./scripts/build.sh` (runs `npm ci` first).

If any lint errors or test failures appear after your changes, fix them before finishing.
Do not suppress linter rules or skip tests to make the pipeline green.

## Key scripts

| Script                                     | Purpose                                               |
| ------------------------------------------ | ----------------------------------------------------- |
| `./scripts/bootstrap.sh`                   | Install deps + create `.env` (run once after cloning) |
| `./scripts/reset.sh`                       | Wipe build outputs + node_modules, then re-bootstrap  |
| `./scripts/dev-webpart.sh`                 | Start SPFx web part dev server                        |
| `./scripts/dev-function.sh`                | Start Azure Function locally                          |
| `./scripts/test.sh`                        | Run tests                                             |
| `./scripts/lint.sh`                        | Run all linters                                       |
| `./scripts/lint-fix.sh`                    | Auto-fix lint issues locally                          |
| `./scripts/build.sh`                       | CI-style clean build → `.sppkg`                       |
| `./scripts/release-notes.sh`               | Preview release notes locally                         |
| `./scripts/set-version.sh v1.x.y`          | Stamp a release version                               |
| `./scripts/set-version.sh v1.x.y --commit` | Stamp, commit, and tag a release                      |
| `./scripts/upgrade-spfx.sh 1.x.y`          | Guided SPFx upgrade                                   |

The release workflow is documented in `docs/development.md` → "Publishing a Release".

## Stack constraints

- SPFx — do not upgrade unless explicitly asked. Use `scripts/upgrade-spfx.sh` when needed.
- Node.js — stay within the range defined in the `engines` field of `package.json`.
- React — the version is pinned; do not change it.
- Build tool: Heft (no Gulp). Use `npm test`, `npm run build`, `npm start` — never raw `npx heft`
  unless diagnosing a build problem.
- Before adding any package, verify compatibility with the current SPFx version.
- **Never run `npm audit fix --force`** — it would downgrade SPFx build-rig packages and break
  the build. Audit warnings in transitive SPFx dependencies cannot be fixed independently.
- **Never run `npm update`** on `@microsoft/sp-*`, `@rushstack/*`, `react`, or `@types/react`.
  These are managed as a coordinated set via `scripts/upgrade-spfx.sh`.

## Feature behaviour

- The web part shows the Microsoft Entra **sponsors** of the currently signed-in guest user.
- In **view mode**: render nothing for non-guest users; render sponsor cards for guests.
- In **edit mode**: always show a lightweight text placeholder — no Graph calls, no photos.
- Guest detection: `#EXT#` marker in `pageContext.user.loginName`.
- The web part has no Microsoft Graph permissions of its own. All data fetching goes through
  the Azure Function via `AadHttpClient`. Do not add `webApiPermissionRequests` entries or
  introduce direct Graph calls.

## Code style

- All code comments and documentation in English. User-facing chat may stay in German.
- No bundled placeholder images. Use live profile photos from Graph; fall back to initials.
- Styles use `makeStyles` + `tokens` from `@fluentui/react-components` (Griffel) for all component-level
  styles. Do not add CSS/SCSS module files.
- Locale strings follow the SPFx AMD `define()` pattern in `loc/*.js`; add new keys to all 17 locale files
  (cs-cz, da-dk, de-de, en-us, es-es, fi-fi, fr-fr, hu-hu, it-it, ja-jp, nb-no, nl-nl, pl-pl, pt-pt, ro-ro, sv-se, zh-cn).

## Shell scripts (`scripts/*.sh`)

- Every script must start with `set -euo pipefail` and
  `cd "$(dirname "${BASH_SOURCE[0]}")/.."`.
- Source `scripts/colors.sh` for all colour output — never copy the colour-detection
  block inline. Colour variables: `C_RED` `C_GRN` `C_YLW` `C_CYN` `C_BLD` `C_DIM` `C_RST`.
- Use the callout box functions from `scripts/colors.sh` for developer-facing messages
  that must not be missed: `hint` (cyan — tips), `next_steps` (green — what to do next),
  `important` (yellow — critical action items). Pass each line as a separate argument.
- Callout boxes are for **interactive** scripts only (ones a developer runs in a terminal).
  Automated hooks (e.g. `azd` pre/post-provision) must use plain `echo` — no `colors.sh`
  dependency, no visual formatting. Prioritise efficiency and robustness there.
- After every change run `npm run lint:sh` (`shellcheck -x`). Fix all warnings — do not
  suppress them with `# shellcheck disable` without a comment explaining why.
- Bash parameter expansions and non-obvious constructs must have an inline comment.
- Scripts that perform side effects (file writes, git ops) should support a dry-run mode
  via a `maybe()` helper that prints `[dry-run] <cmd>` instead of executing.

## Key files

- `src/webparts/guestSponsorInfo/services/SponsorService.ts` — all proxy call logic (Azure Function API)
- `src/webparts/guestSponsorInfo/components/GuestSponsorInfo.tsx` — main component
- `src/webparts/guestSponsorInfo/components/SponsorCard.tsx` — individual card
- `docs/architecture.md` — design decisions and known limitations
