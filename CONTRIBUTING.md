# Contributing to Guest Sponsor Info

Thank you for your interest in contributing! This guide covers everything you
need to know to get started, submit changes, and help keep the project healthy.

## Table of Contents

- [Ways to Contribute](#ways-to-contribute)
- [Licensing](#licensing)
- [Development Setup](#development-setup)
- [Branching and Committing](#branching-and-committing)
- [Code Style and Linting](#code-style-and-linting)
- [Testing](#testing)
- [Locale Strings (i18n)](#locale-strings-i18n)
- [Stack Constraints](#stack-constraints)
- [Pull Request Process](#pull-request-process)
- [Reporting Bugs](#reporting-bugs)
- [Requesting Features](#requesting-features)
- [Security Vulnerabilities](#security-vulnerabilities)

---

## Ways to Contribute

- **Bug reports** — open an issue with a clear description and reproduction
  steps.
- **Feature requests** — open an issue describing the use case and expected
  behaviour.
- **Pull requests** — fix a bug, add a feature, improve documentation, or
  tighten up a locale string.
- **Documentation** — corrections and clarifications in `docs/` or `README.md`
  are always welcome.

---

## Licensing

This project is licensed under the
[PolyForm Shield 1.0.0](LICENSE.md) license. By submitting a pull request, you
agree that your contribution will be distributed under the same license.

## Upstream Contributions

If you improve Guest Sponsor Info in a way that is likely to help other
Microsoft 365 tenants, please consider contributing those changes back upstream
instead of keeping a private fork indefinitely. We are most interested in
changes that solve real, repeatable scenarios: they do not need to apply to
every tenant, but they should be relevant beyond a single organization's unusual
edge case. Bug fixes, accessibility improvements, security hardening,
tenant-agnostic features, documentation updates, and locale corrections are
especially valuable to the wider community.

Tenant-specific configuration, private branding, internal deployment scripts,
or changes that reveal confidential implementation details do not need to be
shared.

---

## Development Setup

The fastest way to get a working environment is to open the repository in the
provided **devcontainer** (VS Code or GitHub Codespaces). All required tools —
Node.js, Azure CLI, Azure Functions Core Tools, PowerShell, ShellCheck,
shfmt, ripgrep, and more — are pre-installed automatically.

### Quickstart (devcontainer)

```bash
# The container runs post-create.sh automatically.
# After it finishes, set your tenant domain:
echo "SPFX_SERVE_TENANT_DOMAIN=<your-tenant>.sharepoint.com" >> .env

# Start the web part dev server (hot-reload):
./scripts/dev-webpart.sh

# Start the Azure Function locally (separate terminal):
./scripts/dev-function.sh
```

### Manual setup (without devcontainer)

1. Install **Node.js 22.x** (stay within the `engines` range in `package.json`).
2. Run the bootstrap script:

   ```bash
   ./scripts/bootstrap.sh
   ```

3. Set `SPFX_SERVE_TENANT_DOMAIN` in `.env`.

For the full prerequisites list, architecture details, and a breakdown of
every dev-environment tool, see [docs/development.md](docs/development.md).

---

## Branching and Committing

### Branches

Work on a **feature branch** forked from `main`:

```bash
git checkout -b feat/my-feature
```

Aim for focused, reviewable branches. One logical change per PR.

### Commit messages

This repository enforces
[Conventional Commits](https://www.conventionalcommits.org/) via
`commitlint`. The pre-commit hook will **reject** messages that do not comply.

**Hard limits:**

| Part | Rule |
| --- | --- |
| Header (first line) | ≤ 100 characters |
| Each body / footer line | ≤ 100 characters (wrap manually) |
| Subject case | lower-case, no trailing period |
| Type | one of the allowed types below |

**Allowed types:**
`fix` · `feat` · `build` · `chore` · `ci` · `docs` · `perf` · `refactor` ·
`revert` · `style` · `test`

**Examples:**

```text
feat(sponsor): show job title below sponsor display name
fix(graph): handle missing profile photo gracefully
chore(deps): bump eslint from 8.x to 9.x
docs: add deployment guide for sovereign cloud tenants
```

Use the **Conventional Commits** VS Code extension (pre-installed in the
devcontainer) to compose messages interactively if you are unsure.

The full ruleset and decision tree for choosing the right type and scope are
in [.github/instructions/commit-message.instructions.md](.github/instructions/commit-message.instructions.md).

---

## Code Style and Linting

All formatting and lint issues are auto-fixed at commit time by `lint-staged`.
Run the targeted lint command that matches the files you changed during
development — do **not** run the full suite after every small edit:

| Changed files | Command |
| --- | --- |
| `src/**/*.{ts,tsx}` | `npm run lint:ts` |
| `azure-function/src/**/*.ts` | `npm run lint:ts:func` |
| `**/*.md` | `npm run lint:md` |
| `.github/**/*.yml`, `azure.yaml` | `npm run lint:yml` |
| `scripts/*.sh` | `npm run lint:sh` |
| `src/**/loc/*.js` | `npm run lint:loc` |

Before opening a PR, run the full validation suite:

```bash
./scripts/lint-fix.sh   # auto-correct formatting
./scripts/lint.sh       # catch anything fix couldn't resolve
./scripts/test.sh       # only when logic, components, or services changed
```

### React / Fluent UI

This project uses **Fluent UI v9** (`@fluentui/react-components`) exclusively:

- Use `makeStyles` + `tokens` for all component-level styles.
- Use `mergeClasses()` for conditional class composition.
- Do **not** add CSS / SCSS module files.
- Do **not** import from `@fluentui/react` (v8 — migration is complete).
- Do **not** hardcode colour values — always use `tokens`.

See [.github/instructions/fluent-ui.instructions.md](.github/instructions/fluent-ui.instructions.md)
for the full rules and v8 → v9 component mapping.

### Shell scripts

Every script in `scripts/*.sh` must:

- Start with `set -euo pipefail` and the canonical `cd "$(dirname …)"` line.
- Source `scripts/colors.sh` for colour output.
- Be idempotent (safe to run multiple times).
- Pass `npm run lint:sh` (ShellCheck) with no warnings.

---

## Testing

```bash
./scripts/test.sh   # run Jest + coverage
```

Add or update tests whenever you change logic, services, or components.
Skip tests for documentation, config, locale, or style-only changes.

Tests live alongside the source files (`*.test.ts` / `*.test.tsx`).
Coverage output is written to `jest-output/`.

---

## Locale Strings (i18n)

Locale strings use the SPFx AMD `define()` pattern in `src/webparts/guestSponsorInfo/loc/*.js`.

**Every new or changed key must be added to all 17 locale files:**

`cs-cz` · `da-dk` · `de-de` · `en-us` · `es-es` · `fi-fi` · `fr-fr` ·
`hu-hu` · `it-it` · `ja-jp` · `nb-no` · `nl-nl` · `pl-pl` · `pt-pt` ·
`ro-ro` · `sv-se` · `zh-cn`

Use UTF-8 native characters — no `\uXXXX` escape sequences. When you add a
key, use a reasonable English string for locales you cannot translate and note
it in your PR description so a native speaker can follow up.

Run `npm run lint:loc` after touching any locale file to catch syntax errors.

---

## Stack Constraints

Please respect these constraints — violations will be caught in review:

- **SPFx version** — do not upgrade unless explicitly requested. Use
  `./scripts/upgrade-spfx.sh` when needed.
- **Node.js** — stay within the `engines` range in `package.json`
  (`>=22.14.0 <23.0.0`).
- **React** — pinned at 17.0.1; do not change it.
- **`@microsoft/sp-*`, `@rushstack/*`, `react`, `@types/react`** — managed as
  a coordinated set. Do **not** update them individually.
- **`npm audit fix --force`** — never use; it downgrades SPFx build-rig
  packages and breaks the build.
- **Build tool** — Heft (no Gulp). Use `npm` scripts; never run `npx heft`
  directly.
- **Graph permissions** — the web part has none of its own. All data fetching
  goes through the Azure Function via `AadHttpClient`. Do not add
  `webApiPermissionRequests` entries or introduce direct Graph calls.

---

## Pull Request Process

1. Fork the repository and create a feature branch from `main`.
2. Make your changes and follow the workflows above (lint, test, commit style).
3. Push and open a pull request against `main`. Fill in the PR template.
4. CI (`ci.yml`) runs automatically — all checks must be green before merge.
5. At least one maintainer review is required before merging.
6. Squash-merge into `main` when approved. The commit message should still
   follow Conventional Commits.

---

## Reporting Bugs

Open a [GitHub issue](https://github.com/workoho/spfx-guest-sponsor-info/issues/new)
and include:

- A concise description of the problem.
- Steps to reproduce.
- Expected behaviour vs. actual behaviour.
- SPFx version, browser, and SharePoint Online region (if relevant).
- Any error messages from the browser console or Azure Function logs.

---

## Requesting Features

Open a [GitHub issue](https://github.com/workoho/spfx-guest-sponsor-info/issues/new)
and describe:

- The use case and the problem it solves.
- Your proposed behaviour or UI (screenshots or mockups are welcome).
- Any alternatives you have considered.

Feature requests are evaluated against the project goals described in
[docs/architecture.md](docs/architecture.md).

---

## Security Vulnerabilities

**Do not open a public issue for security vulnerabilities.**

Please report security issues privately via
[GitHub Security Advisories](https://github.com/workoho/spfx-guest-sponsor-info/security/advisories/new)
or by emailing [security@workoho.com](mailto:security@workoho.com).
This is a volunteer-maintained project — responses are best effort.

See [SECURITY.md](SECURITY.md) and [docs/security-assessment.md](docs/security-assessment.md)
for the full policy, scope, and threat model.
