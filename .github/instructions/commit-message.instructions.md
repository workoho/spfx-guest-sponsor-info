---
description: >
  Commit message style rules for this repository.
  Apply whenever writing or generating a git commit message.
applyTo: "**"
---

# Commit Message Rules

This repository enforces **Conventional Commits** via `commitlint`
(`@commitlint/config-conventional`). The pre-commit hook will reject messages
that violate these rules, so generate correct messages from the start.

## Format

```text
<type>(<optional scope>): <subject>

<optional body>

<optional footer>
```

## Hard Limits (enforced by commitlint — violations block the commit)

| Part                       | Rule                                                                            |
| -------------------------- | ------------------------------------------------------------------------------- |
| Header (entire first line) | **≤ 100 characters**                                                            |
| Each body line             | **≤ 100 characters** — wrap manually; do NOT let a single line exceed 100 chars |
| Each footer line           | **≤ 100 characters**                                                            |
| Subject case               | **lower-case** — never start the subject with a capital letter                  |
| Subject trailing period    | **not allowed**                                                                 |
| Type                       | must be one of the allowed types below                                          |

## Allowed Types

`fix` · `feat` · `build` · `chore` · `ci` · `docs` · `perf` · `refactor` · `revert` · `style` · `test`

## Body Line Wrapping

The 100-character limit applies to **each individual line**, not the total body
length. Break prose naturally at word boundaries so no line exceeds 100 chars.

Good:

```text
feat(sponsor): add version mismatch warning banner

Show a Fluent UI MessageBar when the web part and the Azure Function
are running different versions. The banner is dismissed automatically
when the versions match again after a page reload.
```

Bad (second body line is 113 chars — commitlint will reject):

```text
feat(sponsor): add version mismatch warning banner

Show a Fluent UI MessageBar when the web part and the Azure Function are running different versions.
```

## Complete Examples

```text
feat(sponsor): add version mismatch warning banner

Show a Fluent UI MessageBar when the web part and the
Azure Function are running different versions.
```

```text
fix(map): show external map link even without Azure Maps key

Previously the external link was gated behind the Azure Maps
subscription key check. Split into two separate conditions so
the fallback link renders whenever showAddressMap is enabled.
```

```text
chore(i18n): add native translations for VersionMismatchMessage

Translate into da-dk, fi-fi, ja-jp, nb-no, sv-se, pt-pt,
zh-cn, pl-pl, and nl-nl — all 17 locale files now covered.
```

```text
docs: update architecture decision record for graph scopes
```

## User-facing vs. Internal Changes

Release notes are read by **SharePoint admins, Azure admins, and end users**, not developers.
The commit type and scope together decide in which section a change appears
(or whether it appears at all). Apply these rules:

### Use `feat` / `fix` / `perf` WITHOUT a developer scope

...only when the change is **visible to an admin or end user**:

- New or changed behaviour in the web part UI or its property pane
- New or changed behaviour in the Azure Function API (response shape, new
  endpoint, auth changes)
- A bug fix that users would have noticed (wrong data, broken layout, errors)
- A performance improvement end users can feel

```text
feat(sponsor): show job title below sponsor display name
fix(graph): handle missing profile photo gracefully
perf(card): cache sponsor photos for the page lifetime
```

### Use `feat` / `fix` / `perf` WITH a developer scope → Internal Improvements

When the same types apply to **developer tooling**, git-cliff will place them
under "Internal Improvements" instead of the user-facing sections. Use a
developer scope whenever the change affects:

| Area | Scope to use |
|---|---|
| Release / version scripts | `scripts` |
| GitHub Actions workflows | `ci` or `github` |
| Azure infra / Bicep templates | `infra` |
| Dev container / Dockerfile | `devcontainer` |
| Git hooks (Husky) | `husky` |
| Lint / format config | `lint` or `config` |
| Locale / i18n files only | `i18n` or `loc` |
| Build tooling / Heft config | `build` or `tooling` |
| TypeScript types only | `types` |
| Docker | `docker` |

```text
feat(scripts): add --dry-run flag to set-version.sh
fix(ci): pass GITHUB_TOKEN to release workflow step
feat(devcontainer): pre-install shellcheck and shfmt
```

### Use `chore`, `build`, `ci`, `docs`, `style`, `test`, `refactor`

For changes that are clearly non-functional from a user perspective and do not
fit the developer-scope pattern above. These are filtered out of release notes
entirely (except `chore(deps)` which gets its own section).

```text
chore(deps): bump eslint from 8.x to 9.x
docs: add deployment guide for sovereign cloud tenants
test: add unit tests for SponsorService.fetchSponsors
refactor(card): extract usePhoto hook from SponsorCard
```

### Quick decision tree

```text
Is the change visible to a SharePoint admin, Azure admin, or end user?
├─ Yes → feat / fix / perf  (no developer scope)
└─ No  → Is it an improvement to dev tooling / CI / scripts?
          ├─ Yes → feat / fix / perf  WITH a developer scope from the table
          └─ No  → chore / build / ci / docs / style / test / refactor
```
