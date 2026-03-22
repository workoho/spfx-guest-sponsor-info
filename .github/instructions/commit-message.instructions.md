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

| Part | Rule |
| ---- | ---- |
| Header (entire first line) | **≤ 100 characters** |
| Each body line | **≤ 100 characters** — wrap manually; do NOT let a single line exceed 100 chars |
| Each footer line | **≤ 100 characters** |
| Subject case | **lower-case** — never start the subject with a capital letter |
| Subject trailing period | **not allowed** |
| Type | must be one of the allowed types below |

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

Translate into da-dk, fi-fi, ja-jp, nb-no, sv-se, pt-br,
zh-cn, pl-pl, and nl-nl — all 14 locale files now covered.
```

```text
docs: update architecture decision record for graph scopes
```
