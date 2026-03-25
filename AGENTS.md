# Agent Customization Guide

This file documents tools and best practices for AI coding agents (GitHub Copilot, Claude, etc.) working in this repository.

## DevContainer Tools Available

The following CLI tools are pre-installed in `.devcontainer/Dockerfile` and should be used over their slower alternatives:

### File & Search Tools

| Tool | Alternative | Command Example | Best For |
|------|-------------|-----------------|----------|
| **ripgrep** (`rg`) | `grep` | `rg "ComponentName" src/ --type ts` | Fast code/text search across codebase |
| **fd** | `find` | `fd "\.tsx$" src/` | Finding files by pattern (faster, cleaner syntax) |
| **bat** | `cat` | `bat src/components/MyComponent.tsx` | Syntax-highlighted file preview in terminal |
| **delta** | `git diff` | `git diff HEAD` | Syntax-highlighted git diffs with line numbers |
| **shellcheck** | — | `shellcheck scripts/*.sh` | Static analysis of shell scripts |
| **shfmt** | — | `shfmt -w -i 2 -ci scripts/*.sh` | Format shell scripts (2-space indent, case-indent) |

### Data Processing Tools

| Tool | Use Case | Example |
|------|----------|---------|
| **jq** | JSON parsing | `cat package.json \| jq '.scripts'` |
| **yq** | YAML/config parsing | `yq eval '.devDependencies' package.json` |
| **fzf** | Interactive fuzzy selection | `fd "test.ts" \| fzf` (pick file interactively) |

## Recommended Agent Usage Patterns

### 1. **Searching Code**

```bash
# ✅ Good: Using rg for fast search
rg "SponsorService" src/ --type ts

# ❌ Avoid: Using grep (slower)
grep -r "SponsorService" src/
```

### 2. **Finding Files**

```bash
# ✅ Good: Using fd for clean syntax
fd "GuestSponsorInfo\.tsx"

# ❌ Avoid: Using find (verbose)
find . -name "*GuestSponsorInfo*" -type f
```

### 3. **Reading Configuration**

```bash
# ✅ Good: Using yq/jq for structured parsing
yq eval '.engines' package.json

# ❌ Avoid: Manual regex parsing
grep "engines" package.json | cut -d '"'
```

### 4. **Previewing Files in Terminals**

```bash
# ✅ Good: Using bat for syntax highlighting
bat src/webparts/guestSponsorInfo/GuestSponsorInfo.tsx

# ❌ Avoid: Using cat (no syntax highlighting)
cat src/webparts/guestSponsorInfo/GuestSponsorInfo.tsx
```

## Git Workflow

### Pre-commit Hook

The `.husky/pre-commit` hook automatically:

1. Runs `npm run fix` (auto-corrects formatting issues)
2. Runs `npm run lint` (validates code quality)

**Do not** skip this hook (`--no-verify`) except for meta-changes (infrastructure-only commits).

### Commit Message Format

Use Conventional Commits (enforced by `commitlint` via `.husky/commit-msg`).
See `.github/instructions/commit-message.instructions.md` for the full ruleset.

**Hard limits** (violations abort the commit):

- Header ≤ **100 characters**
- Every body/footer line ≤ **100 characters** — wrap prose manually
- Subject must be **lower-case**, no trailing period
- Type must be one of: `fix` `feat` `build` `chore` `ci` `docs` `perf`
  `refactor` `revert` `style` `test`

```text
chore(i18n): normalize locale files to UTF-8
feat(sponsor): add sponsor card caching
fix(graph): handle missing user profile gracefully
docs: update README configuration section
```

## When to Use npm Scripts vs Direct Commands

### ✅ Use `npm` Scripts (Managed + Linted)

```bash
npm run lint        # ESLint + Stylelint + Markdownlint + locale syntax
npm run fix         # Auto-correct formatting
npm run lint:ts     # TypeScript only
npm run lint:loc    # Locale .js files (node --check syntax only)
npm test            # Jest + coverage
npm start           # Dev server with hot-reload
npm run build       # Full production build + packaging
```

### ⚠️ Avoid Raw Commands Unless Necessary

```bash
# ❌ Don't run raw heft commands
npx heft build

# ✅ Use npm scripts instead
npm run build
```

```bash
# ❌ NEVER call the husky binary directly without -- separator.
#    Husky v9 treats its first positional argument as a target directory.
#    Running 'npx husky --version' (without --) makes husky create a
#    '--version/_' directory and overwrite core.hooksPath, breaking all
#    git hooks (commits and pushes will fail with "Illegal option --").
npx husky --version   # ❌ breaks git hooks!
npx husky list        # ❌ breaks git hooks!

# ✅ Use npm scripts to manage hooks
npm run prepare                           # (re-)install / reset git hooks

# ✅ If you need the husky version, use the -- separator:
node_modules/.bin/husky -- --version      # ✅ safe
```

## Code Validation Checklist

After **every** code change:

1. **Lint** → `npm run lint` (catches 95% of issues)
2. **Test** → `npm test` (if logic changed)
3. **Fix** → `npm run fix` (auto-correct before committing)
4. **Validate** → `npm run lint` again (final check)

**Never push with lint errors or test failures.**

## Shell Script Conventions

All scripts in `scripts/*.sh` follow these rules. Before writing a new script, read
`scripts/lint.sh` for the standard boilerplate (`set -euo pipefail`, `cd`, `source colors.sh`)
and `scripts/set-version.sh` for the `maybe()` dry-run pattern.

### Colour output

Use the variables from `scripts/colors.sh` — never copy the detection block inline.

| Variable | Meaning | Typical use |
|---|---|---|
| `C_GRN` | Green | `✓` success messages |
| `C_RED` | Red | `✗` errors / fatal messages |
| `C_YLW` | Yellow bold | `⚠` warnings, menu highlights |
| `C_CYN` | Cyan | version numbers, file paths |
| `C_BLD` | Bold | section headers |
| `C_DIM` | Dim | secondary info, progress lines |
| `C_RST` | Reset | always close a colour sequence |

Colours are automatically suppressed when `$CI` is set, stdout is not a TTY,
`$NO_COLOR` is set, or `$TERM` is `"dumb"`.

### Comments

Bash is not self-documenting. Comment non-obvious constructs, especially:

- Parameter expansions: `${var%%[-+]*}` → explain what it strips and why
- Conditional flags: explain what a `git` or `npm` flag does if it is not obvious
- Decision points: why a fallback exists, what the edge case is

### Validation

After **every** shell script change:

```bash
npm run lint:sh   # shellcheck -x (SC1091 aware; follows sourced files)
npm run fix       # shfmt auto-format (2-space indent, case-indent)
```

Do not suppress shellcheck warnings with `# shellcheck disable` unless you add
a comment directly above it explaining the specific reason.

## Stack Constraints (Do Not Violate)

- **SPFx version** — Do not upgrade; use `scripts/upgrade-spfx.sh` when explicitly requested
- **Node version** — Must stay within `engines` range in `package.json` (currently `>=22.14.0 <23.0.0`)
- **React version** — Pinned at 17.0.1; do not change
- **@microsoft/** packages — Managed as coordinated set; do not individually update

## Fluent UI v9

This project uses **Fluent UI v9** (`@fluentui/react-components`) exclusively.
For import patterns, check the existing components — `GuestSponsorInfo.tsx` and
`SponsorCard.tsx` are the authoritative reference.

### Styles

Use **`makeStyles`** from `@fluentui/react-components` (Griffel) for all component-level
styles, with **`tokens`** from `@fluentui/react-components` for all colour and spacing values.
Use `mergeClasses()` for conditional class composition.
Do not add CSS/SCSS module files — all styles live in `makeStyles` hooks.

### Forbidden patterns

```typescript
// ❌ v8 components — project has fully migrated to v9
import { Persona, Callout, Panel, ActionButton, IconButton,
         Icon, TooltipHost, MessageBar } from '@fluentui/react';
// ❌ v8 icon font
initializeIcons();
// ❌ v8 styling APIs
import { mergeStyles, mergeStyleSets } from '@fluentui/merge-styles';
// ❌ Hardcoded colour values — use tokens instead
const style = { color: '#0078d4' };
// ❌ CSS/SCSS module files — use makeStyles hooks instead
import styles from './Foo.module.scss';
```

### Key component equivalents

| v8 | v9 |
|---|---|
| `Persona` (avatar display) | `Avatar` with `color="colorful"` |
| `PersonaPresence` enum | `PresenceBadge` `status` string prop |
| `Callout` | `Popover` + `PopoverTrigger` + `PopoverSurface` |
| `Panel` | `OverlayDrawer` + `DrawerHeader` + `DrawerBody` |
| `ActionButton` / `IconButton` | `Button` with `appearance="subtle"` |
| `Icon iconName="Chat"` | `<ChatRegular />` (SVG from `@fluentui/react-icons`) |
| `TooltipHost` | `Tooltip` with `relationship="label"` |
| `Link` | `Link` from `@fluentui/react-components` |
| `MessageBar` + `MessageBarType` | `MessageBar` + `MessageBarBody` with `intent` prop |
| `IButtonStyles` | `makeStyles` with `tokens` (Griffel hook) |

### Theme integration

Always wrap the component tree in `<FluentProvider>` with the site theme:

```tsx
import { FluentProvider } from '@fluentui/react-components';
import { createV9Theme } from '@fluentui/react-migration-v8-v9';
// `theme` is the IReadonlyTheme from the SPFx ThemeProvider service
const v9Theme = theme ? createV9Theme(theme) : undefined;
return <FluentProvider theme={v9Theme}>{children}</FluentProvider>;
```

## Key Files for Reference

- **Lint config** → `config/` directory + `.markdownlint.json`
- **Locale strings** → `src/webparts/guestSponsorInfo/loc/*.js` (UTF-8 native characters, no `\uXXXX` escapes)
- **Components** → `src/webparts/guestSponsorInfo/components/`
- **Services** → `src/webparts/guestSponsorInfo/services/` (Graph API calls in `SponsorService.ts`)

## Azure Infrastructure Scripts

PowerShell scripts in `azure-function/infra/` assist with one-time setup and
post-deployment verification. They require the
`Az.Accounts` and `Az.Resources` PowerShell modules (pre-installed in the dev
container) and an active `Connect-AzAccount` session.

| Script | When to run |
|---|---|
| `setup-app-registration.ps1` | Once — creates the Entra app registration before first deployment |
| `setup-graph-permissions.ps1` | Once after deployment — grants Graph permissions to the Function's managed identity |
| `Verify-DeploymentGuid.ps1` | After every fresh `azuredeploy.json` deployment — confirms CUA attribution works |

### Verify-DeploymentGuid.ps1

Source: [bmoore-msft/Verify-DeploymentGuid.ps1](https://gist.github.com/bmoore-msft/ae6b8226311014d6e7177c5127c7eba1)
(Microsoft Partner Center team)

**Purpose:** After deploying the ARM template, this script follows the
`correlationId` of the `pid-18fb4033-c9f3-41fa-a5db-e3a03b012939` deployment
and lists every Azure resource deployed in the same correlation scope. A
non-empty list confirms that Azure will correctly attribute consumption to the
Workoho Partner Center GUID. An empty list requires investigation (likely the
`pid-*` deployment was created outside a real ARM deployment).

**When to run:**

- After every fresh deployment from `azuredeploy.json` or the "Deploy to
  Azure" button
- Before a Partner Center reporting period to confirm attribution is live
- Not needed for code changes, schema updates, or configuration-only changes

**Usage:**

```powershell
.\azure-function\infra\Verify-DeploymentGuid.ps1 `
  -deploymentName pid-18fb4033-c9f3-41fa-a5db-e3a03b012939 `
  -resourceGroupName <your-resource-group>
```

Expected output: one or more Azure resource ID strings (Function App, Storage
Account, App Service Plan, etc.). See the script header for the full
prerequisites list.
