#!/usr/bin/env bash
# Auto-fix all lint issues (TypeScript/ESLint, SCSS, Markdown, JSON, shell) for both
# the SPFx web part and the Azure Function.
#
# Usage:
#   scripts/lint-fix.sh
#
# Applies auto-fixable corrections in-place. Issues that cannot be fixed
# automatically are reported but do not abort the run (exit 0 always).
# Run scripts/lint.sh afterwards to verify no issues remain.
#
# For CI use scripts/lint.sh instead — it never modifies files.
# Note: Bicep lint (az bicep lint) and shellcheck have no auto-fix mode;
# they are included in scripts/lint.sh but skipped here.

set -euo pipefail

# Always run from the repository root so npm scripts resolve correctly.
cd "$(dirname "${BASH_SOURCE[0]}")/.."

echo "[ 1/6 ] ESLint --fix (TypeScript — web part)..."
npm run fix:ts
echo "  ✓ done"

echo ""
echo "[ 2/6 ] ESLint --fix (TypeScript — Azure Function)..."
npm run fix:ts:func
echo "  ✓ done"

echo ""
echo "[ 3/6 ] Stylelint --fix (SCSS)..."
npm run fix:scss
echo "  ✓ done"

echo ""
echo "[ 4/6 ] Markdownlint --fix (Docs)..."
npm run fix:md
echo "  ✓ done"

echo ""
echo "[ 5/6 ] Prettier --write (JSON/JSONC)..."
npm run fix:json
echo "  ✓ done"

echo ""
echo "[ 6/6 ] shfmt --write (shell scripts)..."
npm run fix:sh
echo "  ✓ done"

echo ""
echo "✓ All fixers ran. Run scripts/lint.sh to verify no issues remain."
