#!/usr/bin/env bash
# Copyright 2026 Workoho GmbH <https://workoho.com>
# Author: Julian Pawlowski <https://github.com/jpawlowski>
# Licensed under PolyForm Shield License 1.0.0 <https://polyformproject.org/licenses/shield/1.0.0>
#
# Auto-fix all lint issues (TypeScript/ESLint, Markdown, locale JS, YAML, JSON, shell)
# for both the SPFx web part, website docs, and the Azure Function.
#
# Usage:
#   scripts/lint-fix.sh
#
# Applies auto-fixable corrections in-place. Issues that cannot be fixed
# automatically are reported but do not abort the run (exit 0 always).
# Run scripts/lint.sh afterwards to verify no issues remain.
#
# For CI use scripts/lint.sh instead — it never modifies files.
# Note: Bicep lint (az bicep lint), shellcheck, actionlint, and PSScriptAnalyzer have no auto-fix
# mode; they are included in scripts/lint.sh but skipped here.

set -euo pipefail

# Always run from the repository root so npm scripts resolve correctly.
cd "$(dirname "${BASH_SOURCE[0]}")/.."

# shellcheck source=scripts/colors.sh
source "$(dirname "${BASH_SOURCE[0]}")/colors.sh"

echo "[ 1/7 ] ESLint --fix (TypeScript — web part)..."
npm run fix:ts
echo "  ✓ done"

echo ""
echo "[ 2/7 ] ESLint --fix (TypeScript — Azure Function)..."
npm run fix:ts:func
echo "  ✓ done"

echo ""
echo "[ 3/7 ] Markdownlint --fix (Docs + Website)..."
npm run fix:md
echo "  ✓ done"

echo ""
echo "[ 4/7 ] Prettier --write (YAML — all)..."
npm run fix:yml
echo "  ✓ done"

echo ""
echo "[ 5/7 ] Prettier --write (JSON/JSONC)..."
npm run fix:json
echo "  ✓ done"

echo ""
echo "[ 6/7 ] Prettier --write (Locale .js files)..."
npm run fix:loc
echo "  ✓ done"

echo ""
echo "[ 7/7 ] shfmt --write (shell scripts)..."
npm run fix:sh
echo "  ✓ done"

echo ""
echo "${C_GRN}✓${C_RST} All fixers ran."
next_steps "Run ${C_BLD}scripts/lint.sh${C_RST} to verify no issues remain."
