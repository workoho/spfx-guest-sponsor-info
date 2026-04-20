#!/usr/bin/env bash
# Copyright 2026 Workoho GmbH <https://workoho.com>
# Author: Julian Pawlowski <https://github.com/jpawlowski>
# Licensed under PolyForm Shield License 1.0.0 <https://polyformproject.org/licenses/shield/1.0.0>
#
# Run all linters (TypeScript/ESLint, Markdown, locale files, YAML, GitHub Actions, Bicep,
# Shell, PowerShell) for both the SPFx web part, website docs, and the Azure Function.
#
# Usage:
#   scripts/lint.sh

set -euo pipefail

# Always run from the repository root so npm scripts resolve correctly.
cd "$(dirname "${BASH_SOURCE[0]}")/.."

# shellcheck source=scripts/colors.sh
source "$(dirname "${BASH_SOURCE[0]}")/colors.sh"

EXIT=0

step() {
  echo "${C_BLD}[ $1 ] $2${C_RST}"
}

step "1/9" "ESLint (TypeScript — web part)…"
gha_group_start "ESLint (TypeScript — web part)"
if npm run lint:ts; then
  echo "  ${C_GRN}✓${C_RST} ESLint passed"
else
  echo "  ${C_RED}✗${C_RST} ESLint found issues"
  gha_warning "ESLint found issues in the web part TypeScript sources."
  EXIT=1
fi
gha_group_end

echo ""
step "2/9" "ESLint (TypeScript — Azure Function)…"
gha_group_start "ESLint (TypeScript — Azure Function)"
if npm run lint:ts:func; then
  echo "  ${C_GRN}✓${C_RST} ESLint passed"
else
  echo "  ${C_RED}✗${C_RST} ESLint found issues"
  gha_warning "ESLint found issues in the Azure Function TypeScript sources."
  EXIT=1
fi
gha_group_end

echo ""
step "3/9" "Markdownlint (Docs + Website)…"
gha_group_start "Markdownlint (Docs + Website)"
if npm run lint:md; then
  echo "  ${C_GRN}✓${C_RST} Markdownlint passed"
else
  echo "  ${C_RED}✗${C_RST} Markdownlint found issues"
  gha_warning "Markdownlint found issues in documentation or website markdown files."
  EXIT=1
fi
gha_group_end

echo ""
step "4/9" "Prettier + consistency (Locale .js files)…"
gha_group_start "Prettier + consistency (Locale .js files)"
if npm run lint:loc; then
  echo "  ${C_GRN}✓${C_RST} Locale file checks passed"
else
  echo "  ${C_RED}✗${C_RST} Locale file checks found issues"
  gha_warning "Locale file formatting, syntax, or key-order checks found issues."
  EXIT=1
fi
gha_group_end

echo ""
step "5/9" "Prettier --check (YAML — all)…"
gha_group_start "Prettier --check (YAML — all)"
if npm run lint:yml; then
  echo "  ${C_GRN}✓${C_RST} YAML formatting check passed"
else
  echo "  ${C_RED}✗${C_RST} YAML formatting check found issues"
  gha_warning "Prettier YAML check found issues in YAML files."
  EXIT=1
fi
gha_group_end

echo ""
step "6/9" "actionlint (GitHub Actions)…"
gha_group_start "actionlint (GitHub Actions)"
if npm run lint:actions; then
  echo "  ${C_GRN}✓${C_RST} actionlint passed"
else
  echo "  ${C_RED}✗${C_RST} actionlint found issues"
  gha_warning "actionlint found issues in GitHub Actions workflow files."
  EXIT=1
fi
gha_group_end

echo ""
step "7/9" "Bicep lint (Azure Function infra)…"
gha_group_start "Bicep lint (Azure Function infra)"
if npm run lint:bicep; then
  echo "  ${C_GRN}✓${C_RST} Bicep lint passed"
else
  echo "  ${C_RED}✗${C_RST} Bicep lint found issues"
  gha_warning "Bicep lint found issues in the Azure Function infrastructure templates."
  EXIT=1
fi
gha_group_end

echo ""
step "8/9" "shellcheck (Shell scripts)…"
gha_group_start "shellcheck (Shell scripts)"
if npm run lint:sh; then
  echo "  ${C_GRN}✓${C_RST} shellcheck passed"
else
  echo "  ${C_RED}✗${C_RST} shellcheck found issues"
  gha_warning "shellcheck found issues in repository shell scripts."
  EXIT=1
fi
gha_group_end

echo ""
step "9/9" "PSScriptAnalyzer (PowerShell)…"
gha_group_start "PSScriptAnalyzer (PowerShell)"
if npm run lint:ps; then
  echo "  ${C_GRN}✓${C_RST} PSScriptAnalyzer passed"
else
  echo "  ${C_RED}✗${C_RST} PSScriptAnalyzer found issues"
  gha_warning "PSScriptAnalyzer found issues in PowerShell scripts."
  EXIT=1
fi
gha_group_end

echo ""
if [[ $EXIT -eq 0 ]]; then
  echo "${C_GRN}✓ All linters passed.${C_RST}"
  gha_notice "All linters passed successfully."
else
  echo "${C_RED}✗ One or more linters reported issues — see above.${C_RST}"
  next_steps "Run ${C_BLD}scripts/lint-fix.sh${C_RST} for auto-fixable issues." \
    "Then rerun ${C_BLD}scripts/lint.sh${C_RST} to verify a clean state."
fi

exit $EXIT
