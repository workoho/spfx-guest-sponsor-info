#!/usr/bin/env bash
# Copyright 2026 Workoho GmbH <https://workoho.com>
# Author: Julian Pawlowski <https://github.com/jpawlowski>
# Licensed under PolyForm Shield License 1.0.0 <https://polyformproject.org/licenses/shield/1.0.0>
#
# Reset the development environment to a clean state.
#
# Usage:
#   scripts/reset.sh
#
# Removes all build outputs, caches, and node_modules for both the SPFx
# web part and the Azure Function, then re-installs dependencies via
# scripts/bootstrap.sh. Useful after branch switches with major dependency
# changes or when builds produce unexpected results.

set -euo pipefail

# Always run from the repository root so paths resolve correctly.
cd "$(dirname "${BASH_SOURCE[0]}")/.."

# shellcheck source=scripts/colors.sh
source "$(dirname "${BASH_SOURCE[0]}")/colors.sh"

echo "${C_BLD}Resetting development environment…${C_RST}"
echo "${C_DIM}Cleaning web part build outputs…${C_RST}"
npm run clean
echo "${C_GRN}✓${C_RST} Web part build outputs cleaned."

echo ""
echo "${C_DIM}Removing web part node_modules…${C_RST}"
rm -rf node_modules
echo "${C_GRN}✓${C_RST} Web part node_modules removed."

echo ""
echo "${C_DIM}Cleaning Azure Function build output and node_modules…${C_RST}"
rm -rf azure-function/dist azure-function/node_modules
echo "${C_GRN}✓${C_RST} Azure Function dist/node_modules removed."

echo ""
echo "${C_DIM}Re-installing dependencies…${C_RST}"
./scripts/bootstrap.sh
echo ""
next_steps "${C_BLD}./scripts/dev-webpart.sh${C_RST}    # start the SPFx dev server" \
  "${C_BLD}./scripts/dev-function.sh${C_RST}   # start the Azure Function locally"
