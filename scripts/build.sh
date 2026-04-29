#!/usr/bin/env bash
# Copyright 2026 Workoho GmbH <https://workoho.com>
# Author: Julian Pawlowski <https://github.com/jpawlowski>
# Licensed under PolyForm Shield License 1.0.0 <https://polyformproject.org/licenses/shield/1.0.0>
#
# Full production build — intended for CI and release.
#
# Usage:
#   scripts/build.sh
#
# Produces: sharepoint/solution/guest-sponsor-info.sppkg
#
# Runs npm ci (clean install) for both the SPFx web part and the Azure
# Function, then builds both. For the web part this compiles TypeScript,
# bundles assets, runs the Jest test suite, and creates the .sppkg package.
# For the Azure Function it compiles TypeScript to dist/. Fails with a
# non-zero exit code if the expected artifact is not produced.
#
# For local development, use 'npm run build' directly to skip the slow
# npm ci reinstall.

set -euo pipefail

# Always run from the repository root so paths resolve correctly.
cd "$(dirname "${BASH_SOURCE[0]}")/.."

# shellcheck source=scripts/colors.sh
source "$(dirname "${BASH_SOURCE[0]}")/colors.sh"

step() {
  echo "${C_BLD}[ $1 ] $2${C_RST}"
}

step "1/4" "Installing web part dependencies…"
gha_group_start "Install web part dependencies"
npm ci
gha_group_end

step "2/4" "Building solution (compile · bundle · test · package)…"
gha_group_start "Build web part solution"
npm run build
gha_group_end

PKG="sharepoint/solution/guest-sponsor-info.sppkg"
if [[ ! -f "$PKG" ]]; then
  echo "${C_RED}ERROR:${C_RST} Expected artifact not found: ${PKG}" >&2
  gha_error "Expected artifact not found: ${PKG}"
  exit 1
fi

step "3/4" "Installing Azure Function dependencies…"
gha_group_start "Install Azure Function dependencies"
npm ci --prefix azure-function
gha_group_end

step "4/4" "Building Azure Function…"
gha_group_start "Build Azure Function"
npm run build --prefix azure-function
gha_group_end

echo "${C_GRN}✓${C_RST} Artifact ready: ${C_BLD}$(du -sh "$PKG" | cut -f1)${C_RST}  ${PKG}"
gha_notice "Build complete: ${PKG} is ready."
next_steps "Build complete." \
  "Deploy package: ${C_BLD}sharepoint/solution/guest-sponsor-info.sppkg${C_RST}" \
  "Optional checks: ${C_BLD}scripts/lint.sh${C_RST} and ${C_BLD}scripts/test.sh${C_RST}"
