#!/usr/bin/env bash
# Copyright 2026 Workoho GmbH <https://workoho.com>
# Author: Julian Pawlowski <https://github.com/jpawlowski>
# Licensed under PolyForm Shield License 1.0.0 <https://polyformproject.org/licenses/shield/1.0.0>
#
# Guided SPFx version upgrade.
#
# Usage:
#   scripts/upgrade-spfx.sh
#   scripts/upgrade-spfx.sh --upgrade-latest
#   scripts/upgrade-spfx.sh <new-spfx-version>
#
# Example:
#   scripts/upgrade-spfx.sh
#   scripts/upgrade-spfx.sh --upgrade-latest
#   scripts/upgrade-spfx.sh 1.23.0
#
# SPFx is NOT upgraded with 'npm update'. The @microsoft/sp-* packages form a
# tightly coupled suite that must all move to exactly the same version at once,
# and the Yeoman generator must also be updated to regenerate scaffolded config.
#
# This script guides you through the process and checks preconditions, but some
# steps (reviewing generated config diffs) require human judgment.

set -euo pipefail

# Always run from the repository root so npm scripts resolve correctly.
cd "$(dirname "${BASH_SOURCE[0]}")/.."

# shellcheck source=scripts/colors.sh
source "$(dirname "${BASH_SOURCE[0]}")/colors.sh"

header() {
  echo "${C_BLD}═══════════════════════════════════════════════════════════${C_RST}"
  echo "${C_BLD}  $*${C_RST}"
  echo "${C_BLD}═══════════════════════════════════════════════════════════${C_RST}"
  echo ""
}

step() {
  echo "${C_CYN}[ $1 ]${C_RST} $2"
}

ok() {
  echo "  ${C_GRN}✓${C_RST} $*"
}

fail() {
  echo "  ${C_RED}ERROR:${C_RST} $*" >&2
}

NEW_VERSION="${1:-}"

CURRENT_CONFIG_SPFX="$(node -p "require('./package.json').config?.spfxVersion || ''")"
CURRENT_DEP_SPFX="$(node -p "require('./package.json').dependencies['@microsoft/sp-core-library'] || ''")"
YO_VERSION="$(node -p "require('./package.json').config?.yoVersion || '7.0.0'")"

if [[ -z "${CURRENT_CONFIG_SPFX}" ]]; then
  fail "package.json config.spfxVersion is not set."
  important "Add ${C_BLD}config.spfxVersion${C_RST} to ${C_BLD}package.json${C_RST} before running this script."
  exit 1
fi

if [[ "${CURRENT_CONFIG_SPFX}" != "${CURRENT_DEP_SPFX}" ]]; then
  fail "package.json is inconsistent."
  important \
    "config.spfxVersion=${CURRENT_CONFIG_SPFX}" \
    "@microsoft/sp-core-library=${CURRENT_DEP_SPFX}" \
    "" \
    "Align them first, then rerun ${C_BLD}scripts/upgrade-spfx.sh${C_RST}."
  exit 1
fi

if [[ -z "${NEW_VERSION}" ]]; then
  LATEST_SPFX="$(npm view @microsoft/sp-core-library version 2>/dev/null || true)"
  if [[ -z "${LATEST_SPFX}" ]]; then
    fail "Could not fetch latest SPFx version from npm."
    hint "Check your network connection and run again."
    exit 1
  fi

  header "SPFx version check"
  echo "  Current (package.json config.spfxVersion): ${CURRENT_CONFIG_SPFX}"
  echo "  Latest  (@microsoft/sp-core-library):      ${LATEST_SPFX}"
  echo ""

  if [[ "${CURRENT_CONFIG_SPFX}" == "${LATEST_SPFX}" ]]; then
    ok "You are up to date."
    hint "No action needed."
  else
    important "A newer SPFx version is available: ${CURRENT_CONFIG_SPFX} -> ${LATEST_SPFX}"
    next_steps \
      "${C_BLD}scripts/upgrade-spfx.sh ${LATEST_SPFX}${C_RST}         # upgrade to latest" \
      "${C_BLD}scripts/upgrade-spfx.sh <version>${C_RST}      # choose specific target" \
      "${C_BLD}npm view @microsoft/sp-core-library versions${C_RST} # list available versions"
  fi
  exit 0
fi

if [[ "${NEW_VERSION}" == "--upgrade-latest" ]]; then
  LATEST_SPFX="$(npm view @microsoft/sp-core-library version 2>/dev/null || true)"
  if [[ -z "${LATEST_SPFX}" ]]; then
    fail "Could not fetch latest SPFx version from npm."
    hint "Check your network connection and run again."
    exit 1
  fi

  if [[ "${CURRENT_CONFIG_SPFX}" == "${LATEST_SPFX}" ]]; then
    ok "Already up to date (${CURRENT_CONFIG_SPFX}). Nothing to upgrade."
    exit 0
  fi

  hint "Auto mode: upgrading from ${CURRENT_CONFIG_SPFX} to latest ${LATEST_SPFX}."
  NEW_VERSION="${LATEST_SPFX}"
fi

# Strip leading 'v' if provided.
NEW_VERSION="${NEW_VERSION#v}"

header "SPFx upgrade guide -> ${NEW_VERSION}"

# ── Step 1: Check for a clean working tree ────────────────────────────────
step "1/6" "Checking working tree..."
if ! git diff --quiet || ! git diff --cached --quiet; then
  fail "You have uncommitted changes. Commit or stash them first."
  exit 1
fi
ok "Working tree is clean."
echo ""

# ── Step 2: Verify the target SPFx version exists on npm ─────────────────
step "2/6" "Verifying SPFx ${NEW_VERSION} exists on npm..."
if ! npm view "@microsoft/sp-core-library@${NEW_VERSION}" version >/dev/null 2>&1; then
  fail "@microsoft/sp-core-library@${NEW_VERSION} not found on npm."
  hint "Check available versions: ${C_BLD}npm view @microsoft/sp-core-library versions${C_RST}"
  exit 1
fi
ok "Version ${NEW_VERSION} exists."
echo ""

# ── Step 3: Check Node.js compatibility ──────────────────────────────────
step "3/6" "Node.js compatibility..."
echo "  Current Node: $(node --version)"
echo "  → Check the SPFx ${NEW_VERSION} release notes for the required Node range:"
echo "    https://learn.microsoft.com/sharepoint/dev/spfx/compatibility"
echo "  If a different Node version is required, update:"
echo "    - .devcontainer/devcontainer.json  (NODE_VERSION build arg)"
echo "    - .devcontainer/Dockerfile         (NODE_VERSION ARG default)"
echo "    - package.json                     (engines.node field)"
echo "    - .github/workflows/ci.yml         (node-version)"
echo "    - .github/workflows/release.yml    (node-version)"
echo "    - .nvmrc"
echo ""

# ── Step 4: Install updated packages ─────────────────────────────────────
step "4/6" "Installing SPFx ${NEW_VERSION} packages..."
echo "  Updating package.json config.spfxVersion..."
npm pkg set "config.spfxVersion=${NEW_VERSION}" >/dev/null

SPFX_DEPS=(
  "@microsoft/sp-component-base@${NEW_VERSION}"
  "@microsoft/sp-core-library@${NEW_VERSION}"
  "@microsoft/sp-http@${NEW_VERSION}"
  "@microsoft/sp-lodash-subset@${NEW_VERSION}"
  "@microsoft/sp-office-ui-fabric-core@${NEW_VERSION}"
  "@microsoft/sp-property-pane@${NEW_VERSION}"
  "@microsoft/sp-webpart-base@${NEW_VERSION}"
)
SPFX_DEV_DEPS=(
  "@microsoft/eslint-config-spfx@${NEW_VERSION}"
  "@microsoft/eslint-plugin-spfx@${NEW_VERSION}"
  "@microsoft/sp-module-interfaces@${NEW_VERSION}"
  "@microsoft/spfx-heft-plugins@${NEW_VERSION}"
  "@microsoft/spfx-web-build-rig@${NEW_VERSION}"
)

npm install "${SPFX_DEPS[@]}"
npm install --save-dev "${SPFX_DEV_DEPS[@]}"

if [[ -f ".yo-rc.json" ]]; then
  echo "  Updating .yo-rc.json generator version..."
  tmp_file="$(mktemp)"
  jq --arg version "${NEW_VERSION}" '."@microsoft/generator-sharepoint".version = $version' \
    .yo-rc.json >"${tmp_file}"
  mv "${tmp_file}" .yo-rc.json
fi

# Install matching Rushstack packages (versions are coordinated with SPFx).
echo ""
echo "  Rushstack packages are managed by @microsoft/spfx-web-build-rig."
echo "  Run 'npm install' to let npm resolve compatible versions from the lockfile."
npm install
ok "Packages installed."
echo ""

# ── Step 5: Run yo to regenerate scaffolded config ───────────────────────
step "5/6" "Regenerating config via Yeoman..."
echo "  The Yeoman generator updates config files (tsconfig, heft configs, etc.)."
echo "  You will be asked whether to overwrite each file — review diffs carefully."
hint \
  "Install the updated generator first:" \
  "${C_BLD}npm install --global yo@${YO_VERSION} @microsoft/generator-sharepoint@${NEW_VERSION}${C_RST}" \
  "Then run:" \
  "${C_BLD}yo @microsoft/sharepoint --skip-install${C_RST}" \
  "" \
  "Note: .devcontainer/Dockerfile and Copilot setup derive generator version" \
  "from package.json config.spfxVersion."

# ── Step 6: Verify the build ─────────────────────────────────────────────
step "6/6" "Verify the upgrade..."
next_steps \
  "Run the full build and test suite:" \
  "${C_BLD}npm test${C_RST}" \
  "${C_BLD}npm run lint${C_RST}" \
  "${C_BLD}npm run build${C_RST}  # produces the .sppkg" \
  "" \
  "If everything passes, commit with:" \
  "${C_BLD}git add package.json package-lock.json config/${C_RST}" \
  "${C_BLD}git commit -m \"chore: upgrade SPFx to ${NEW_VERSION}\"${C_RST}"

ok "Upgrade preparation complete. Manual steps remain (see hints/next steps above)."
