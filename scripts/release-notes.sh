#!/usr/bin/env bash
# Copyright 2026 Workoho GmbH <https://workoho.com>
# Author: Julian Pawlowski <https://github.com/jpawlowski>
# Licensed under PolyForm Shield License 1.0.0 <https://polyformproject.org/licenses/shield/1.0.0>
#
# Preview or generate release notes from Conventional Commit history.
#
# Uses git-cliff (https://git-cliff.org) with the project's cliff.toml config.
# On Linux (x86-64 and arm64) the script auto-installs git-cliff into ~/.local/bin
# is not already on PATH.  On other platforms install manually:
#   https://git-cliff.org/docs/installation
#
# Usage:
#   ./scripts/release-notes.sh              # commits not yet tagged (preview)
#   ./scripts/release-notes.sh --latest     # changes in the most recent tag
#   ./scripts/release-notes.sh --tag v1.2.3 # simulate notes for an upcoming tag
#
# Any extra arguments are forwarded verbatim to git-cliff.
# Output goes to stdout; redirect to a file as needed:
#   ./scripts/release-notes.sh > RELEASE_NOTES.md

set -euo pipefail

# Always run from the repository root so relative paths resolve correctly.
cd "$(dirname "${BASH_SOURCE[0]}")/.."

# shellcheck source=scripts/colors.sh
source "$(dirname "${BASH_SOURCE[0]}")/colors.sh"

# ── Pinned version ───────────────────────────────────────────────────────────
# Update this string to upgrade git-cliff.
# The devcontainer post-create.sh and the release workflow both read
# this value from here, so this is the single place to change.
GIT_CLIFF_VERSION="2.12.0"

# ── Auto-install on Linux x86-64 (devcontainer + GitHub Actions runner) ──────
if ! command -v git-cliff &>/dev/null; then
  ARCH="$(uname -m)"
  OS="$(uname -s)"
  if [[ "${OS}" != "Linux" ]]; then
    echo "${C_RED}ERROR:${C_RST} git-cliff is not installed and auto-install is only supported on Linux." >&2
    gha_error "git-cliff is not installed and auto-install is only supported on Linux."
    important "Install git-cliff manually:" \
      "${C_BLD}https://git-cliff.org/docs/installation${C_RST}"
    exit 1
  fi
  case "${ARCH}" in
    x86_64) TRIPLE="x86_64-unknown-linux-musl" ;;
    aarch64) TRIPLE="aarch64-unknown-linux-musl" ;;
    *)
      echo "${C_RED}ERROR:${C_RST} git-cliff auto-install is not supported for ${ARCH}." >&2
      gha_error "git-cliff auto-install is not supported for ${ARCH}."
      important "Install git-cliff manually:" \
        "${C_BLD}https://git-cliff.org/docs/installation${C_RST}"
      exit 1
      ;;
  esac
  INSTALL_DIR="${HOME}/.local/bin"
  TARBALL="git-cliff-${GIT_CLIFF_VERSION}-${TRIPLE}.tar.gz"
  echo "${C_DIM}git-cliff not found — installing v${GIT_CLIFF_VERSION} into ${INSTALL_DIR}…${C_RST}" >&2
  mkdir -p "${INSTALL_DIR}"
  TARBALL_TMP="$(mktemp)"
  # SC2064: expand $TARBALL_TMP now so the trap captures the exact path.
  # shellcheck disable=SC2064
  trap "rm -f '${TARBALL_TMP}'" EXIT
  curl -fsSL \
    "https://github.com/orhun/git-cliff/releases/download/v${GIT_CLIFF_VERSION}/${TARBALL}" \
    -o "${TARBALL_TMP}"
  tar -xz -C "${INSTALL_DIR}" \
    --strip-components=1 \
    "git-cliff-${GIT_CLIFF_VERSION}/git-cliff" <"${TARBALL_TMP}"
  rm -f "${TARBALL_TMP}"
  trap - EXIT
  chmod +x "${INSTALL_DIR}/git-cliff"
  export PATH="${INSTALL_DIR}:${PATH}"
  echo "${C_GRN}✓${C_RST} git-cliff ${GIT_CLIFF_VERSION} installed." >&2
  gha_notice "git-cliff ${GIT_CLIFF_VERSION} installed successfully."
fi

REPO_ROOT="$(pwd)"

if [[ "${1:-}" != "--help" && "${1:-}" != "-h" ]]; then
  hint "Generating release notes from ${C_BLD}${REPO_ROOT}/cliff.toml${C_RST}" \
    "Output is written to stdout. Redirect with ${C_BLD}> RELEASE_NOTES.md${C_RST} if needed." >&2
fi

git-cliff --config "${REPO_ROOT}/cliff.toml" --strip header "${@:---unreleased}"
