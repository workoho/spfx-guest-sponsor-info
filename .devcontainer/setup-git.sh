#!/usr/bin/env bash
# Configure git identity inside the dev container.
#
# Strategy:
#   - In GitHub Codespaces: identity is already set by GitHub — nothing to do.
#   - In a local dev container: read user.name and user.email from the host
#     gitconfig (bind-mounted read-only as ~/.gitconfig.host) and copy them
#     into the container's own global gitconfig.
#
# This script is idempotent and safe to run multiple times.

set -euo pipefail

# shellcheck source=scripts/colors.sh
source "$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)/scripts/colors.sh"

echo "Setting up Git configuration..."

# ── Codespaces: GitHub already configured everything ──────────────────────
if [[ -n "${CODESPACES:-}" ]]; then
  name="$(git config --global user.name 2>/dev/null || true)"
  email="$(git config --global user.email 2>/dev/null || true)"
  echo "  Running in GitHub Codespaces — Git already configured by GitHub"
  echo "  user.name:  ${name:-"(not set)"}"
  echo "  user.email: ${email:-"(not set)"}"
  exit 0
fi

# ── Local dev container: copy identity from host gitconfig ────────────────
HOST_GITCONFIG="${HOME}/.gitconfig.host"

if [[ -f "${HOST_GITCONFIG}" ]]; then
  # Extract values using git itself (handles includes, aliases, encoding).
  HOST_NAME="$(GIT_CONFIG_GLOBAL="${HOST_GITCONFIG}" git config --global user.name 2>/dev/null || true)"
  HOST_EMAIL="$(GIT_CONFIG_GLOBAL="${HOST_GITCONFIG}" git config --global user.email 2>/dev/null || true)"
else
  HOST_NAME=""
  HOST_EMAIL=""
fi

if [[ -n "${HOST_NAME}" ]]; then
  git config --global user.name "${HOST_NAME}"
  echo "  user.name:  ${HOST_NAME}"
else
  echo "  user.name:  (not found in host gitconfig)"
fi

if [[ -n "${HOST_EMAIL}" ]]; then
  git config --global user.email "${HOST_EMAIL}"
  echo "  user.email: ${HOST_EMAIL}"
else
  echo "  user.email: (not found in host gitconfig)"
fi

# ── Warn if identity is still missing ────────────────────────────────────
FINAL_NAME="$(git config --global user.name 2>/dev/null || true)"
FINAL_EMAIL="$(git config --global user.email 2>/dev/null || true)"

if [[ -z "${FINAL_NAME}" || -z "${FINAL_EMAIL}" ]]; then
  important "Git identity incomplete. Set it once on your host machine:" \
    "" \
    "  ${C_BLD}git config --global user.name 'Your Name'${C_RST}" \
    "  ${C_BLD}git config --global user.email 'you@example.com'${C_RST}" \
    "" \
    "Then rebuild the container and it will be picked up automatically."
fi

echo "${C_GRN}✓${C_RST} Git configuration complete"
