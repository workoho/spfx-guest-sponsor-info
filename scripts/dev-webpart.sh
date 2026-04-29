#!/usr/bin/env bash
# Copyright 2026 Workoho GmbH <https://workoho.com>
# Author: Julian Pawlowski <https://github.com/jpawlowski>
# Licensed under PolyForm Shield License 1.0.0 <https://polyformproject.org/licenses/shield/1.0.0>
#
# Start the SPFx web part local development server (hot-reload dev mode).
#
# Usage:
#   scripts/dev-webpart.sh
#
# Requires SPFX_SERVE_TENANT_DOMAIN to be set. Two ways to provide it
# (in order of preference):
#
#   1. Host OS environment variable (persistent across container rebuilds):
#        export SPFX_SERVE_TENANT_DOMAIN=contoso.sharepoint.com
#      Add this to ~/.bashrc or ~/.profile on your host machine. The
#      devcontainer picks it up automatically via containerEnv.
#
#   2. Local .env file (git-ignored, persistent within the container):
#        cp .env.example .env
#      Then set SPFX_SERVE_TENANT_DOMAIN=<your-tenant>.sharepoint.com in .env.
#
# NOTE: The local workbench (/temp/workbench.html) was removed in SPFx 1.17.
# The dev server only serves the JS bundle; testing requires the hosted
# workbench on a real SharePoint Online tenant.
#
# The hosted workbench URL is printed on startup:
#   https://<your-tenant>.sharepoint.com/_layouts/15/workbench.aspx
#
# Prerequisites: accept the dev certificate warning in your browser the first
# time by navigating to https://localhost:4321 and confirming the certificate.

set -euo pipefail

# Always run from the repository root so paths resolve correctly.
cd "$(dirname "${BASH_SOURCE[0]}")/.."

# shellcheck source=scripts/colors.sh
source "$(dirname "${BASH_SOURCE[0]}")/colors.sh"

# Load .env if present (overrides any containerEnv value for this session).
ENV_FILE=".env"
if [[ -f "${ENV_FILE}" ]]; then
  set -a
  # shellcheck source=/dev/null
  source "${ENV_FILE}"
  set +a
fi

SPFX_SERVE_TENANT_DOMAIN="${SPFX_SERVE_TENANT_DOMAIN:-}"

if [[ -z "${SPFX_SERVE_TENANT_DOMAIN}" ]]; then
  echo "${C_RED}ERROR:${C_RST} SPFX_SERVE_TENANT_DOMAIN is not set."
  important "The local workbench was removed in SPFx 1.17." \
    "A SharePoint Online tenant is required to test the web part." \
    "" \
    "${C_BLD}Option 1${C_RST} — this terminal session only (lost when the terminal closes):" \
    "  export SPFX_SERVE_TENANT_DOMAIN=contoso.sharepoint.com" \
    "" \
    "${C_BLD}Option 2${C_RST} — persistent in this container (all terminals, survives VS Code restarts):" \
    "  cp -n .env.example .env   # skip if .env already exists" \
    "  echo 'SPFX_SERVE_TENANT_DOMAIN=contoso.sharepoint.com' >> .env" \
    "  ${C_DIM}This script sources .env on every run, so it works with any shell.${C_RST}" \
    "" \
    "${C_BLD}Option 3${C_RST} — permanent across future container rebuilds (set once on your host OS):" \
    "  echo 'export SPFX_SERVE_TENANT_DOMAIN=contoso.sharepoint.com' >> ~/.bashrc" \
    "  ${C_DIM}Use ~/.zshrc for zsh. Takes effect the next time this container is rebuilt.${C_RST}"
  exit 1
fi

# --- Dependencies ---

if [[ ! -d "node_modules" ]]; then
  echo "Installing web part dependencies..."
  npm install
  echo ""
fi

echo "Tenant: ${SPFX_SERVE_TENANT_DOMAIN}"
echo "Starting local development server..."
hint "Hosted workbench:" \
  "  ${C_BLD}https://${SPFX_SERVE_TENANT_DOMAIN}/_layouts/15/workbench.aspx${C_RST}" \
  "" \
  "→ Accept the certificate at ${C_BLD}https://localhost:4321${C_RST} first (once per browser)"
echo "Press Ctrl+C to stop."
echo ""

# Node ≥17 resolves 'localhost' to ::1 (IPv6) by default, but the devcontainer
# port-forwarding tunnel binds to 127.0.0.1 (IPv4). Force IPv4-first DNS
# ordering so the dev server listens on 127.0.0.1:4321 where VS Code can reach it.
# This only affects this process and its children — no global IPv6 changes.
export NODE_OPTIONS="${NODE_OPTIONS:+${NODE_OPTIONS} }--dns-result-order=ipv4first"

npm start
