#!/usr/bin/env bash
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
  echo "ERROR: SPFX_SERVE_TENANT_DOMAIN is not set."
  echo "  The local workbench was removed in SPFx 1.17."
  echo "  A SharePoint Online tenant is required to test the web part."
  echo ""
  echo "  Option 1 — this terminal session only (lost when the terminal closes):"
  echo "    export SPFX_SERVE_TENANT_DOMAIN=contoso.sharepoint.com"
  echo ""
  echo "  Option 2 — persistent in this container (all terminals, survives VS Code restarts):"
  echo "    cp -n .env.example .env   # skip if .env already exists"
  echo "    echo 'SPFX_SERVE_TENANT_DOMAIN=contoso.sharepoint.com' >> .env"
  echo "    This script sources .env on every run, so it works with any shell."
  echo ""
  echo "  Option 3 — permanent across future container rebuilds (set once on your host OS):"
  echo "    echo 'export SPFX_SERVE_TENANT_DOMAIN=contoso.sharepoint.com' >> ~/.bashrc"
  echo "    Use ~/.zshrc for zsh. Takes effect the next time this container is rebuilt."
  exit 1
fi

echo "Tenant: ${SPFX_SERVE_TENANT_DOMAIN}"
echo "Starting local development server..."
echo "Hosted workbench: https://${SPFX_SERVE_TENANT_DOMAIN}/_layouts/15/workbench.aspx"
echo "  → Accept the certificate at https://localhost:4321 first (once per browser)"
echo ""
echo "Press Ctrl+C to stop."
echo ""

# Node ≥17 resolves 'localhost' to ::1 (IPv6) by default, but the devcontainer
# port-forwarding tunnel binds to 127.0.0.1 (IPv4). Force IPv4-first DNS
# ordering so the dev server listens on 127.0.0.1:4321 where VS Code can reach it.
# This only affects this process and its children — no global IPv6 changes.
export NODE_OPTIONS="${NODE_OPTIONS:+${NODE_OPTIONS} }--dns-result-order=ipv4first"

npm start
