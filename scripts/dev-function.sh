#!/usr/bin/env bash
# Copyright 2026 Workoho GmbH <https://workoho.com>
# Author: Julian Pawlowski <https://github.com/jpawlowski>
# Licensed under PolyForm Shield License 1.0.0 <https://polyformproject.org/licenses/shield/1.0.0>
#
# Start the Azure Function locally for development.
#
# Usage:
#   scripts/dev-function.sh
#
# Prerequisites:
#   - Azure Functions Core Tools (func) — pre-installed in the dev container
#   - Azure CLI (az login) — for Graph API credentials via DefaultAzureCredential
#   - local.settings.json in azure-function/ (created from template on first run)
#
# The function starts on http://localhost:7071 by default.
# EasyAuth is not active locally, so pass the caller OID via header:
#
#   curl http://localhost:7071/api/getGuestSponsors \
#     -H "X-Dev-User-OID: <guest-user-oid>"
#
# To connect the SPFx web part to this local function, set the Function URL
# in the web part property pane to "localhost:7071".
# Note: the web part enforces HTTPS, so either:
#   - use VS Code port forwarding (automatic in dev containers / Codespaces), or
#   - start with: scripts/dev-function.sh --useHttps

set -euo pipefail

SCRIPT_DIR="${BASH_SOURCE[0]%/*}"
ROOT_DIR="${SCRIPT_DIR}/.."
FUNC_DIR="${ROOT_DIR}/azure-function"

# shellcheck source=scripts/colors.sh
source "${SCRIPT_DIR}/colors.sh"

# Load .env if present (shares SPFX_SERVE_TENANT_DOMAIN etc. with dev-webpart.sh).
ENV_FILE="${ROOT_DIR}/.env"
if [[ -f "${ENV_FILE}" ]]; then
  set -a
  # shellcheck source=/dev/null
  source "${ENV_FILE}"
  set +a
fi

# --- Preflight checks ---

if ! command -v func &>/dev/null; then
  echo "${C_RED}ERROR:${C_RST} Azure Functions Core Tools (func) not found."
  echo "  Install: npm i -g azure-functions-core-tools@4 --unsafe-perm true"
  echo "  In the dev container it is pre-installed."
  exit 1
fi

if ! command -v az &>/dev/null; then
  echo "${C_YLW}WARNING:${C_RST} Azure CLI (az) not found."
  echo "  The function uses DefaultAzureCredential which falls back to"
  echo "  Azure CLI for local Graph API access. Install it or set"
  echo "  AZURE_CLIENT_ID / AZURE_TENANT_ID / AZURE_CLIENT_SECRET for"
  echo "  service-principal auth."
  echo ""
elif ! az account show &>/dev/null 2>&1; then
  echo "${C_YLW}WARNING:${C_RST} Azure CLI is not logged in."
  echo "  Run 'az login' so the function can call Microsoft Graph locally."
  echo "  The function uses DefaultAzureCredential which tries Azure CLI"
  echo "  credentials when Managed Identity is not available."
  echo ""
fi

# --- local.settings.json ---

SETTINGS_FILE="${FUNC_DIR}/local.settings.json"
SETTINGS_EXAMPLE="${FUNC_DIR}/local.settings.json.example"

if [[ ! -f "${SETTINGS_FILE}" ]]; then
  echo "Creating local.settings.json from template..."
  cp "${SETTINGS_EXAMPLE}" "${SETTINGS_FILE}"
  important "Edit ${C_BLD}azure-function/local.settings.json${C_RST}" \
    "" \
    "Required:" \
    "  ${C_BLD}TENANT_ID${C_RST}          — your Entra tenant ID (GUID)" \
    "  ${C_BLD}ALLOWED_AUDIENCE${C_RST}   — client ID of the app registration" \
    "  ${C_BLD}CORS_ALLOWED_ORIGIN${C_RST} — https://<tenant>.sharepoint.com" \
    "" \
    "${C_DIM}The file is in .gitignore and will not be committed.${C_RST}"
fi

# --- Dependencies ---

if [[ ! -d "${FUNC_DIR}/node_modules" ]]; then
  echo "Installing Azure Function dependencies..."
  npm --prefix "${FUNC_DIR}" install
  echo ""
fi

# --- Build ---

echo "Building Azure Function..."
npm --prefix "${FUNC_DIR}" run build
echo ""

# --- Start ---

# Ensure NODE_ENV is not 'production' so X-Dev-User-OID header is accepted.
export NODE_ENV="${NODE_ENV:-development}"

# Node ≥17 resolves 'localhost' to ::1 (IPv6) by default, but the devcontainer
# port-forwarding tunnel binds to 127.0.0.1 (IPv4). Force IPv4-first DNS
# ordering so the function listens where VS Code can reach it.
export NODE_OPTIONS="${NODE_OPTIONS:+${NODE_OPTIONS} }--dns-result-order=ipv4first"

echo "Starting Azure Function..."
hint "Endpoint: ${C_BLD}http://localhost:7071/api/getGuestSponsors${C_RST}" \
  "" \
  "Test with:" \
  "  curl http://localhost:7071/api/getGuestSponsors \\" \
  "    -H 'X-Dev-User-OID: <guest-user-oid>'" \
  "" \
  "To connect the SPFx web part, set the Function URL property to:" \
  "  ${C_BLD}localhost:7071${C_RST}   (use VS Code port forwarding for HTTPS)"
echo "Press Ctrl+C to stop."
echo ""

cd "${FUNC_DIR}"
func start "$@"
