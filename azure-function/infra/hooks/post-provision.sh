#!/usr/bin/env bash
# SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
# SPDX-FileCopyrightText: 2026 Julian Pawlowski <https://github.com/jpawlowski>
# SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0
#
# Post-provision hook for Azure Developer CLI (azd).
# Runs after Bicep deployment to:
#   - Restart the Function App so the Managed Identity token cache picks up
#     the Graph application permissions that Bicep just assigned.
#   - Print the web part configuration values.
#
# The Entra App Registration and all Microsoft Graph application role
# assignments (User.Read.All, Presence.Read.All, MailboxSettings.Read,
# TeamMember.Read.All) are now managed by the Bicep template via the
# Microsoft Graph Bicep extension v1.0 — no manual permission grants here.
#
# Bicep outputs (functionAppUrl, webPartClientId) are available as environment
# variables via 'azd env get-values' after provisioning.
#
# All operations are idempotent — safe to re-run.

set -euo pipefail

# shellcheck disable=SC1090  # process substitution: no static path to specify
source <(azd env get-values)

# azd writes Bicep output names verbatim (camelCase) to the .env file and
# preloads them into the hook process environment with the same casing.
# Create SCREAMING_SNAKE_CASE aliases so the rest of this script uses a
# consistent naming convention alongside the AZURE_* env vars.
FUNCTION_APP_URL="${functionAppUrl:-${FUNCTION_APP_URL:-}}"
if [[ -z "${FUNCTION_APP_URL:-}" && -n "${sponsorApiEndpointUrl:-}" ]]; then
  FUNCTION_APP_URL="$(printf '%s' "${sponsorApiEndpointUrl}" | sed 's#/api/getGuestSponsors$##')"
fi
if [[ -z "${FUNCTION_APP_URL:-}" && -n "${sponsorApiUrl:-}" ]]; then
  FUNCTION_APP_URL="$(printf '%s' "${sponsorApiUrl}" | sed 's#/api/getGuestSponsors$##')"
fi
WEB_PART_CLIENT_ID="${webPartClientId:-${WEB_PART_CLIENT_ID:-}}"
# functionAppName is now a Bicep output (camelCase); fall back to the azd env
# var for deployments that still have AZURE_FUNCTION_APP_NAME persisted.
FUNCTION_APP_NAME="${functionAppName:-${AZURE_FUNCTION_APP_NAME:-}}"

# azd can retain a stale webPartClientId in the env file. Resolve the EasyAuth
# App Registration directly by its deterministic uniqueName and sync the azd
# environment so both this hook and deploy-azure.ps1 print the real client ID.
if [[ -n "${FUNCTION_APP_NAME:-}" ]]; then
  app_reg_unique_name="guest-sponsor-info-proxy-${FUNCTION_APP_NAME}"
  if resolved_client_id="$(az ad app list --filter "uniqueName eq '${app_reg_unique_name}'" --query '[0].appId' -o tsv 2>/dev/null)"; then
    if [[ -n "${resolved_client_id}" && "${resolved_client_id}" != "null" ]]; then
      existing_env_client_id="${webPartClientId:-}"
      WEB_PART_CLIENT_ID="${resolved_client_id}"
      export WEB_PART_CLIENT_ID
      export webPartClientId="${resolved_client_id}"
      if [[ "${AZURE_WEB_PART_CLIENT_ID:-}" != "${resolved_client_id}" ]]; then
        azd env set AZURE_WEB_PART_CLIENT_ID "${resolved_client_id}" >/dev/null
        export AZURE_WEB_PART_CLIENT_ID="${resolved_client_id}"
      fi
      if [[ "${existing_env_client_id}" != "${resolved_client_id}" ]]; then
        azd env set webPartClientId "${resolved_client_id}" >/dev/null
      fi
    fi
  fi
fi

# ── Restart Function App ──────────────────────────────────────────────────────
# Bicep assigns Graph app roles as part of the deployment.  A restart ensures
# the Managed Identity token cache is cleared and the new permissions are
# activated immediately.  Without this, the first invocations after a
# fresh deployment may fail until the token naturally expires.
if [ -n "${FUNCTION_APP_NAME:-}" ] && [ -n "${AZURE_RESOURCE_GROUP:-}" ]; then
  echo ""
  echo "Restarting Function App '${FUNCTION_APP_NAME}' to activate Graph permissions..."
  az functionapp restart \
    --name "${FUNCTION_APP_NAME}" \
    --resource-group "${AZURE_RESOURCE_GROUP}" \
    >/dev/null
  echo "  Function App restarted."
else
  echo ""
  echo "Note: Could not restart the Function App automatically"
  echo "(functionAppName output or AZURE_RESOURCE_GROUP not set)."
  echo "Restart it manually to ensure Graph permissions are activated."
fi

# ── Print web part configuration values ──────────────────────────────────────
echo ""
echo "Paste these values into the SPFx web part property pane"
echo "(Edit web part → Guest Sponsor API):"
echo ""
echo "  Guest Sponsor API Base URL              : ${FUNCTION_APP_URL}"
echo "  Guest Sponsor API Client ID (App Reg.)  : ${WEB_PART_CLIENT_ID}"
echo ""
echo "Note: Storage role assignment propagation can take 1–2 minutes."
echo "If the function returns errors immediately after deployment,"
echo "wait a moment and retry — no redeployment is needed."

# ── Deferred Graph permissions reminder ───────────────────────────────────────
# When deploy-azure.ps1 was used with SkipGraphRoleAssignments, the Bicep
# parameter skipGraphRoleAssignments=true was passed and
# AZURE_SKIP_GRAPH_ROLE_ASSIGNMENTS was written to the azd env.
# Remind the operator to run the follow-up script.
if [ "${AZURE_SKIP_GRAPH_ROLE_ASSIGNMENTS:-false}" = "true" ]; then
  echo ""
  echo "IMPORTANT: Graph role assignments are DEFERRED."
  echo "The Function App Managed Identity does not yet have the Microsoft Graph"
  echo "application permissions it needs. Run setup-graph-permissions.ps1 to assign them:"
  echo ""
  echo "  -ManagedIdentityObjectId : ${managedIdentityObjectId:-<see azd env get-values>}"
  echo "  -TenantId                : ${AZURE_TENANT_ID:-<see azd env get-values>}"
  echo ""
  echo "The web part will return errors until those permissions are assigned."
fi
