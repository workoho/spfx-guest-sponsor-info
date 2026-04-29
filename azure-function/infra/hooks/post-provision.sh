#!/usr/bin/env bash
# SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
# SPDX-FileCopyrightText: 2026 Julian Pawlowski <https://github.com/jpawlowski>
# SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0
#
# Post-provision hook for Azure Developer CLI (azd).
# Runs after Bicep deployment to:
#   - Restart the Function App when this deployment run managed Microsoft Graph
#     permissions and the Managed Identity token cache should be refreshed.
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

sync_azd_env_value() {
  local name="$1" value="$2"
  if [[ -z "${value}" ]]; then
    return
  fi
  export "${name}=${value}"
  azd env set "${name}" "${value}" >/dev/null
}

print_summary_line() {
  local label="$1" value="${2:-}"
  if [[ -z "${value}" ]]; then
    value="(not available)"
  fi
  printf '  %-28s: %s\n' "${label}" "${value}"
}

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
MANAGED_IDENTITY_OBJECT_ID="${managedIdentityObjectId:-${MANAGED_IDENTITY_OBJECT_ID:-}}"

if [[ -z "${FUNCTION_APP_URL:-}" && -n "${FUNCTION_APP_NAME:-}" && -n "${AZURE_RESOURCE_GROUP:-}" ]]; then
  if default_host_name="$(az functionapp show --name "${FUNCTION_APP_NAME}" --resource-group "${AZURE_RESOURCE_GROUP}" --query defaultHostName -o tsv 2>/dev/null)"; then
    if [[ -n "${default_host_name}" && "${default_host_name}" != "null" ]]; then
      FUNCTION_APP_URL="https://${default_host_name}"
      export FUNCTION_APP_URL
      export functionAppUrl="${FUNCTION_APP_URL}"
      sync_azd_env_value functionAppUrl "${FUNCTION_APP_URL}"
    fi
  fi
fi

if [[ -z "${MANAGED_IDENTITY_OBJECT_ID:-}" && -n "${FUNCTION_APP_NAME:-}" && -n "${AZURE_RESOURCE_GROUP:-}" ]]; then
  if principal_id="$(az functionapp identity show --name "${FUNCTION_APP_NAME}" --resource-group "${AZURE_RESOURCE_GROUP}" --query principalId -o tsv 2>/dev/null)"; then
    if [[ -n "${principal_id}" && "${principal_id}" != "null" ]]; then
      MANAGED_IDENTITY_OBJECT_ID="${principal_id}"
      export MANAGED_IDENTITY_OBJECT_ID
      export managedIdentityObjectId="${principal_id}"
      sync_azd_env_value managedIdentityObjectId "${principal_id}"
    fi
  fi
fi

# azd can retain a stale webPartClientId in the env file. Resolve the EasyAuth
# App Registration directly by its deterministic uniqueName and sync the azd
# environment so both this hook and deploy-azure.ps1 print the real client ID.
if [[ -n "${FUNCTION_APP_NAME:-}" ]]; then
  app_reg_unique_name="guest-sponsor-info-proxy-${FUNCTION_APP_NAME}"
  if resolved_client_id="$(az ad app list --filter "uniqueName eq '${app_reg_unique_name}'" --query '[0].appId' -o tsv 2>/dev/null)"; then
    if [[ -n "${resolved_client_id}" && "${resolved_client_id}" != "null" ]]; then
      WEB_PART_CLIENT_ID="${resolved_client_id}"
      export WEB_PART_CLIENT_ID
      export webPartClientId="${resolved_client_id}"
      export AZURE_WEB_PART_CLIENT_ID="${resolved_client_id}"
      sync_azd_env_value AZURE_WEB_PART_CLIENT_ID "${resolved_client_id}"
      sync_azd_env_value webPartClientId "${resolved_client_id}"
    fi
  fi
fi

# ── Restart Function App ──────────────────────────────────────────────────────
# When Microsoft Graph permissions are managed as part of this deployment run,
# a restart ensures the Managed Identity token cache is cleared and fresh roles
# are picked up immediately. In deferred mode the permissions are managed
# separately, so an automatic restart here would be unnecessary.
SKIP_ROLE_ASSIGNMENTS="${AZURE_SKIP_GRAPH_ROLE_ASSIGNMENTS:-false}"
RESTART_STATUS="restart manually if needed"
if [ "${SKIP_ROLE_ASSIGNMENTS}" = "true" ]; then
  RESTART_STATUS="not run in this deployment mode"
  echo ""
  echo "Skipping automatic Function App restart because Microsoft Graph permissions are managed separately in this deployment mode."
elif [ -n "${FUNCTION_APP_NAME:-}" ] && [ -n "${AZURE_RESOURCE_GROUP:-}" ]; then
  echo ""
  echo "Restarting Function App '${FUNCTION_APP_NAME}' to activate Graph permissions..."
  az functionapp restart \
    --name "${FUNCTION_APP_NAME}" \
    --resource-group "${AZURE_RESOURCE_GROUP}" \
    >/dev/null
  RESTART_STATUS="completed"
else
  echo ""
  echo "Skipping automatic Function App restart (function app name or resource group missing)."
fi

# ── Print concise post-provision summary ─────────────────────────────────────
echo ""
echo "Post-provision summary"
echo "----------------------"
print_summary_line "Function app restart" "${RESTART_STATUS}"
print_summary_line "Guest Sponsor API Base URL" "${FUNCTION_APP_URL}"
print_summary_line "Guest Sponsor API Client ID" "${WEB_PART_CLIENT_ID}"

# ── Deferred Graph permissions reminder ───────────────────────────────────────
# When deploy-azure.ps1 was used with SkipGraphRoleAssignments, the Bicep
# parameter skipGraphRoleAssignments=true was passed and
# AZURE_SKIP_GRAPH_ROLE_ASSIGNMENTS was written to the azd env.
# Remind the operator to run the follow-up script.
if [ "${SKIP_ROLE_ASSIGNMENTS}" = "true" ]; then
  print_summary_line "Microsoft Graph permissions" "managed separately in this mode"
  print_summary_line "Managed identity object ID" "${MANAGED_IDENTITY_OBJECT_ID}"
  print_summary_line "TenantId" "${AZURE_TENANT_ID:-}"
  print_summary_line "If needed, run this script" "setup-graph-permissions.ps1"
  print_summary_line "When to run it" "if the web part shows permission errors"
fi
echo ""
echo "Note: Storage role assignment propagation can take 1-2 minutes."
if [ "${SKIP_ROLE_ASSIGNMENTS}" = "true" ]; then
  echo "If the web part shows permission errors, run setup-graph-permissions.ps1 and then restart the Function App once."
else
  echo "If you see errors right after deployment, wait a moment and try again."
fi
