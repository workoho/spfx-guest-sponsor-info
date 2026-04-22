#!/usr/bin/env bash
# SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
# SPDX-FileCopyrightText: 2026 Julian Pawlowski <https://github.com/jpawlowski>
# SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0
#
# Post-provision hook for Azure Developer CLI (azd).
# Runs after Bicep deployment to grant the Function App's Managed Identity
# the required Microsoft Graph application roles:
#   - User.Read.All        (required; read any user's sponsors, profile, and photos)
#   - Presence.Read.All    (optional; requires Microsoft Teams)
#   - MailboxSettings.Read (optional; filters shared/room/equipment mailboxes)
#
# Role GUIDs are resolved dynamically from the Graph service principal so that
# no hardcoded IDs need to be maintained here.
#
# Bicep outputs (managedIdentityObjectId, sponsorApiUrl) are available
# as environment variables via 'azd env get-values' after provisioning.
#
# All operations are idempotent — safe to re-run.

set -euo pipefail

# shellcheck disable=SC1090  # process substitution: no static path to specify
source <(azd env get-values)

MANAGED_IDENTITY_OBJECT_ID="${MANAGED_IDENTITY_OBJECT_ID:?Bicep output MANAGED_IDENTITY_OBJECT_ID missing — did provisioning succeed?}"

GRAPH_APP_ID="00000003-0000-0000-c000-000000000000"

echo "Resolving Microsoft Graph service principal..."
GRAPH_SP_ID=$(az ad sp show --id "${GRAPH_APP_ID}" --query "id" -o tsv)

# Resolve a Graph app role ID by permission name (Application type only).
resolve_role_id() {
  az rest \
    --method GET \
    --url "https://graph.microsoft.com/v1.0/servicePrincipals/${GRAPH_SP_ID}/appRoles" \
    --query "value[?value=='${1}' && contains(allowedMemberTypes, 'Application')].id | [0]" \
    -o tsv 2>/dev/null || true
}

# Assign a Graph app role to the Managed Identity; skip if already assigned.
# Usage: assign_role <name> [optional=false]
assign_role() {
  local ROLE_NAME="${1}"
  local OPTIONAL="${2:-false}"
  local ROLE_ID
  ROLE_ID=$(resolve_role_id "${ROLE_NAME}")

  if [ -z "${ROLE_ID:-}" ]; then
    if [ "${OPTIONAL}" = "true" ]; then
      echo "  ⚠ ${ROLE_NAME} not found in this tenant — skipping (optional)."
      return
    else
      echo "  ✗ Required role ${ROLE_NAME} not found on the Graph service principal." >&2
      exit 1
    fi
  fi

  echo "Checking app role ${ROLE_NAME}..."
  EXISTING=$(az rest \
    --method GET \
    --url "https://graph.microsoft.com/v1.0/servicePrincipals/${MANAGED_IDENTITY_OBJECT_ID}/appRoleAssignments" \
    --query "value[?appRoleId=='${ROLE_ID}'].id | [0]" \
    -o tsv 2>/dev/null || true)

  if [ -n "${EXISTING:-}" ]; then
    echo "  ${ROLE_NAME} already assigned — skipping."
  else
    az rest \
      --method POST \
      --url "https://graph.microsoft.com/v1.0/servicePrincipals/${MANAGED_IDENTITY_OBJECT_ID}/appRoleAssignments" \
      --body "{\"principalId\":\"${MANAGED_IDENTITY_OBJECT_ID}\",\"resourceId\":\"${GRAPH_SP_ID}\",\"appRoleId\":\"${ROLE_ID}\"}" \
      >/dev/null
    echo "  ${ROLE_NAME} assigned."
  fi
}

assign_role "User.Read.All"
assign_role "Presence.Read.All" "true"
assign_role "MailboxSettings.Read" "true"

# ── Print web part configuration values ──────────────────────────────────────
echo ""
echo "Paste these values into the SPFx web part property pane"
echo "(Edit web part → Guest Sponsor API):"
echo ""
echo "  Guest Sponsor API Base URL              : ${SPONSOR_API_URL}"
echo "  Guest Sponsor API Client ID (App Reg.)  : ${AZURE_WEB_PART_CLIENT_ID}"
echo ""
echo "Note: Storage role assignment propagation can take 1–2 minutes."
echo "If the function returns errors immediately after deployment,"
echo "wait a moment and retry — no redeployment is needed."
