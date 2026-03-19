#!/usr/bin/env bash
# Post-provision hook for Azure Developer CLI (azd).
# Runs after Bicep deployment to grant the Function App's Managed Identity
# the required Microsoft Graph application roles:
#   - User.Read.All     (read any user's sponsors, profile, and photos)
#   - Presence.Read.All (read sponsor presence status)
#
# Bicep outputs (managedIdentityObjectId, sponsorApiUrl) are available
# as environment variables via 'azd env get-values' after provisioning.
#
# All operations are idempotent — safe to re-run.

set -euo pipefail

source <(azd env get-values)

MANAGED_IDENTITY_OBJECT_ID="${MANAGED_IDENTITY_OBJECT_ID:?Bicep output MANAGED_IDENTITY_OBJECT_ID missing — did provisioning succeed?}"

GRAPH_APP_ID="00000003-0000-0000-c000-000000000000"
ROLE_USER_READ_ALL="df021288-bdef-4463-88db-98f22de89214"
ROLE_PRESENCE_READ_ALL="9c7a330d-35b3-4aa1-963d-cb2b055962cc"

echo "Resolving Microsoft Graph service principal..."
GRAPH_SP_ID=$(az ad sp show --id "${GRAPH_APP_ID}" --query "id" -o tsv)

for ROLE_ID in "${ROLE_USER_READ_ALL}" "${ROLE_PRESENCE_READ_ALL}"; do
  echo "Checking app role ${ROLE_ID}..."
  EXISTING=$(az rest \
    --method GET \
    --url "https://graph.microsoft.com/v1.0/servicePrincipals/${MANAGED_IDENTITY_OBJECT_ID}/appRoleAssignments" \
    --query "value[?appRoleId=='${ROLE_ID}'].id | [0]" \
    -o tsv 2>/dev/null || true)

  if [ -n "${EXISTING:-}" ]; then
    echo "  Role ${ROLE_ID} already assigned — skipping."
  else
    az rest \
      --method POST \
      --url "https://graph.microsoft.com/v1.0/servicePrincipals/${MANAGED_IDENTITY_OBJECT_ID}/appRoleAssignments" \
      --body "{\"principalId\":\"${MANAGED_IDENTITY_OBJECT_ID}\",\"resourceId\":\"${GRAPH_SP_ID}\",\"appRoleId\":\"${ROLE_ID}\"}" \
      > /dev/null
    echo "  Role ${ROLE_ID} assigned."
  fi
done

# ── Print web part configuration values ──────────────────────────────────────
echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "Deployment complete. Paste these values into the SPFx web part"
echo "property pane (Edit web part → Sponsor API configuration):"
echo ""
echo "  Sponsor API URL   : ${SPONSOR_API_URL}"
echo "  Function Client ID: ${AZURE_FUNCTION_CLIENT_ID}"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""
echo "Note: storage role assignment propagation can take 1–2 minutes."
echo "If the function returns errors immediately after deployment, wait"
echo "a moment and retry — no redeployment is needed."
