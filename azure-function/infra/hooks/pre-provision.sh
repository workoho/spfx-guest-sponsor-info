#!/usr/bin/env bash
# SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
# SPDX-FileCopyrightText: 2026 Julian Pawlowski <https://github.com/jpawlowski>
# SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0
#
# Pre-provision hook for Azure Developer CLI (azd).
# Runs before Bicep deployment to:
#   1. Derive a default Function App name from the azd environment name.
#   2. Detect or prompt for the SharePoint tenant name.
#   3. Create (or reuse) the Entra App Registration required for EasyAuth,
#      and store its client ID as AZURE_WEB_PART_CLIENT_ID in the azd environment.
#
# All operations are idempotent — safe to re-run on 'azd provision' or 'azd up'.

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/../../.." && pwd)"
cd "${REPO_ROOT}"

APP_DISPLAY_NAME="Guest Sponsor Info - SharePoint Web Part Auth"
APP_DESCRIPTION="EasyAuth identity provider for the \"Guest Sponsor Info\" SharePoint Online web part (SPFx). Authenticates requests from the web part to the Azure Function proxy, which calls Microsoft Graph on behalf of signed-in guest users to retrieve their Entra sponsor information. Tokens are acquired silently via pre-authorized SharePoint Online Web Client Extensibility. Source: https://github.com/workoho/spfx-guest-sponsor-info"

# ── 0a. Check Azure RBAC permission ─────────────────────────────────────────
# Contributor (or Owner) on the subscription is needed to register resource
# providers and to deploy Bicep resources.  The check is informational — a
# missing role does not abort the script, but it surfaces the gap early so
# the operator can activate a PIM role or request access before the actual
# deployment runs.
echo ''
echo 'Checking Azure role assignment...'
# Use the env var set by azd (from the .env file) with fallback to parsing azd env.
SUBSCRIPTION_ID="${AZURE_SUBSCRIPTION_ID:-$(azd env get-values 2>/dev/null | grep '^AZURE_SUBSCRIPTION_ID=' | cut -d'=' -f2 | tr -d '"' || true)}"
if [[ -n "${SUBSCRIPTION_ID:-}" ]]; then
  USER_ID="$(az ad signed-in-user show --query id -o tsv 2>/dev/null || true)"
  if [[ -n "${USER_ID:-}" ]]; then
    RBAC_ROLES="$(az role assignment list \
      --scope "/subscriptions/${SUBSCRIPTION_ID}" \
      --assignee "${USER_ID}" \
      --include-inherited \
      --query "[?contains(['Owner','Contributor'], roleDefinitionName)].roleDefinitionName" \
      -o tsv 2>/dev/null || true)"
    if [[ -n "${RBAC_ROLES:-}" ]]; then
      # Collapse newlines to a comma-separated list for display.
      RBAC_LIST="$(echo "${RBAC_ROLES}" | sort -u | tr '\n' ',' | sed 's/,$//' | sed 's/,/, /g')"
      echo "  ✓ Azure RBAC: ${RBAC_LIST} on subscription."
    else
      echo '  ! Azure RBAC: no Contributor or Owner role found on this subscription.'
      echo '    Both are required for resource provider registration and Bicep deployment.'
      echo '    Contact your subscription owner to request Contributor access or activate'
      echo '    an eligible role via Azure PIM before re-running azd provision.'
      echo '    Azure PIM: https://portal.azure.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/azurerbac'
    fi
  else
    echo '  ! Azure RBAC: could not identify the signed-in user — skipping check.'
    echo '    Required: Contributor or Owner on the subscription.'
  fi
else
  echo '  ! Azure RBAC: AZURE_SUBSCRIPTION_ID not yet set — skipping role check.'
  echo '    Required: Contributor or Owner on the subscription.'
fi
echo ''

# ── 0. Validate required Azure resource providers ───────────────────────────
# Keep these defaults aligned with azure-function/infra/main.parameters.json.
HOSTING_PLAN='Consumption'
DEPLOY_AZURE_MAPS='true'
REQUIRED_PROVIDERS=(
  'Microsoft.AlertsManagement'
  'Microsoft.Authorization'
  'Microsoft.Insights'
  'Microsoft.ManagedIdentity'
  'Microsoft.OperationalInsights'
  'Microsoft.Resources'
  'Microsoft.Storage'
  'Microsoft.Web'
)

if [[ "${HOSTING_PLAN}" == 'FlexConsumption' ]]; then
  REQUIRED_PROVIDERS+=(
    'Microsoft.ContainerInstance'
  )
fi

if [[ "${DEPLOY_AZURE_MAPS,,}" == 'true' ]]; then
  REQUIRED_PROVIDERS+=(
    'Microsoft.Maps'
  )
fi

mapfile -t REQUIRED_PROVIDERS < <(printf '%s\n' "${REQUIRED_PROVIDERS[@]}" | sort -u)

echo 'Checking required Azure resource providers...'
MISSING_PROVIDERS=()
for provider in "${REQUIRED_PROVIDERS[@]}"; do
  state="$(az provider show --namespace "${provider}" --query registrationState -o tsv 2>/dev/null || true)"

  case "${state}" in
    Registered)
      echo "  ✓ ${provider} is registered."
      ;;
    Registering)
      echo "  ! ${provider} is still registering. Deployment can usually continue."
      ;;
    NotRegistered | Unregistered | '')
      echo "  ! ${provider} is not registered."
      MISSING_PROVIDERS+=("${provider}")
      ;;
    *)
      echo "  ! ${provider} returned state: ${state}"
      MISSING_PROVIDERS+=("${provider}")
      ;;
  esac
done

if [[ ${#MISSING_PROVIDERS[@]} -gt 0 ]]; then
  echo 'Registering missing Azure resource providers...'
  for provider in "${MISSING_PROVIDERS[@]}"; do
    echo "  -> az provider register --namespace ${provider} --wait"
    if az provider register --namespace "${provider}" --wait >/dev/null; then
      echo "  ✓ ${provider} registered."
    else
      echo "ERROR: Could not register ${provider}." >&2
      echo 'This usually means your account lacks subscription-level register permission.' >&2
      echo 'Minimum built-in role: Contributor. Owner also works.' >&2
      exit 1
    fi
  done
else
  echo '  ✓ All required resource providers are ready.'
fi

# ── 1. Derive a default Function App name ────────────────────────────────────
if ! azd env get-values | grep -q "^AZURE_FUNCTION_APP_NAME="; then
  ENV_NAME=$(azd env get-values | grep "^AZURE_ENV_NAME=" | cut -d'=' -f2 | tr -d '"')
  DEFAULT_APP_NAME="guest-sponsor-${ENV_NAME}"
  echo "Function App name not set — using: ${DEFAULT_APP_NAME}"
  azd env set AZURE_FUNCTION_APP_NAME "${DEFAULT_APP_NAME}"
fi

# ── 2. Detect or prompt for SharePoint tenant name ───────────────────────────
if ! azd env get-values | grep -q "^AZURE_SHAREPOINT_TENANT_NAME="; then
  # Try to derive from the default verified domain (e.g. contoso.onmicrosoft.com → contoso).
  DERIVED=$(az rest \
    --method GET \
    --url "https://graph.microsoft.com/v1.0/organization?\$select=verifiedDomains" \
    --query "value[0].verifiedDomains[?isInitial].name | [0]" \
    -o tsv 2>/dev/null | sed 's/\.onmicrosoft\.com//' || true)

  if [ -n "${DERIVED:-}" ]; then
    echo "Detected SharePoint tenant name: ${DERIVED}"
    azd env set AZURE_SHAREPOINT_TENANT_NAME "${DERIVED}"
  else
    read -rp "Enter your SharePoint tenant name (e.g. 'contoso' for contoso.sharepoint.com): " TENANT_NAME
    azd env set AZURE_SHAREPOINT_TENANT_NAME "${TENANT_NAME}"
  fi
fi

# ── 3. Create or reuse the App Registration ───────────────────────────────────
# ── 3a. Check Entra role for App Registration ────────────────────────────────
# Creating or modifying an App Registration requires an active Entra admin
# role.  'az rest' reuses the already-authenticated az CLI session to query
# the current user's directory role memberships.  The check is informational
# — a missing role does not abort the script, but it surfaces the gap early
# so the operator can activate an eligible PIM role before proceeding.
echo ''
echo 'Checking Entra role for App Registration...'
ENTRA_ROLES="$(az rest \
  --method GET \
  --url "https://graph.microsoft.com/v1.0/me/transitiveMemberOf/microsoft.graph.directoryRole?\$select=displayName" \
  --query 'value[*].displayName' \
  -o tsv 2>/dev/null || true)"
if [[ -n "${ENTRA_ROLES:-}" ]]; then
  # Filter to only the roles relevant to App Registration management.
  ACTIVE_ADMIN_ROLES="$(echo "${ENTRA_ROLES}" | grep -E \
    '^(Cloud Application Administrator|Application Administrator|Global Administrator)$' || true)"
  if [[ -n "${ACTIVE_ADMIN_ROLES:-}" ]]; then
    ACTIVE_LIST="$(echo "${ACTIVE_ADMIN_ROLES}" | tr '\n' ',' | sed 's/,$//' | sed 's/,/, /g')"
    echo "  ✓ Entra role: ${ACTIVE_LIST} — active."
  else
    echo '  ! Entra role: no required admin role is active for your account.'
    echo '    Required (one of):'
    echo '      Cloud Application Administrator'
    echo '      Application Administrator'
    echo '      Global Administrator'
    echo '    If your role is eligible (PIM): activate it before continuing.'
    echo '    PIM → My roles → Entra roles:'
    echo '    https://entra.microsoft.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/aadRoles'
    echo '    The App Registration step below will fail without one of these roles.'
  fi
else
  echo '  ! Entra role: check could not be completed — continuing anyway.'
  echo '    Required (one of): Cloud Application Administrator,'
  echo '    Application Administrator, or Global Administrator.'
fi
echo ''

echo "Checking for existing App Registration '${APP_DISPLAY_NAME}'..."

# Helper: show a clear, actionable error when an App Registration az command
# fails — commonly caused by a missing or inactive Entra admin role.
_app_reg_fail() {
  echo '' >&2
  echo 'ERROR: App Registration step failed.' >&2
  echo '  If the output above shows "Insufficient privileges" or "Forbidden",' >&2
  echo '  your account lacks the required Entra admin role. Required (one of):' >&2
  echo '    Cloud Application Administrator' >&2
  echo '    Application Administrator' >&2
  echo '    Global Administrator' >&2
  echo '' >&2
  echo '  If your role is eligible (PIM): activate it, then re-run azd provision.' >&2
  echo '  PIM → My roles → Entra roles:' >&2
  echo '  https://entra.microsoft.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/aadRoles' >&2
  exit 1
}

EXISTING_CLIENT_ID="$(az ad app list \
  --display-name "${APP_DISPLAY_NAME}" \
  --query "[0].appId" \
  -o tsv 2>/dev/null)" || _app_reg_fail

if [[ -n "${EXISTING_CLIENT_ID:-}" ]]; then
  echo "App Registration already exists. Client ID: ${EXISTING_CLIENT_ID}"
  CLIENT_ID="${EXISTING_CLIENT_ID}"
else
  echo "Creating App Registration '${APP_DISPLAY_NAME}'..."
  CLIENT_ID="$(az ad app create \
    --display-name "${APP_DISPLAY_NAME}" \
    --sign-in-audience "AzureADMyOrg" \
    --description "${APP_DESCRIPTION}" \
    --query "appId" \
    -o tsv)" || _app_reg_fail

  APP_ID_URI="api://guest-sponsor-info-proxy/${CLIENT_ID}"
  az ad app update \
    --id "${CLIENT_ID}" \
    --identifier-uris "${APP_ID_URI}" || _app_reg_fail

  echo "App Registration created. App ID URI: ${APP_ID_URI}"
fi

# Ensure accessTokenAcceptedVersion is set to 2 (v2 tokens — aud = bare clientId).
CURRENT_VERSION="$(az ad app show \
  --id "${CLIENT_ID}" \
  --query "api.requestedAccessTokenVersion" \
  -o tsv 2>/dev/null || true)"

if [[ "${CURRENT_VERSION:-}" != "2" ]]; then
  echo "Setting accessTokenAcceptedVersion to 2..."
  az rest --method PATCH \
    --url "https://graph.microsoft.com/v1.0/applications(appId='${CLIENT_ID}')" \
    --body '{"api":{"requestedAccessTokenVersion":2}}' || _app_reg_fail
fi

azd env set AZURE_WEB_PART_CLIENT_ID "${CLIENT_ID}"
echo "AZURE_WEB_PART_CLIENT_ID set to ${CLIENT_ID}"
