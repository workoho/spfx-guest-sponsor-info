#!/usr/bin/env bash
# SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
# SPDX-FileCopyrightText: 2026 Julian Pawlowski <https://github.com/jpawlowski>
# SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0
#
# Pre-provision hook for Azure Developer CLI (azd).
# Runs before Bicep deployment to:
#   1. Derive a default Function App name from the azd environment name.
#   2. Detect or prompt for the SharePoint tenant name.
#
# The Entra App Registration and Microsoft Graph permission assignments are
# now managed declaratively by the Bicep template (Microsoft Graph Bicep
# extension v1.0).  The deploying principal needs:
#   - Application.ReadWrite.All  (Cloud Application Administrator,
#                                  Application Administrator, or Global Administrator)
#   - AppRoleAssignment.ReadWrite.All  (Privileged Role Administrator
#                                        or Global Administrator)
#
# All operations are idempotent — safe to re-run on 'azd provision' or 'azd up'.

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
INVOCATION_DIR="$(pwd)"

find_azd_project_root() {
  local candidate

  if [[ -n "${AZD_PROJECT_PATH:-}" && -f "${AZD_PROJECT_PATH}/azure.yaml" ]]; then
    printf '%s\n' "${AZD_PROJECT_PATH}"
    return 0
  fi

  if [[ -f "${INVOCATION_DIR}/azure.yaml" ]]; then
    printf '%s\n' "${INVOCATION_DIR}"
    return 0
  fi

  candidate="${SCRIPT_DIR}"
  while [[ "${candidate}" != '/' ]]; do
    if [[ -f "${candidate}/azure.yaml" ]]; then
      printf '%s\n' "${candidate}"
      return 0
    fi
    candidate="$(dirname "${candidate}")"
  done

  return 1
}

if ! PROJECT_ROOT="$(find_azd_project_root)"; then
  echo 'ERROR: azure.yaml not found for azd pre-provision hook.' >&2
  exit 1
fi

cd "${PROJECT_ROOT}"

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
# Read from azd env — set by deploy-azure.ps1 via 'azd env set' before running provision.
# Fall back to main.parameters.json defaults when running azd directly without the wizard.
_ENV_VALUES="$(azd env get-values 2>/dev/null || true)"
HOSTING_PLAN="$(echo "${_ENV_VALUES}" | grep '^AZURE_HOSTING_PLAN=' | cut -d'=' -f2 | tr -d '"' || true)"
# Default to Consumption when the variable is absent (direct azd invocation).
HOSTING_PLAN="${HOSTING_PLAN:-Consumption}"
DEPLOY_AZURE_MAPS="$(echo "${_ENV_VALUES}" | grep '^AZURE_DEPLOY_AZURE_MAPS=' | cut -d'=' -f2 | tr -d '"' || true)"
DEPLOY_AZURE_MAPS="${DEPLOY_AZURE_MAPS:-true}"
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

# ── 1. Detect or prompt for SharePoint tenant name ──────────────────────────
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

if ! azd env get-values | grep -q '^AZURE_HOSTING_PLAN='; then
  azd env set AZURE_HOSTING_PLAN 'Consumption'
fi

if ! azd env get-values | grep -q '^AZURE_DEPLOY_AZURE_MAPS='; then
  azd env set AZURE_DEPLOY_AZURE_MAPS 'true'
fi

if ! azd env get-values | grep -q '^AZURE_TAG_ENVIRONMENT='; then
  azd env set AZURE_TAG_ENVIRONMENT ''
fi

if ! azd env get-values | grep -q '^AZURE_TAG_CRITICALITY='; then
  azd env set AZURE_TAG_CRITICALITY ''
fi

if ! azd env get-values | grep -q '^AZURE_APP_VERSION='; then
  azd env set AZURE_APP_VERSION 'latest'
fi

if ! azd env get-values | grep -q '^AZURE_ENABLE_MONITORING='; then
  azd env set AZURE_ENABLE_MONITORING 'true'
fi

if ! azd env get-values | grep -q '^AZURE_ENABLE_FAILURE_ANOMALIES_ALERT='; then
  azd env set AZURE_ENABLE_FAILURE_ANOMALIES_ALERT 'false'
fi

if ! azd env get-values | grep -q '^AZURE_ALWAYS_READY_INSTANCES='; then
  azd env set AZURE_ALWAYS_READY_INSTANCES '1'
fi

if ! azd env get-values | grep -q '^AZURE_MAXIMUM_FLEX_INSTANCES='; then
  azd env set AZURE_MAXIMUM_FLEX_INSTANCES '10'
fi

if ! azd env get-values | grep -q '^AZURE_INSTANCE_MEMORY_MB='; then
  azd env set AZURE_INSTANCE_MEMORY_MB '2048'
fi

# ── 3. Entra role check ──────────────────────────────────────────────────────
# Cloud Application Administrator (or Application Administrator / Global Admin)
# is always required — Bicep creates and manages the App Registration.
#
# Privileged Role Administrator (or Global Admin) is required for Graph app role
# assignments to the Managed Identity. If that role is not available, defer by
# setting AZURE_SKIP_GRAPH_ROLE_ASSIGNMENTS=true before running azd provision:
#
#   azd env set AZURE_SKIP_GRAPH_ROLE_ASSIGNMENTS true
#
# Then run setup-graph-permissions.ps1 after deployment with the
# managedIdentityObjectId Bicep output (azd env get-values).
echo ''
echo 'Checking Entra roles...'
ENTRA_ROLES="$(az rest \
  --method GET \
  --url "https://graph.microsoft.com/v1.0/me/transitiveMemberOf/microsoft.graph.directoryRole?\$select=displayName" \
  --query 'value[*].displayName' \
  -o tsv 2>/dev/null || true)"
SKIP_ROLE_ASSIGNMENTS="${AZURE_SKIP_GRAPH_ROLE_ASSIGNMENTS:-}"
if [[ -n "${ENTRA_ROLES:-}" ]]; then
  HAS_APP_REG_ROLE="$(echo "${ENTRA_ROLES}" | grep -E \
    '^(Cloud Application Administrator|Application Administrator|Global Administrator)$' | head -1 || true)"
  HAS_ASSIGNMENT_ROLE=''
  if [[ "${SKIP_ROLE_ASSIGNMENTS}" != 'true' ]]; then
    HAS_ASSIGNMENT_ROLE="$(echo "${ENTRA_ROLES}" | grep -E \
      '^(Privileged Role Administrator|Global Administrator)$' | head -1 || true)"
  fi
  if [[ -z "${HAS_APP_REG_ROLE:-}" ]]; then
    echo '  ! Missing: Cloud Application Administrator, Application Administrator,'
    echo '    or Global Administrator — required to create/update the App Registration.'
    echo '    Bicep will fail without this role. Activate via PIM before re-running:'
    echo '    https://entra.microsoft.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/aadRoles'
  fi
  if [[ "${SKIP_ROLE_ASSIGNMENTS}" != 'true' && -z "${HAS_ASSIGNMENT_ROLE:-}" ]]; then
    echo '  ! Missing: Privileged Role Administrator (or Global Administrator) —'
    echo '    needed to assign Graph app roles to the Managed Identity.'
    echo '    Either activate the role via PIM, or defer the assignments:'
    echo '      azd env set AZURE_SKIP_GRAPH_ROLE_ASSIGNMENTS true'
    echo '    Then run setup-graph-permissions.ps1 after deployment.'
    echo '    PIM → My roles → Entra roles:'
    echo '    https://entra.microsoft.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/aadRoles'
  fi
  if [[ -n "${HAS_APP_REG_ROLE:-}" ]]; then
    if [[ "${SKIP_ROLE_ASSIGNMENTS}" == 'true' ]]; then
      echo "  ✓ Entra role: ${HAS_APP_REG_ROLE} — App Registration management covered."
      echo '    Graph role assignments: deferred to setup-graph-permissions.ps1.'
    elif [[ -n "${HAS_ASSIGNMENT_ROLE:-}" ]]; then
      if [[ "${HAS_APP_REG_ROLE}" == "${HAS_ASSIGNMENT_ROLE}" ]]; then
        echo "  ✓ Entra role: ${HAS_APP_REG_ROLE} — covers both required permissions."
      else
        echo "  ✓ Entra roles: ${HAS_APP_REG_ROLE} + ${HAS_ASSIGNMENT_ROLE} — both required roles active."
      fi
    fi
  fi
else
  echo '  ! Entra role check could not be completed — continuing anyway.'
  echo '    Required: Cloud Application Administrator (or similar).'
  if [[ "${SKIP_ROLE_ASSIGNMENTS}" != 'true' ]]; then
    echo '    Also required: Privileged Role Administrator (or Global Administrator).'
    echo '    To defer Graph role assignments: azd env set AZURE_SKIP_GRAPH_ROLE_ASSIGNMENTS true'
  fi
fi
echo ''
