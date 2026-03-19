#!/usr/bin/env bash
# Pre-provision hook for Azure Developer CLI (azd).
# Runs before Bicep deployment to:
#   1. Derive a default Function App name from the azd environment name.
#   2. Detect or prompt for the SharePoint tenant name.
#   3. Create (or reuse) the Entra App Registration required for EasyAuth,
#      and store its client ID as AZURE_FUNCTION_CLIENT_ID in the azd environment.
#
# All operations are idempotent — safe to re-run on 'azd provision' or 'azd up'.

set -euo pipefail

APP_DISPLAY_NAME="Guest Sponsor Info Proxy"

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
    --query "value[0].verifiedDomains[?isDefault].name | [0]" \
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
echo "Checking for existing App Registration '${APP_DISPLAY_NAME}'..."
EXISTING_CLIENT_ID=$(az ad app list \
  --display-name "${APP_DISPLAY_NAME}" \
  --query "[0].appId" \
  -o tsv 2>/dev/null || true)

if [ -n "${EXISTING_CLIENT_ID:-}" ]; then
  echo "App Registration already exists. Client ID: ${EXISTING_CLIENT_ID}"
  CLIENT_ID="${EXISTING_CLIENT_ID}"
else
  echo "Creating App Registration '${APP_DISPLAY_NAME}'..."
  CLIENT_ID=$(az ad app create \
    --display-name "${APP_DISPLAY_NAME}" \
    --sign-in-audience "AzureADMyOrg" \
    --query "appId" \
    -o tsv)

  APP_ID_URI="api://guest-sponsor-info-proxy/${CLIENT_ID}"
  az ad app update \
    --id "${CLIENT_ID}" \
    --identifier-uris "${APP_ID_URI}"

  echo "App Registration created. App ID URI: ${APP_ID_URI}"
fi

azd env set AZURE_FUNCTION_CLIENT_ID "${CLIENT_ID}"
echo "AZURE_FUNCTION_CLIENT_ID set to ${CLIENT_ID}"
