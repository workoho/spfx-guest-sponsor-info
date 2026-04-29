#!/usr/bin/env pwsh
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

$ErrorActionPreference = 'Stop'

# Load azd environment. azd writes Bicep output names verbatim (camelCase) to
# the .env file and preloads them into the hook process with the same casing.
foreach ($line in (azd env get-values)) {
  if ($line -match '^([A-Za-z_][A-Za-z0-9_]*)=(.*)$') {
    [Environment]::SetEnvironmentVariable($Matches[1], $Matches[2].Trim('"'))
  }
}
# Create SCREAMING_SNAKE_CASE aliases for camelCase Bicep outputs so the rest of
# this script uses a consistent naming convention alongside the AZURE_* env vars.
if (-not $env:FUNCTION_APP_URL) {
  if ($env:functionAppUrl) {
    $env:FUNCTION_APP_URL = $env:functionAppUrl
  }
  elseif ($env:sponsorApiEndpointUrl) {
    $env:FUNCTION_APP_URL = $env:sponsorApiEndpointUrl -replace '/api/getGuestSponsors$', ''
  }
  elseif ($env:sponsorApiUrl) {
    $env:FUNCTION_APP_URL = $env:sponsorApiUrl -replace '/api/getGuestSponsors$', ''
  }
}
if (-not $env:WEB_PART_CLIENT_ID) { $env:WEB_PART_CLIENT_ID = $env:webPartClientId }
# functionAppName is now a Bicep output (camelCase); fall back to the azd env
# var for deployments that still have AZURE_FUNCTION_APP_NAME persisted.
if (-not $env:FUNCTION_APP_NAME) {
  $env:FUNCTION_APP_NAME = if ($env:functionAppName) { $env:functionAppName } else { $env:AZURE_FUNCTION_APP_NAME }
}

# azd can retain a stale webPartClientId in the env file. Resolve the EasyAuth
# App Registration directly by its deterministic uniqueName and sync the azd
# environment so both this hook and deploy-azure.ps1 print the real client ID.
if ($env:FUNCTION_APP_NAME) {
  try {
    $_appRegUniqueName = "guest-sponsor-info-proxy-$($env:FUNCTION_APP_NAME)"
    $_resolvedClientId = (az ad app list --filter "uniqueName eq '$_appRegUniqueName'" --query '[0].appId' -o tsv 2>$null).Trim()
    if ($_resolvedClientId -and $_resolvedClientId -ne 'null') {
      $_existingEnvClientId = $env:webPartClientId
      $_existingLegacyClientId = $env:AZURE_WEB_PART_CLIENT_ID
      $env:WEB_PART_CLIENT_ID = $_resolvedClientId
      $env:webPartClientId = $_resolvedClientId
      if ($env:AZURE_WEB_PART_CLIENT_ID -ne $_resolvedClientId) {
        azd env set AZURE_WEB_PART_CLIENT_ID $_resolvedClientId | Out-Null
        $env:AZURE_WEB_PART_CLIENT_ID = $_resolvedClientId
      }
      if ($_existingEnvClientId -ne $_resolvedClientId) {
        azd env set webPartClientId $_resolvedClientId | Out-Null
      }
    }
  }
  catch {
    Write-Verbose "Could not resolve EasyAuth App Registration client ID from Entra: $_"
  }
}

# ── Restart Function App ──────────────────────────────────────────────────────
# Bicep assigns Graph app roles as part of the deployment.  A restart ensures
# the Managed Identity token cache is cleared and the new permissions are
# activated immediately.  Without this, the first invocations after a
# fresh deployment may fail until the token naturally expires.
$functionAppName = $env:FUNCTION_APP_NAME
$resourceGroup = $env:AZURE_RESOURCE_GROUP
if ($functionAppName -and $resourceGroup) {
  Write-Host ''
  Write-Host "Restarting Function App '$functionAppName' to activate Graph permissions..."
  az functionapp restart --name $functionAppName --resource-group $resourceGroup | Out-Null
  Write-Host '  Function App restarted.'
}
else {
  Write-Host ''
  Write-Host 'Note: Could not restart the Function App automatically'
  Write-Host '(functionAppName output or AZURE_RESOURCE_GROUP not set).'
  Write-Host 'Restart it manually to ensure Graph permissions are activated.'
}

# ── Print web part configuration values ──────────────────────────────────────
Write-Host ''
Write-Host 'Paste these values into the SPFx web part property pane'
Write-Host '(Edit web part → Guest Sponsor API):'
Write-Host ''
Write-Host "  Guest Sponsor API Base URL              : $($env:FUNCTION_APP_URL)"
Write-Host "  Guest Sponsor API Client ID (App Reg.)  : $($env:WEB_PART_CLIENT_ID)"
Write-Host ''
Write-Host 'Note: Storage role assignment propagation can take 1-2 minutes.'
Write-Host 'If the function returns errors immediately after deployment,'
Write-Host 'wait a moment and retry - no redeployment is needed.'

# ── Deferred Graph permissions reminder ───────────────────────────────────────
# When deploy-azure.ps1 was used with SkipGraphRoleAssignments, the Bicep
# parameter skipGraphRoleAssignments=true was passed and AZURE_SKIP_GRAPH_ROLE_ASSIGNMENTS
# was written to the azd env. Remind the operator to run the follow-up script.
$_skipRoles = $env:AZURE_SKIP_GRAPH_ROLE_ASSIGNMENTS -eq 'true'
if ($_skipRoles) {
  Write-Host ''
  Write-Host 'IMPORTANT: Graph role assignments are DEFERRED.'
  Write-Host 'The Function App Managed Identity does not yet have the Microsoft Graph'
  Write-Host 'application permissions it needs. Run setup-graph-permissions.ps1 to assign them:'
  Write-Host ''
  Write-Host "  -ManagedIdentityObjectId : $($env:managedIdentityObjectId)"
  Write-Host "  -TenantId                : $($env:AZURE_TENANT_ID)"
  Write-Host ''
  Write-Host 'The web part will return errors until those permissions are assigned.'
}
