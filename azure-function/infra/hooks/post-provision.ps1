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

function Sync-AzdEnvValue {
  param(
    [Parameter(Mandatory)][string]$Name,
    [AllowEmptyString()][string]$Value
  )

  if ([string]::IsNullOrWhiteSpace($Value)) {
    return
  }

  [Environment]::SetEnvironmentVariable($Name, $Value)
  azd env set $Name $Value | Out-Null
}

function Write-SummaryLine {
  param(
    [Parameter(Mandatory)][string]$Label,
    [AllowEmptyString()][string]$Value
  )

  $_displayValue = if ([string]::IsNullOrWhiteSpace($Value)) { '(not available)' } else { $Value }
  Write-Host ('  {0,-28}: {1}' -f $Label, $_displayValue)
}

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
if (-not $env:MANAGED_IDENTITY_OBJECT_ID) {
  $env:MANAGED_IDENTITY_OBJECT_ID = $env:managedIdentityObjectId
}

if (-not $env:FUNCTION_APP_URL -and $env:FUNCTION_APP_NAME -and $env:AZURE_RESOURCE_GROUP) {
  try {
    $_defaultHostName = (az functionapp show --name $env:FUNCTION_APP_NAME --resource-group $env:AZURE_RESOURCE_GROUP --query defaultHostName -o tsv 2>$null).Trim()
    if ($_defaultHostName -and $_defaultHostName -ne 'null') {
      $env:FUNCTION_APP_URL = "https://$_defaultHostName"
      $env:functionAppUrl = $env:FUNCTION_APP_URL
      Sync-AzdEnvValue -Name 'functionAppUrl' -Value $env:FUNCTION_APP_URL
    }
  }
  catch {
    Write-Verbose "Could not resolve Function App base URL from Azure: $_"
  }
}

if (-not $env:MANAGED_IDENTITY_OBJECT_ID -and $env:FUNCTION_APP_NAME -and $env:AZURE_RESOURCE_GROUP) {
  try {
    $_principalId = (az functionapp identity show --name $env:FUNCTION_APP_NAME --resource-group $env:AZURE_RESOURCE_GROUP --query principalId -o tsv 2>$null).Trim()
    if ($_principalId -and $_principalId -ne 'null') {
      $env:MANAGED_IDENTITY_OBJECT_ID = $_principalId
      $env:managedIdentityObjectId = $_principalId
      Sync-AzdEnvValue -Name 'managedIdentityObjectId' -Value $_principalId
    }
  }
  catch {
    Write-Verbose "Could not resolve Managed Identity object ID from Azure: $_"
  }
}

# azd can retain a stale webPartClientId in the env file. Resolve the EasyAuth
# App Registration directly by its deterministic uniqueName and sync the azd
# environment so both this hook and deploy-azure.ps1 print the real client ID.
if ($env:FUNCTION_APP_NAME) {
  try {
    $_appRegUniqueName = "guest-sponsor-info-proxy-$($env:FUNCTION_APP_NAME)"
    $_resolvedClientId = (az ad app list --filter "uniqueName eq '$_appRegUniqueName'" --query '[0].appId' -o tsv 2>$null).Trim()
    if ($_resolvedClientId -and $_resolvedClientId -ne 'null') {
      $env:WEB_PART_CLIENT_ID = $_resolvedClientId
      $env:webPartClientId = $_resolvedClientId
      $env:AZURE_WEB_PART_CLIENT_ID = $_resolvedClientId
      Sync-AzdEnvValue -Name 'AZURE_WEB_PART_CLIENT_ID' -Value $_resolvedClientId
      Sync-AzdEnvValue -Name 'webPartClientId' -Value $_resolvedClientId
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
$restartStatus = 'restart manually if needed'
if ($functionAppName -and $resourceGroup) {
  Write-Host ''
  Write-Host "Restarting Function App '$functionAppName' to activate Graph permissions..."
  az functionapp restart --name $functionAppName --resource-group $resourceGroup | Out-Null
  $restartStatus = 'completed'
}
else {
  Write-Host ''
  Write-Host 'Skipping automatic Function App restart (function app name or resource group missing).'
}

# ── Print concise post-provision summary ─────────────────────────────────────
Write-Host ''
Write-Host 'Post-provision summary'
Write-Host '----------------------' -ForegroundColor DarkGray
Write-SummaryLine -Label 'Function app restart' -Value $restartStatus
Write-SummaryLine -Label 'Guest Sponsor API Base URL' -Value $env:FUNCTION_APP_URL
Write-SummaryLine -Label 'Guest Sponsor API Client ID' -Value $env:WEB_PART_CLIENT_ID

# ── Deferred Graph permissions reminder ───────────────────────────────────────
# When deploy-azure.ps1 was used with SkipGraphRoleAssignments, the Bicep
# parameter skipGraphRoleAssignments=true was passed and AZURE_SKIP_GRAPH_ROLE_ASSIGNMENTS
# was written to the azd env. Remind the operator to run the follow-up script.
$_skipRoles = $env:AZURE_SKIP_GRAPH_ROLE_ASSIGNMENTS -eq 'true'
if ($_skipRoles) {
  Write-SummaryLine -Label 'Microsoft Graph permissions' -Value 'one more admin step is needed'
  Write-SummaryLine -Label 'Managed identity object ID' -Value $env:MANAGED_IDENTITY_OBJECT_ID
  Write-SummaryLine -Label 'TenantId' -Value $env:AZURE_TENANT_ID
  Write-SummaryLine -Label 'Next step' -Value 'run setup-graph-permissions.ps1 to finish Microsoft Graph permissions'
}
Write-Host ''
Write-Host 'Note: Storage role assignment propagation can take 1-2 minutes.'
if ($_skipRoles) {
  Write-Host 'The web part may show errors until you finish the Microsoft Graph permissions step.'
}
else {
  Write-Host 'If you see errors right after deployment, wait a moment and try again.'
}
