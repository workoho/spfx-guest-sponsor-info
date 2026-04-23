#!/usr/bin/env pwsh
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
if (-not $env:MANAGED_IDENTITY_OBJECT_ID) { $env:MANAGED_IDENTITY_OBJECT_ID = $env:managedIdentityObjectId }
if (-not $env:SPONSOR_API_URL) { $env:SPONSOR_API_URL = $env:sponsorApiUrl }

$miObjectId = $env:MANAGED_IDENTITY_OBJECT_ID
if (-not $miObjectId) {
  throw "Bicep output managedIdentityObjectId is missing — did provisioning succeed?"
}

$GRAPH_APP_ID = '00000003-0000-0000-c000-000000000000'
$roles = @(
  @{ Name = 'User.Read.All'; Optional = $false }
  @{ Name = 'Presence.Read.All'; Optional = $true }   # requires Microsoft Teams
  @{ Name = 'MailboxSettings.Read'; Optional = $true }   # filters shared/room/equipment mailboxes
)

Write-Host "Resolving Microsoft Graph service principal..."
$graphSpId = az ad sp show --id $GRAPH_APP_ID --query 'id' -o tsv

$newRolesAssigned = $false
foreach ($role in $roles) {
  # Resolve the app role ID dynamically by name — avoids hardcoded GUIDs.
  $roleId = az rest `
    --method GET `
    --url "https://graph.microsoft.com/v1.0/servicePrincipals/$graphSpId/appRoles" `
    --query "value[?value=='$($role.Name)' && contains(allowedMemberTypes, 'Application')].id | [0]" `
    -o tsv 2>$null

  if (-not $roleId) {
    if ($role.Optional) {
      Write-Host "  ⚠ $($role.Name) not found in this tenant — skipping (optional)."
      continue
    }
    else {
      throw "Required role $($role.Name) not found on the Microsoft Graph service principal."
    }
  }

  Write-Host "Checking app role $($role.Name)..."
  $existing = az rest `
    --method GET `
    --url "https://graph.microsoft.com/v1.0/servicePrincipals/$miObjectId/appRoleAssignments" `
    --query "value[?appRoleId=='$roleId'].id | [0]" `
    -o tsv 2>$null

  if ($existing) {
    Write-Host "  $($role.Name) already assigned — skipping."
  }
  else {
    az rest `
      --method POST `
      --url "https://graph.microsoft.com/v1.0/servicePrincipals/$miObjectId/appRoleAssignments" `
      --body "{`"principalId`":`"$miObjectId`",`"resourceId`":`"$graphSpId`",`"appRoleId`":`"$roleId`"}" `
    | Out-Null
    Write-Host "  $($role.Name) assigned."
    $newRolesAssigned = $true
  }
}

# ── Restart Function App if new permissions were granted ─────────────────────
# New Graph app role assignments are not picked up until the managed identity
# token cache is cleared — a restart is the fastest way to do that.
if ($newRolesAssigned) {
  $functionAppName = $env:AZURE_FUNCTION_APP_NAME
  $resourceGroup = $env:AZURE_RESOURCE_GROUP
  if ($functionAppName -and $resourceGroup) {
    Write-Host ''
    Write-Host "Restarting Function App '$functionAppName' to activate new Graph permissions..."
    az functionapp restart --name $functionAppName --resource-group $resourceGroup | Out-Null
    Write-Host '  Function App restarted.'
  }
  else {
    Write-Host ''
    Write-Host "Note: New Graph permissions were granted. Restart the Function App manually"
    Write-Host "to activate them (AZURE_FUNCTION_APP_NAME or AZURE_RESOURCE_GROUP not set)."
  }
}

# ── Print web part configuration values ──────────────────────────────────────
Write-Host ''
Write-Host 'Paste these values into the SPFx web part property pane'
Write-Host '(Edit web part → Guest Sponsor API):'
Write-Host ''
Write-Host "  Guest Sponsor API Base URL              : $($env:SPONSOR_API_URL)"
Write-Host "  Guest Sponsor API Client ID (App Reg.)  : $($env:AZURE_WEB_PART_CLIENT_ID)"
Write-Host ''
Write-Host 'Note: Storage role assignment propagation can take 1-2 minutes.'
Write-Host 'If the function returns errors immediately after deployment,'
Write-Host 'wait a moment and retry - no redeployment is needed.'
