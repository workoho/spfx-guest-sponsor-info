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

# Load azd environment (includes Bicep outputs converted to SCREAMING_SNAKE_CASE).
foreach ($line in (azd env get-values)) {
  if ($line -match '^([A-Za-z_][A-Za-z0-9_]*)=(.*)$') {
    [Environment]::SetEnvironmentVariable($Matches[1], $Matches[2].Trim('"'))
  }
}

$miObjectId = $env:MANAGED_IDENTITY_OBJECT_ID
if (-not $miObjectId) {
  throw "Bicep output MANAGED_IDENTITY_OBJECT_ID is missing — did provisioning succeed?"
}

$GRAPH_APP_ID = '00000003-0000-0000-c000-000000000000'
$roles = @(
  @{ Name = 'User.Read.All'; Optional = $false }
  @{ Name = 'Presence.Read.All'; Optional = $true }   # requires Microsoft Teams
  @{ Name = 'MailboxSettings.Read'; Optional = $true }   # filters shared/room/equipment mailboxes
)

Write-Host "Resolving Microsoft Graph service principal..."
$graphSpId = az ad sp show --id $GRAPH_APP_ID --query 'id' -o tsv

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
