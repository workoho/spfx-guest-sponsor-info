#!/usr/bin/env pwsh
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
    @{ Id = 'df021288-bdef-4463-88db-98f22de89214'; Name = 'User.Read.All' },
    @{ Id = '9c7a330d-35b3-4aa1-963d-cb2b055962cc'; Name = 'Presence.Read.All' }
)

Write-Host "Resolving Microsoft Graph service principal..."
$graphSpId = az ad sp show --id $GRAPH_APP_ID --query 'id' -o tsv

foreach ($role in $roles) {
    Write-Host "Checking app role $($role.Name)..."
    $existing = az rest `
        --method GET `
        --url "https://graph.microsoft.com/v1.0/servicePrincipals/$miObjectId/appRoleAssignments" `
        --query "value[?appRoleId=='$($role.Id)'].id | [0]" `
        -o tsv 2>$null

    if ($existing) {
        Write-Host "  $($role.Name) already assigned — skipping."
    } else {
        az rest `
            --method POST `
            --url "https://graph.microsoft.com/v1.0/servicePrincipals/$miObjectId/appRoleAssignments" `
            --body "{`"principalId`":`"$miObjectId`",`"resourceId`":`"$graphSpId`",`"appRoleId`":`"$($role.Id)`"}" `
            | Out-Null
        Write-Host "  $($role.Name) assigned."
    }
}

# ── Print web part configuration values ──────────────────────────────────────
Write-Host ""
Write-Host ("━" * 67)
Write-Host "Deployment complete. Paste these values into the SPFx web part"
Write-Host "property pane (Edit web part → Sponsor API configuration):"
Write-Host ""
Write-Host "  Sponsor API URL   : $($env:SPONSOR_API_URL)"
Write-Host "  Function Client ID: $($env:AZURE_FUNCTION_CLIENT_ID)"
Write-Host ("━" * 67)
Write-Host ""
Write-Host "Note: storage role assignment propagation can take 1–2 minutes."
Write-Host "If the function returns errors immediately after deployment, wait"
Write-Host "a moment and retry — no redeployment is needed."
