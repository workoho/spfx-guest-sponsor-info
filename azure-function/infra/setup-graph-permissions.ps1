<#
.SYNOPSIS
    Grants the Function App's Managed Identity the required Microsoft Graph application roles.

.DESCRIPTION
    After deploying the Azure Function, run this script to assign:
      - User.Read.All     (read any user's profile and sponsors)
      - Presence.Read.All (read sponsor presence status)

    The Managed Identity object ID is shown in the Azure Portal (Function App → Identity)
    and is also emitted as an output of the Bicep/ARM deployment.

.PARAMETER ManagedIdentityObjectId
    The object ID (not the client ID) of the Function App's system-assigned Managed Identity.

.PARAMETER TenantId
    The Entra tenant ID (GUID).

.EXAMPLE
    ./setup-graph-permissions.ps1 -ManagedIdentityObjectId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -TenantId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy"
#>
param(
    [Parameter(Mandatory)][string]$ManagedIdentityObjectId,
    [Parameter(Mandatory)][string]$TenantId
)

$ErrorActionPreference = 'Stop'

# Ensure Microsoft.Graph module is available.
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Write-Host "Installing Microsoft.Graph.Authentication module..." -ForegroundColor Cyan
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
}
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Applications)) {
    Write-Host "Installing Microsoft.Graph.Applications module..." -ForegroundColor Cyan
    Install-Module Microsoft.Graph.Applications -Scope CurrentUser -Force
}

Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Applications

Connect-MgGraph -TenantId $TenantId -Scopes "AppRoleAssignment.ReadWrite.All"

Write-Host "Resolving Microsoft Graph service principal..." -ForegroundColor Cyan
$graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
if (-not $graphSp) {
    throw "Could not find the Microsoft Graph service principal in tenant '$TenantId'."
}

# Resolve role IDs dynamically from the Graph service principal's app roles.
# This avoids hardcoded GUIDs and correctly detects unavailable permissions.
$requiredRoles = @(
    @{ Name = 'User.Read.All'; Optional = $false }
    @{ Name = 'Presence.Read.All'; Optional = $true }   # requires Microsoft Teams; function degrades gracefully without it
)

$assignedRoles = @()
$skippedRoles  = @()

foreach ($role in $requiredRoles) {
    Write-Host "Assigning $($role.Name) ..." -ForegroundColor Cyan

    $appRole = $graphSp.AppRoles | Where-Object { $_.Value -eq $role.Name -and $_.AllowedMemberTypes -contains 'Application' }
    if (-not $appRole) {
        if ($role.Optional) {
            Write-Host "  ⚠ $($role.Name) is not available as an Application permission in this tenant (Microsoft Teams may not be licensed). Skipping — sponsors will be shown without presence status." -ForegroundColor Yellow
            $skippedRoles += $role.Name
            continue
        } else {
            throw "Required permission '$($role.Name)' was not found on the Microsoft Graph service principal."
        }
    }

    try {
        $null = New-MgServicePrincipalAppRoleAssignment `
            -ServicePrincipalId $ManagedIdentityObjectId `
            -PrincipalId $ManagedIdentityObjectId `
            -ResourceId $graphSp.Id `
            -AppRoleId $appRole.Id `
            -ErrorAction Stop
        Write-Host "  ✓ $($role.Name) assigned." -ForegroundColor Green
        $assignedRoles += $role.Name
    } catch {
        if ($_.Exception.Message -like "*Permission being assigned already exists*") {
            Write-Host "  ✓ $($role.Name) already assigned — skipping." -ForegroundColor Yellow
            $assignedRoles += $role.Name
        } else {
            throw
        }
    }
}

Write-Host "`nDone. The Managed Identity can now call Microsoft Graph with:" -ForegroundColor Green
foreach ($r in $assignedRoles) {
    Write-Host "  - $r" -ForegroundColor Green
}
if ($skippedRoles.Count -gt 0) {
    Write-Host "`nSkipped (not available in this tenant):" -ForegroundColor Yellow
    foreach ($r in $skippedRoles) {
        Write-Host "  - $r" -ForegroundColor Yellow
    }
}
