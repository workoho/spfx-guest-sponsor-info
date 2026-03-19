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

$roleAssignments = @(
    @{
        Id   = 'df021288-bdef-4463-88db-98f22de89214'
        Name = 'User.Read.All'
    },
    @{
        Id   = '9c7a330d-35b3-4aa1-963d-cb2b055962cc'
        Name = 'Presence.Read.All'
    }
)

foreach ($role in $roleAssignments) {
    Write-Host "Assigning $($role.Name) ..." -ForegroundColor Cyan
    try {
        New-MgServicePrincipalAppRoleAssignment `
            -ServicePrincipalId $ManagedIdentityObjectId `
            -PrincipalId $ManagedIdentityObjectId `
            -ResourceId $graphSp.Id `
            -AppRoleId $role.Id | Out-Null
        Write-Host "  ✓ $($role.Name) assigned." -ForegroundColor Green
    } catch {
        if ($_.Exception.Message -like "*Permission being assigned already exists*") {
            Write-Host "  ✓ $($role.Name) already assigned — skipping." -ForegroundColor Yellow
        } else {
            throw
        }
    }
}

Write-Host "`nDone. The Managed Identity can now call Microsoft Graph with:" -ForegroundColor Green
Write-Host "  - User.Read.All" -ForegroundColor Green
Write-Host "  - Presence.Read.All" -ForegroundColor Green
