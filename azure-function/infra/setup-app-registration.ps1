<#
.SYNOPSIS
    Creates the Entra App Registration needed for the Azure Function proxy (EasyAuth).

.DESCRIPTION
    This script creates an App Registration named "Guest Sponsor Info Proxy" with:
      - Supported account types: single tenant
      - App ID URI: api://guest-sponsor-info-proxy/<clientId>

    The resulting Client ID must be provided as a parameter to the Bicep/ARM deployment.

.PARAMETER TenantId
    The Entra tenant ID (GUID).

.PARAMETER DisplayName
    Display name for the App Registration. Defaults to "Guest Sponsor Info Proxy".

.EXAMPLE
    ./setup-app-registration.ps1 -TenantId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy"
#>
param(
    [Parameter(Mandatory)][string]$TenantId,
    [string]$DisplayName = 'Guest Sponsor Info Proxy'
)

$ErrorActionPreference = 'Stop'

# Ensure Microsoft.Graph modules are available.
foreach ($module in @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Applications')) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installing $module module..." -ForegroundColor Cyan
        Install-Module $module -Scope CurrentUser -Force
    }
    Import-Module $module
}

Connect-MgGraph -TenantId $TenantId -Scopes "Application.ReadWrite.All"

Write-Host "Checking for existing App Registration '$DisplayName'..." -ForegroundColor Cyan
$existing = Get-MgApplication -Filter "displayName eq '$DisplayName'" -Top 1

if ($existing) {
    $clientId = $existing.AppId
    Write-Host "App Registration already exists. Client ID: $clientId" -ForegroundColor Yellow
} else {
    Write-Host "Creating App Registration '$DisplayName'..." -ForegroundColor Cyan

    $app = New-MgApplication -DisplayName $DisplayName `
        -SignInAudience 'AzureADMyOrg'

    $clientId = $app.AppId
    $objectId = $app.Id

    # Set the App ID URI.
    $appIdUri = "api://guest-sponsor-info-proxy/$clientId"
    Update-MgApplication -ApplicationId $objectId -IdentifierUris @($appIdUri)

    Write-Host "App Registration created:" -ForegroundColor Green
    Write-Host "  Display Name : $DisplayName" -ForegroundColor Green
    Write-Host "  Client ID    : $clientId" -ForegroundColor Green
    Write-Host "  App ID URI   : $appIdUri" -ForegroundColor Green
}

Write-Host "`n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "Copy this Client ID and use it as the 'functionClientId' parameter" -ForegroundColor Cyan
Write-Host "when deploying the ARM template, and in the SPFx web part property pane." -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Function Client ID: $clientId" -ForegroundColor White -BackgroundColor DarkGreen
Write-Host ""
