#!/usr/bin/env pwsh
# SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
# SPDX-FileCopyrightText: 2026 Julian Pawlowski <https://github.com/jpawlowski>
# SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0
#
# Pre-provision hook for Azure Developer CLI (azd).
# Runs before Bicep deployment to:
#   1. Derive a default Function App name from the azd environment name.
#   2. Detect or prompt for the SharePoint tenant name.
#   3. Create (or reuse) the Entra App Registration required for EasyAuth,
#      and store its client ID as AZURE_WEB_PART_CLIENT_ID in the azd environment.
#
# All operations are idempotent — safe to re-run on 'azd provision' or 'azd up'.

$ErrorActionPreference = 'Stop'

Set-Location -Path (Resolve-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath '../../..')).Path

$APP_DISPLAY_NAME = 'Guest Sponsor Info - SharePoint Web Part Auth'
$APP_DESCRIPTION = @(
  'EasyAuth identity provider for the "Guest Sponsor Info"',
  'SharePoint Online web part (SPFx). Authenticates requests from the',
  'web part to the Azure Function proxy, which calls Microsoft Graph on',
  'behalf of signed-in guest users to retrieve their Entra sponsor',
  'information. Tokens are acquired silently via pre-authorized',
  'SharePoint Online Web Client Extensibility.',
  'Source: https://github.com/workoho/spfx-guest-sponsor-info'
) -join ' '
$envValues = azd env get-values

# ── 0. Validate required Azure resource providers ───────────────────────────
# Keep these defaults aligned with azure-function/infra/main.parameters.json.
$hostingPlan = 'Consumption'
$deployAzureMaps = $true
$requiredProviders = @(
  'Microsoft.AlertsManagement',
  'Microsoft.Authorization',
  'Microsoft.Insights',
  'Microsoft.ManagedIdentity',
  'Microsoft.OperationalInsights',
  'Microsoft.Resources',
  'Microsoft.Storage',
  'Microsoft.Web'
)

if ($hostingPlan -eq 'FlexConsumption') {
  $requiredProviders += 'Microsoft.ContainerInstance'
}

if ($deployAzureMaps) {
  $requiredProviders += 'Microsoft.Maps'
}

$requiredProviders = $requiredProviders | Sort-Object -Unique
$missingProviders = @()

Write-Host 'Checking required Azure resource providers...'
foreach ($provider in $requiredProviders) {
  $state = az provider show --namespace $provider --query registrationState -o tsv 2>$null

  switch ($state) {
    'Registered' {
      Write-Host "  + $provider is registered."
    }
    'Registering' {
      Write-Host "  ! $provider is still registering. Deployment can usually continue."
    }
    'NotRegistered' {
      Write-Host "  ! $provider is not registered."
      $missingProviders += $provider
    }
    'Unregistered' {
      Write-Host "  ! $provider is not registered."
      $missingProviders += $provider
    }
    '' {
      Write-Host "  ! $provider returned no state."
      $missingProviders += $provider
    }
    default {
      Write-Host "  ! $provider returned state: $state"
      $missingProviders += $provider
    }
  }
}

if ($missingProviders.Count -gt 0) {
  Write-Host 'Registering missing Azure resource providers...'
  foreach ($provider in $missingProviders) {
    Write-Host "  -> az provider register --namespace $provider --wait"
    try {
      az provider register --namespace $provider --wait | Out-Null
      Write-Host "  + $provider registered."
    }
    catch {
      throw @(
        "Could not register $provider.",
        'This usually means your account lacks subscription-level register permission.',
        'Minimum built-in role: Contributor. Owner also works.'
      ) -join ' '
    }
  }
}
else {
  Write-Host '  + All required resource providers are ready.'
}

# ── 1. Derive a default Function App name ────────────────────────────────────
if ($envValues -notmatch 'AZURE_FUNCTION_APP_NAME=') {
  $envName = ($envValues | Select-String '^AZURE_ENV_NAME=(.+)$').Matches[0].Groups[1].Value.Trim('"')
  $defaultAppName = "guest-sponsor-$envName"
  Write-Host "Function App name not set — using: $defaultAppName"
  azd env set AZURE_FUNCTION_APP_NAME $defaultAppName
}

# ── 2. Detect or prompt for SharePoint tenant name ───────────────────────────
if ($envValues -notmatch 'AZURE_SHAREPOINT_TENANT_NAME=') {
  $derived = $null
  try {
    $raw = az rest `
      --method GET `
      --url 'https://graph.microsoft.com/v1.0/organization?$select=verifiedDomains' `
      --query 'value[0].verifiedDomains[?isDefault].name | [0]' `
      -o tsv 2>$null
    $derived = $raw -replace '\.onmicrosoft\.com$', ''
  }
  catch {
    # az rest for tenant detection is best-effort; failure is handled by the
    # Read-Host prompt below.
    Write-Verbose "Tenant name detection via az rest failed: $_"
  }

  if ($derived) {
    Write-Host "Detected SharePoint tenant name: $derived"
    azd env set AZURE_SHAREPOINT_TENANT_NAME $derived
  }
  else {
    $tenantName = Read-Host "Enter your SharePoint tenant name (e.g. 'contoso' for contoso.sharepoint.com)"
    azd env set AZURE_SHAREPOINT_TENANT_NAME $tenantName
  }
}

# ── 3. Create or reuse the App Registration ───────────────────────────────────
Write-Host "Checking for existing App Registration '$APP_DISPLAY_NAME'..."
$existingClientId = az ad app list `
  --display-name $APP_DISPLAY_NAME `
  --query '[0].appId' `
  -o tsv 2>$null

if ($existingClientId) {
  Write-Host "App Registration already exists. Client ID: $existingClientId"
  $clientId = $existingClientId
}
else {
  Write-Host "Creating App Registration '$APP_DISPLAY_NAME'..."
  $clientId = az ad app create `
    --display-name $APP_DISPLAY_NAME `
    --sign-in-audience 'AzureADMyOrg' `
    --description $APP_DESCRIPTION `
    --query 'appId' `
    -o tsv

  $appIdUri = "api://guest-sponsor-info-proxy/$clientId"
  az ad app update --id $clientId --identifier-uris $appIdUri | Out-Null
  Write-Host "App Registration created. App ID URI: $appIdUri"
}

# Ensure accessTokenAcceptedVersion is set to 2 (v2 tokens — aud = bare clientId).
$appObj = az ad app show --id $clientId --query 'api.requestedAccessTokenVersion' -o tsv 2>$null
if ($appObj -ne '2') {
  Write-Host "Setting accessTokenAcceptedVersion to 2..."
  az rest --method PATCH `
    --url "https://graph.microsoft.com/v1.0/applications(appId='$clientId')" `
    --body '{"api":{"requestedAccessTokenVersion":2}}' | Out-Null
}

azd env set AZURE_WEB_PART_CLIENT_ID $clientId
Write-Host "AZURE_WEB_PART_CLIENT_ID set to $clientId"
