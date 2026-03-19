#!/usr/bin/env pwsh
# Pre-provision hook for Azure Developer CLI (azd).
# Runs before Bicep deployment to:
#   1. Derive a default Function App name from the azd environment name.
#   2. Detect or prompt for the SharePoint tenant name.
#   3. Create (or reuse) the Entra App Registration required for EasyAuth,
#      and store its client ID as AZURE_FUNCTION_CLIENT_ID in the azd environment.
#
# All operations are idempotent — safe to re-run on 'azd provision' or 'azd up'.

$ErrorActionPreference = 'Stop'

$APP_DISPLAY_NAME = 'Guest Sponsor Info Proxy'
$envValues = azd env get-values

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
    } catch { }

    if ($derived) {
        Write-Host "Detected SharePoint tenant name: $derived"
        azd env set AZURE_SHAREPOINT_TENANT_NAME $derived
    } else {
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
} else {
    Write-Host "Creating App Registration '$APP_DISPLAY_NAME'..."
    $clientId = az ad app create `
        --display-name $APP_DISPLAY_NAME `
        --sign-in-audience 'AzureADMyOrg' `
        --query 'appId' `
        -o tsv

    $appIdUri = "api://guest-sponsor-info-proxy/$clientId"
    az ad app update --id $clientId --identifier-uris $appIdUri | Out-Null
    Write-Host "App Registration created. App ID URI: $appIdUri"
}

azd env set AZURE_FUNCTION_CLIENT_ID $clientId
Write-Host "AZURE_FUNCTION_CLIENT_ID set to $clientId"
