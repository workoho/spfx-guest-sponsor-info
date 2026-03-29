<#
.SYNOPSIS
    Creates or updates the Entra App Registration needed for the Azure Function
    proxy (EasyAuth).

.DESCRIPTION
    Idempotent script that ensures an App Registration named
    "Guest Sponsor Info - SharePoint Web Part Auth" exists with the
    correct configuration:

      - Supported account types: single tenant (AzureADMyOrg)
      - App ID URI: api://guest-sponsor-info-proxy/<clientId>
      - Access token version: v2 (aud = bare clientId GUID)
      - Description populated for Ops team discoverability

    When the registration already exists the script verifies every setting and
    updates anything that drifted.  Re-running the script is always safe.

    The resulting Client ID must be provided as a parameter to the
    Bicep/ARM deployment.

.PARAMETER TenantId
    The Entra tenant ID (GUID).

.PARAMETER DisplayName
    Display name for the App Registration. Defaults to
    "Guest Sponsor Info - SharePoint Web Part Auth".

.EXAMPLE
    ./setup-app-registration.ps1 -TenantId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy"

.NOTES
    Copyright 2026 Workoho GmbH <https://workoho.com>
    Author: Julian Pawlowski <https://github.com/jpawlowski>
    Licensed under PolyForm Shield License 1.0.0
    <https://polyformproject.org/licenses/shield/1.0.0>
#>
param(
  [string]$TenantId,
  [string]$DisplayName = 'Guest Sponsor Info - SharePoint Web Part Auth'
)

$ErrorActionPreference = 'Stop'

# Dot-source callout box helpers when running from a local clone.
# When executed via iwr (remote one-liner), $PSScriptRoot is empty and the
# file won't exist — fall back to plain Write-Host stubs so the script still
# runs without visual callout boxes.
$calloutFile = Join-Path $PSScriptRoot 'Write-Callout.ps1'
if ($PSScriptRoot -and (Test-Path $calloutFile)) {
  . $calloutFile
}
else {
  # Minimal stubs — print lines without the fancy box frame.
  function Write-Hint { param([Parameter(ValueFromRemainingArguments)][string[]]$L) Write-Host ''; foreach ($l in $L) { if ($l) { Write-Host "  $l" } }; Write-Host '' }
  function Write-NextSteps { param([Parameter(ValueFromRemainingArguments)][string[]]$L) Write-Host ''; foreach ($l in $L) { if ($l) { Write-Host "  $l" } }; Write-Host '' }
  function Write-Important { param([Parameter(ValueFromRemainingArguments)][string[]]$L) Write-Host ''; foreach ($l in $L) { if ($l) { Write-Host "  $l" -ForegroundColor Yellow } }; Write-Host '' }
}

# ── Module bootstrap ─────────────────────────────────────────────────────────
foreach ($module in @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Applications'
  )) {
  if (-not (Get-Module -ListAvailable -Name $module)) {
    Write-Host "Installing $module module..." -ForegroundColor Cyan
    Install-Module $module -Scope CurrentUser -Force
  }
  Import-Module $module
}

# ── Interactive parameter prompts ─────────────────────────────────────────────
# Each prompt shows a title, a short description, and where to find the value,
# then re-prompts until a valid GUID is entered.
$_guidPattern = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'

if (-not $TenantId) {
  Write-Host ''
  Write-Host '  Required: Entra Tenant ID' -ForegroundColor Cyan
  Write-Host '  ───────────────────────────────────────────────────────' -ForegroundColor DarkGray
  Write-Host '  Your Microsoft Entra tenant ID (a GUID).'
  Write-Host '  Where to find it:'
  Write-Host '    Microsoft Entra admin center → Overview → Tenant ID'
  Write-Host '    https://entra.microsoft.com' -ForegroundColor DarkCyan
  Write-Host ''
  do {
    $TenantId = (Read-Host '  Tenant ID').Trim()
    if (-not $TenantId) {
      Write-Host '  ⚠ Value is required.' -ForegroundColor Yellow
    }
    elseif ($TenantId -notmatch $_guidPattern) {
      Write-Host '  ⚠ Expected a GUID: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx' -ForegroundColor Yellow
      $TenantId = ''
    }
  } while (-not $TenantId)
  Write-Host ''
}

Connect-MgGraph -TenantId $TenantId -Scopes "Application.ReadWrite.All"

# ── Desired state ─────────────────────────────────────────────────────────────
# We pin accessTokenAcceptedVersion to 2 (modern v2 tokens) so the aud
# claim in access tokens equals the bare clientId GUID.  ALLOWED_AUDIENCE
# and the EasyAuth allowedAudiences configuration must use the same GUID.
$desiredTokenVersion = 2

# GitHub base URL for all info links.
$repoUrl = 'https://github.com/workoho/spfx-guest-sponsor-info'

# Description shown on the App Registration overview in Entra admin center.
# Helps Ops teams identify the purpose of this registration at a glance.
$appDescription = @(
  'EasyAuth identity provider for the "Guest Sponsor Info"',
  'SharePoint Online web part (SPFx). Authenticates requests from the',
  'web part to the Azure Function proxy, which calls Microsoft Graph on',
  'behalf of signed-in guest users to retrieve their Entra sponsor',
  'information. Tokens are acquired silently via pre-authorized',
  'SharePoint Online Web Client Extensibility.',
  "Source: $repoUrl"
) -join ' '

# Internal notes for the App Registration (visible on the Overview blade).
$appNotes = @(
  'Managed by: Workoho GmbH.',
  'Do not delete — the "Guest Sponsor Info" SharePoint web part depends',
  'on this for guest sponsor lookups via Microsoft Graph.',
  'The associated Azure Function uses a system-assigned Managed Identity',
  'for Graph API calls (User.Read.All, Presence.Read.All,',
  'MailboxSettings.Read, TeamMember.Read.All).',
  "Source & docs: $repoUrl"
) -join ' '

# Info URLs shown on the App Registration "Branding & properties" blade.
$desiredInfo = @{
  TermsOfServiceUrl   = "$repoUrl/blob/main/docs/terms-of-use.md"
  PrivacyStatementUrl = "$repoUrl/blob/main/docs/privacy-policy.md"
  SupportUrl          = "$repoUrl/issues"
  MarketingUrl        = $repoUrl
}

# Logo file — uploaded from the repo if present.
$logoPath = Join-Path $PSScriptRoot '../../sharepoint/images/icon-300.png'

# ── Find or create ────────────────────────────────────────────────────────────
Write-Host "Checking for existing App Registration '$DisplayName'..." `
  -ForegroundColor Cyan
$app = Get-MgApplication -Filter "displayName eq '$DisplayName'" -Top 1

if ($app) {
  $clientId = $app.AppId
  $objectId = $app.Id
  Write-Host "App Registration already exists. Client ID: $clientId" `
    -ForegroundColor Yellow
}
else {
  Write-Host "Creating App Registration '$DisplayName'..." `
    -ForegroundColor Cyan
  $app = New-MgApplication -DisplayName $DisplayName `
    -SignInAudience 'AzureADMyOrg' `
    -Description $appDescription `
    -Notes $appNotes `
    -Info $desiredInfo `
    -Web @{ HomePageUrl = $repoUrl } `
    -Api @{ RequestedAccessTokenVersion = $desiredTokenVersion }
  $clientId = $app.AppId
  $objectId = $app.Id
  Write-Host "  Created — Client ID: $clientId" -ForegroundColor Green
}

# ── Converge to desired state (idempotent) ────────────────────────────────────
$changes = @()

# 1. Description
if ($app.Description -ne $appDescription) {
  Write-Host "  Fixing Description" -ForegroundColor Yellow
  Update-MgApplication -ApplicationId $objectId `
    -Description $appDescription
  $changes += 'Description'
}

# 2. Notes
if ($app.Notes -ne $appNotes) {
  Write-Host "  Fixing Notes" -ForegroundColor Yellow
  Update-MgApplication -ApplicationId $objectId `
    -Notes $appNotes
  $changes += 'Notes'
}

# 3. Info URLs (terms, privacy, support, marketing/homepage)
$infoChanged = $false
if ($app.Info.TermsOfServiceUrl -ne $desiredInfo.TermsOfServiceUrl) { $infoChanged = $true }
if ($app.Info.PrivacyStatementUrl -ne $desiredInfo.PrivacyStatementUrl) { $infoChanged = $true }
if ($app.Info.SupportUrl -ne $desiredInfo.SupportUrl) { $infoChanged = $true }
if ($app.Info.MarketingUrl -ne $desiredInfo.MarketingUrl) { $infoChanged = $true }
if ($infoChanged) {
  Write-Host "  Fixing Info URLs" -ForegroundColor Yellow
  Update-MgApplication -ApplicationId $objectId -Info $desiredInfo
  $changes += 'Info'
}

# 4. Homepage URL (Web section)
if ($app.Web.HomePageUrl -ne $repoUrl) {
  Write-Host "  Fixing HomePageUrl" -ForegroundColor Yellow
  Update-MgApplication -ApplicationId $objectId `
    -Web @{ HomePageUrl = $repoUrl }
  $changes += 'HomePageUrl'
}

# 5. SignInAudience
if ($app.SignInAudience -ne 'AzureADMyOrg') {
  throw ("Existing App Registration '$DisplayName' is not single-tenant " +
    "(SignInAudience=$($app.SignInAudience)). " +
    "Configure it to AzureADMyOrg before using this solution.")
}

# 6. Identifier URI
$expectedUri = "api://guest-sponsor-info-proxy/$clientId"
$currentUris = $app.IdentifierUris ?? @()
if ($currentUris -notcontains $expectedUri) {
  Write-Host "  Fixing IdentifierUris: adding $expectedUri" `
    -ForegroundColor Yellow
  Update-MgApplication -ApplicationId $objectId `
    -IdentifierUris @($expectedUri)
  $changes += 'IdentifierUris'
}

# 7. Access-token version
$currentVersion = $app.Api.RequestedAccessTokenVersion
if ($currentVersion -ne $desiredTokenVersion) {
  Write-Host ("  Fixing RequestedAccessTokenVersion: " +
    "$currentVersion -> $desiredTokenVersion") -ForegroundColor Yellow
  Update-MgApplication -ApplicationId $objectId `
    -Api @{ RequestedAccessTokenVersion = $desiredTokenVersion }
  $changes += 'RequestedAccessTokenVersion'
}

# ── Summary ───────────────────────────────────────────────────────────────────
if ($changes.Count -eq 0) {
  Write-Host "  All settings are correct — nothing to update." `
    -ForegroundColor Green
}
else {
  Write-Host ("  Updated: " + ($changes -join ', ')) `
    -ForegroundColor Green
}

# ── Logo upload ───────────────────────────────────────────────────────────────
# The Graph SDK does not support logo upload via Update-MgApplication.
# Use Invoke-MgGraphRequest to PUT the image bytes directly.
if (Test-Path $logoPath) {
  Write-Host "  Uploading logo from $logoPath ..." -ForegroundColor Cyan
  $logoBytes = [System.IO.File]::ReadAllBytes((Resolve-Path $logoPath))
  Invoke-MgGraphRequest -Method PUT `
    -Uri "https://graph.microsoft.com/v1.0/applications/$objectId/logo" `
    -ContentType 'image/png' `
    -Body $logoBytes `
    -ErrorAction Stop
  Write-Host "  ✓ Logo uploaded." -ForegroundColor Green
}
else {
  Write-Host "  ⚠ Logo file not found at $logoPath — skipping." `
    -ForegroundColor Yellow
}

Write-Important `
  'Copy this Client ID and use it as the ''functionClientId''' `
  'parameter when deploying the ARM template. In the SPFx web part,' `
  'paste it into the ''Guest Sponsor API Client ID (App Registration)''' `
  'field (property pane → Guest Sponsor API).' `
  '' `
  "  Guest Sponsor API Client ID (App Registration): $clientId"
