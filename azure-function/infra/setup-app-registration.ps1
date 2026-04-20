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
#Requires -Version 5.1
param(
  [string]$TenantId,
  [string]$DisplayName = 'Guest Sponsor Info - SharePoint Web Part Auth'
)

$ErrorActionPreference = 'Stop'

# ── Console output encoding ───────────────────────────────────────────────────
# Switch to UTF-8 early so box-drawing characters and symbols (✓, ⚠) render
# correctly on Windows PowerShell 5.1 which defaults to an ANSI code page.
if ([Console]::OutputEncoding.CodePage -ne 65001) {
  try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8
  }
  catch { $null = $_ <# Non-interactive host; ignore — encoding failure is non-fatal. #> }
}

# ── Unicode output capability ─────────────────────────────────────────────────
# After the UTF-8 encoding block above, [Console]::OutputEncoding is UTF-8 on
# every host (ConsoleHost, VS Code Extension, ISE, …).  Verify by checking that
# U+2500 (BOX DRAWINGS LIGHT HORIZONTAL) encodes to more than one byte — a
# single byte would mean a legacy ANSI code page is still active.
$_u = $false
try { $_u = ([Console]::OutputEncoding.GetBytes([char]0x2500)).Length -gt 1 }
catch { $_u = $false }
$_chk = if ($_u) { [char]0x2713 } else { '[+]' }  # ✓
$_wrn = if ($_u) { [char]0x26A0 } else { '[!]' }  # ⚠
$_arr = if ($_u) { [char]0x2192 } else { '>' }    # →
$_sep = '  ' + $(if ($_u) { [string][char]0x2500 * 53 } else { '-' * 53 })

# Embedded directly so the script works on any machine without Write-Callout.ps1,
# whether run from a local clone, via iwr, or on a bare system with no repo files.
#
# Write-Host vs Write-Output in this script:
#   Write-Host   → Information stream (stream 6). Goes straight to the operator's
#                  console; cannot be captured by $x = <cmd> or piped downstream.
#                  Correct for: status messages, prompts, and colored callout boxes —
#                  anything that is purely for the operator's eyes.
#   Write-Output → Success stream (stream 1). Values flow into the pipeline and
#                  can be captured by $x = <cmd>. Use only when a function must
#                  hand data back to its caller. This script has no such functions,
#                  so Write-Output is not used.
# PSAvoidUsingWriteHost is suppressed in PSScriptAnalyzerSettings.psd1.
function Write-Box {
  param(
    [Parameter(Mandatory)][string]$Title,
    [Parameter(Mandatory)][ConsoleColor]$Color,
    [Parameter(ValueFromRemainingArguments)][string[]]$Lines
  )
  # Use the script-scope Unicode capability flag set at startup.
  if ($_u) {
    $H = [char]0x2500; $TL = [char]0x256D; $V = [char]0x2502; $BL = [char]0x2570
  }
  else {
    $H = '-'; $TL = '+'; $V = '|'; $BL = '+'
  }
  $dashes = 56 - $Title.Length
  if ($dashes -lt 4) { $dashes = 4 }
  Write-Host ''
  Write-Host "  $TL$H " -ForegroundColor $Color -NoNewline
  Write-Host $Title -ForegroundColor $Color -NoNewline
  Write-Host " $($H * $dashes)" -ForegroundColor $Color
  Write-Host "  $V" -ForegroundColor $Color
  foreach ($line in $Lines) {
    if ([string]::IsNullOrEmpty($line)) {
      Write-Host "  $V" -ForegroundColor $Color
    }
    else {
      Write-Host "  $V" -ForegroundColor $Color -NoNewline
      Write-Host "  $line"
    }
  }
  Write-Host "  $V" -ForegroundColor $Color
  Write-Host "  $BL$($H * 59)" -ForegroundColor $Color
  Write-Host ''
}
function Write-Hint {
  param([Parameter(ValueFromRemainingArguments)][string[]]$Lines)
  Write-Box -Title 'HINT' -Color Cyan @Lines
}
function Write-NextStep {
  param([Parameter(ValueFromRemainingArguments)][string[]]$Lines)
  Write-Box -Title 'NEXT STEPS' -Color Green @Lines
}
function Write-Important {
  param([Parameter(ValueFromRemainingArguments)][string[]]$Lines)
  Write-Box -Title 'IMPORTANT' -Color Yellow @Lines
}

# ── Module prerequisite helper ────────────────────────────────────────────────
# Checks whether a required PowerShell module is installed and, if not, offers
# to install it from the PowerShell Gallery. Handles:
#   - Windows PowerShell 5.1 and PowerShell 7+ on Windows, Linux, and macOS
#   - Scope selection (CurrentUser / AllUsers) on Windows
#   - OneDrive Known Folder Move (KFM) warning when CurrentUser is chosen
#   - NuGet package provider bootstrap (required by Install-Module on PS 5.1)
#   - Prefers Install-PSResource (PSResourceGet) on PS 7+ when available
function Install-RequiredModule {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Name
  )

  if (Get-Module -ListAvailable -Name $Name) {
    Import-Module -Name $Name
    return
  }

  Write-Host ''
  Write-Host "  Module '$Name' is not installed." -ForegroundColor Yellow
  $answer = (Read-Host "  Install '$Name' from the PowerShell Gallery? [Y/n]").Trim()
  if ($answer -ne '' -and $answer -notmatch '^[Yy]') {
    Write-Host "  Aborted. '$Name' is required — cannot continue." -ForegroundColor Red
    exit 1
  }

  # ── Choose installation scope ─────────────────────────────────────────────
  # On Windows (PS 5.1, which only runs on Windows, and PS 7+ when $IsWindows)
  # prompt the user to choose a scope. On Linux/macOS only CurrentUser applies.
  $scope = 'CurrentUser'
  # PS 5.1 has no $IsWindows; major version < 6 implies Windows-only runtime.
  $onWindows = ($PSVersionTable.PSVersion.Major -lt 6) -or $IsWindows

  if ($onWindows) {
    $isAdmin = ([Security.Principal.WindowsPrincipal] `
        [Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    Write-Host ''
    Write-Host '  Installation scope:' -ForegroundColor Cyan
    Write-Host '    [1] CurrentUser  – installed in your profile folder (no admin required)'
    Write-Host '    [2] AllUsers     – installed system-wide               (requires admin)'
    Write-Host ''

    if (-not $isAdmin) {
      Write-Host '  Note: You are not running as administrator.' -ForegroundColor DarkGray
      Write-Host '  AllUsers scope requires elevation. Defaulting to CurrentUser.' -ForegroundColor DarkGray
      $scopeChoice = '1'
    }
    else {
      do {
        $scopeChoice = (Read-Host '  Scope [1/2, default: 1]').Trim()
        if ($scopeChoice -eq '') { $scopeChoice = '1' }
      } while ($scopeChoice -notin @('1', '2'))
    }

    $scope = if ($scopeChoice -eq '2') { 'AllUsers' } else { 'CurrentUser' }

    # ── OneDrive Known Folder Move (KFM) warning ──────────────────────────
    # When OneDrive "Folder Backup" (Known Folder Move) redirects the Documents
    # folder to OneDrive, the CurrentUser PowerShell module path lives inside
    # that synced folder. Installed module DLLs are then uploaded to OneDrive,
    # which can cause DLL-lock conflicts during sync and slow first-use on every
    # machine sharing the same OneDrive account.
    if ($scope -eq 'CurrentUser') {
      $docsPath = [Environment]::GetFolderPath(
        [Environment+SpecialFolder]::MyDocuments)
      # Check all known OneDrive environment variables (personal, commercial, generic).
      $oneDrivePaths = @($env:OneDriveConsumer, $env:OneDriveCommercial, $env:OneDrive) |
      Where-Object { $_ }
      $kfmActive = $false
      foreach ($odPath in $oneDrivePaths) {
        if ($docsPath.StartsWith($odPath, [StringComparison]::OrdinalIgnoreCase)) {
          $kfmActive = $true
          break
        }
      }

      if ($kfmActive) {
        Write-Host ''
        Write-Host "  $_wrn  OneDrive Folder Backup is active (Known Folder Move / KFM)." `
          -ForegroundColor Yellow
        Write-Host '     Your Documents folder is currently synced to OneDrive:' -ForegroundColor Yellow
        Write-Host "     $docsPath" -ForegroundColor DarkCyan
        Write-Host ''
        Write-Host '  PowerShell modules installed in CurrentUser scope are stored inside'
        Write-Host '  your Documents folder and will therefore be synced to OneDrive.'
        Write-Host '  This may cause:'
        Write-Host '    - DLL files locked by the OneDrive sync client during module load'
        Write-Host '    - Sync conflicts when the same account is active on another computer'
        Write-Host '    - Slow availability after installation until OneDrive sync completes'
        Write-Host ''
        Write-Host '  Recommendation: choose AllUsers scope (option 2) instead, or pause'
        Write-Host '  OneDrive sync while installing the modules.'
        Write-Host ''
        $cont = (Read-Host '  Continue with CurrentUser scope anyway? [y/N]').Trim()
        if ($cont -notmatch '^[Yy]') {
          Write-Host '  Aborted.' -ForegroundColor Red
          exit 1
        }
      }
    }

    # ── NuGet package provider (required by Install-Module on PS 5.1) ──────
    # Windows PowerShell 5.1 ships without the NuGet provider; Install-Module
    # silently fails unless the provider (>= 2.8.5.201) is present.
    if ($PSVersionTable.PSVersion.Major -lt 6) {
      $nuget = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
      if (-not $nuget -or $nuget.Version -lt [Version]'2.8.5.201') {
        Write-Host ''
        Write-Host '  Installing NuGet package provider (required on Windows PowerShell…)' `
          -ForegroundColor Cyan
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 `
          -Scope $scope -Force | Out-Null
        Write-Host "  $_chk NuGet provider installed." -ForegroundColor Green
      }
    }
  }

  # Trust PSGallery to avoid interactive confirmation prompts.
  $gallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
  if ($gallery -and $gallery.InstallationPolicy -ne 'Trusted') {
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
  }

  Write-Host ''
  Write-Host "  Installing '$Name' (scope: $scope) …" -ForegroundColor Cyan

  # PowerShell 7+ with PSResourceGet: use Install-PSResource when available.
  # PSResourceGet is the modern package manager that ships with PS 7.4+ and
  # handles dependency resolution, side-by-side versions, and parallel downloads.
  if ($PSVersionTable.PSVersion.Major -ge 7 -and
    (Get-Command Install-PSResource -ErrorAction SilentlyContinue)) {
    Install-PSResource -Name $Name -Scope $scope -TrustRepository -Quiet `
      -ErrorAction Stop
  }
  else {
    # PowerShell 5.1 and PS 7 without PSResourceGet: use Install-Module.
    Install-Module -Name $Name -Scope $scope -Force -AllowClobber -ErrorAction Stop
  }

  Import-Module -Name $Name
  Write-Host "  $_chk '$Name' installed and imported." -ForegroundColor Green
  Write-Host ''
}

# ── Module bootstrap ──────────────────────────────────────────────────────────
Install-RequiredModule -Name 'Microsoft.Graph.Authentication'
Install-RequiredModule -Name 'Microsoft.Graph.Applications'

# ── Interactive parameter prompts ─────────────────────────────────────────────
# Each prompt shows a title, a short description, and where to find the value,
# then re-prompts until a valid GUID is entered.
$_guidPattern = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'

if (-not $TenantId) {
  Write-Host ''
  Write-Host '  Required: Entra Tenant ID' -ForegroundColor Cyan
  Write-Host $_sep -ForegroundColor DarkGray
  Write-Host '  Your Microsoft Entra tenant ID (a GUID).'
  Write-Host '  Where to find it:'
  Write-Host "    Microsoft Entra admin center $_arr Overview $_arr Tenant ID"
  Write-Host '    https://entra.microsoft.com' -ForegroundColor DarkCyan
  Write-Host ''
  do {
    $TenantId = (Read-Host '  Tenant ID').Trim()
    if (-not $TenantId) {
      Write-Host "  $_wrn Value is required." -ForegroundColor Yellow
    }
    elseif ($TenantId -notmatch $_guidPattern) {
      Write-Host "  $_wrn Expected a GUID: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -ForegroundColor Yellow
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
# The null-coalescing operator ?? is PS7+ only; use if/else for PS5.1 compat.
$currentUris = if ($null -eq $app.IdentifierUris) { @() } else { $app.IdentifierUris }
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
  Write-Host "  $_chk Logo uploaded." -ForegroundColor Green
}
else {
  Write-Host "  $_wrn Logo file not found at $logoPath — skipping." `
    -ForegroundColor Yellow
}

Write-Important -Lines @(
  'Copy this Client ID and use it as the ''functionClientId'''
  'parameter when deploying the ARM template. In the SPFx web part,'
  'paste it into the ''Guest Sponsor API Client ID (App Registration)'''
  'field (property pane → Guest Sponsor API).'
  ''
  "  Guest Sponsor API Client ID (App Registration): $clientId"
)
