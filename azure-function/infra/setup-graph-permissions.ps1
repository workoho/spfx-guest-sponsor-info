<#
.SYNOPSIS
    Grants the Function App's Managed Identity the required Microsoft Graph application roles
    and configures the App Registration so the SharePoint web part can silently acquire tokens.

.DESCRIPTION
    After deploying the Azure Function, run this script to:

      1. Assign Microsoft Graph application permissions to the Managed Identity:
           - User.Read.All          (read any user's profile, sponsors, and accountEnabled status)
           - Presence.Read.All      (read sponsor presence status; optional, requires Teams)
           - MailboxSettings.Read   (filter shared/room/equipment mailboxes via userPurpose; optional, function fails open without it)

      2. Expose a 'user_impersonation' API scope on the EasyAuth App Registration and
         pre-authorize 'SharePoint Online Web Client Extensibility' to call it.
         This allows the SPFx web part to acquire tokens silently without prompting the
         user for consent or redirecting the page.

    The Managed Identity object ID is shown in the Azure Portal (Function App → Identity)
    and is also emitted as an output of the Bicep/ARM deployment.

.PARAMETER ManagedIdentityObjectId
    The object ID (not the client ID) of the Function App's system-assigned Managed Identity.

.PARAMETER TenantId
    The Entra tenant ID (GUID).

.PARAMETER FunctionAppClientId
    The client ID (application ID) of the EasyAuth App Registration created in the pre-step.
    Required to expose the API scope and pre-authorize the SharePoint client.

.EXAMPLE
    ./setup-graph-permissions.ps1 `
      -ManagedIdentityObjectId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
      -TenantId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
      -FunctionAppClientId "zzzzzzzz-zzzz-zzzz-zzzz-zzzzzzzzzzzz"

.NOTES
    Copyright 2026 Workoho GmbH <https://workoho.com>
    Author: Julian Pawlowski <https://github.com/jpawlowski>
    Licensed under PolyForm Shield License 1.0.0
    <https://polyformproject.org/licenses/shield/1.0.0>
#>
#Requires -Version 5.1
param(
  [string]$ManagedIdentityObjectId,
  [string]$TenantId,
  [string]$FunctionAppClientId
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
# Hoist to script scope so all Write-Host calls share the same capability check.
# ConsoleHost (Windows Terminal, pwsh.exe, etc.) can render Unicode box-drawing
# chars and symbols; VS Code's PowerShell Extension host cannot.
$_u = $false
try { $_u = ($Host.Name -eq 'ConsoleHost') -and ([Console]::OutputEncoding.GetBytes([char]0x2500)).Length -gt 1 }
catch { $_u = $false }
$_chk = if ($_u) { [char]0x2713 } else { '[+]' }  # ✓
$_wrn = if ($_u) { [char]0x26A0 } else { '[!]' }  # ⚠
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

# ── Module prerequisite helper ────────────────────────────────────────────
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

  # ── Choose installation scope ──────────────────────────────────────────────────
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
    Write-Host "    [1] CurrentUser  – installed in your profile folder (no admin required)"
    Write-Host "    [2] AllUsers     – installed system-wide               (requires admin)"
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

    # ── OneDrive Known Folder Move (KFM) warning ──────────────────────────────
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

    # ── NuGet package provider (required by Install-Module on PS 5.1) ──────────
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

# ── Module bootstrap ───────────────────────────────────────────────────
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
  Write-Host '    Microsoft Entra admin center → Overview → Tenant ID'
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

if (-not $ManagedIdentityObjectId) {
  Write-Host '  Required: Function App Managed Identity — Object ID' -ForegroundColor Cyan
  Write-Host $_sep -ForegroundColor DarkGray
  Write-Host '  The Object ID of the system-assigned Managed Identity of the'
  Write-Host '  Azure Function App (a GUID). This is NOT the App Client ID.'
  Write-Host '  Where to find it:'
  Write-Host '    Azure Portal → your Function App → Settings → Identity → Object (principal) ID'
  Write-Host '    or: the deployment outputs → managedIdentityObjectId'
  Write-Host '    https://portal.azure.com' -ForegroundColor DarkCyan
  Write-Host ''
  do {
    $ManagedIdentityObjectId = (Read-Host '  Managed Identity Object ID').Trim()
    if (-not $ManagedIdentityObjectId) {
      Write-Host "  $_wrn Value is required." -ForegroundColor Yellow
    }
    elseif ($ManagedIdentityObjectId -notmatch $_guidPattern) {
      Write-Host "  $_wrn Expected a GUID: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -ForegroundColor Yellow
      $ManagedIdentityObjectId = ''
    }
  } while (-not $ManagedIdentityObjectId)
  Write-Host ''
}

if (-not $FunctionAppClientId) {
  Write-Host '  Required: App Registration Client ID' -ForegroundColor Cyan
  Write-Host $_sep -ForegroundColor DarkGray
  Write-Host '  The Client ID (Application ID) of the App Registration created'
  Write-Host '  in the previous step (setup-app-registration.ps1). It was'
  Write-Host '  printed at the end of that script.'
  Write-Host '  Where to find it:'
  Write-Host '    Entra admin center → App registrations →'
  Write-Host '    "Guest Sponsor Info - SharePoint Web Part Auth" → Application (client) ID'
  Write-Host '    https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade' -ForegroundColor DarkCyan
  Write-Host ''
  do {
    $FunctionAppClientId = (Read-Host '  App Registration Client ID').Trim()
    if (-not $FunctionAppClientId) {
      Write-Host "  $_wrn Value is required." -ForegroundColor Yellow
    }
    elseif ($FunctionAppClientId -notmatch $_guidPattern) {
      Write-Host "  $_wrn Expected a GUID: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -ForegroundColor Yellow
      $FunctionAppClientId = ''
    }
  } while (-not $FunctionAppClientId)
  Write-Host ''
}

Connect-MgGraph -TenantId $TenantId -Scopes "AppRoleAssignment.ReadWrite.All", "Application.ReadWrite.All"

Write-Host "Resolving Microsoft Graph service principal..." -ForegroundColor Cyan
$graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
if (-not $graphSp) {
  throw "Could not find the Microsoft Graph service principal in tenant '$TenantId'."
}

# Resolve role IDs dynamically from the Graph service principal's app roles.
# This avoids hardcoded GUIDs and correctly detects unavailable permissions.
$requiredRoles = @(
  @{ Name = 'User.Read.All'; Optional = $false }
  @{ Name = 'Presence.Read.All'; Optional = $true }        # requires Microsoft Teams; function degrades gracefully without it
  @{ Name = 'MailboxSettings.Read'; Optional = $true }    # optional; filters shared/room/equipment mailboxes via userPurpose; without it the filter is simply skipped
  @{ Name = 'TeamMember.Read.All'; Optional = $true }     # optional; checks if the guest has joined any Team (Teams provisioning signal); fall back to presence without it
)

$assignedRoles = @()
$skippedRoles = @()

foreach ($role in $requiredRoles) {
  Write-Host "Assigning $($role.Name) ..." -ForegroundColor Cyan

  $appRole = $graphSp.AppRoles | Where-Object { $_.Value -eq $role.Name -and $_.AllowedMemberTypes -contains 'Application' }
  if (-not $appRole) {
    if ($role.Optional) {
      Write-Host "  $_wrn $($role.Name) is not available as an Application permission in this tenant (Microsoft Teams may not be licensed). Skipping — sponsors will be shown without presence status." -ForegroundColor Yellow
      $skippedRoles += $role.Name
      continue
    }
    else {
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
    Write-Host "  $_chk $($role.Name) assigned." -ForegroundColor Green
    $assignedRoles += $role.Name
  }
  catch {
    if ($_.Exception.Message -like "*Permission being assigned already exists*") {
      Write-Host "  $_chk $($role.Name) already assigned — skipping." -ForegroundColor Yellow
      $assignedRoles += $role.Name
    }
    else {
      throw
    }
  }
}

Write-Host "`nConfiguring App Registration for silent token acquisition by the SharePoint web part..." -ForegroundColor Cyan

# The SharePoint Online Web Client Extensibility app is the MSAL client that SPFx uses
# internally to acquire tokens on behalf of the signed-in user. Pre-authorizing it on the
# EasyAuth App Registration allows silent token acquisition without user consent prompts
# or full-page redirects.
#
# The actual app IDs vary across SharePoint Online environments. We resolve them dynamically
# from the tenant rather than hardcoding, then fall back to the two known canonical IDs.
Write-Host "  Resolving SharePoint Online Web Client Extensibility service principal(s)..." -ForegroundColor Cyan
$spWebClientSps = Get-MgServicePrincipal -Filter "displayName eq 'SharePoint Online Web Client Extensibility'" -All -ErrorAction SilentlyContinue
if ($spWebClientSps) {
  $spWebClientAppIds = @($spWebClientSps | Select-Object -ExpandProperty AppId)
  Write-Host "  Found $($spWebClientAppIds.Count) SP(s): $($spWebClientAppIds -join ', ')" -ForegroundColor Cyan
}
else {
  # Fall back to the two well-known first-party Microsoft app IDs used across SharePoint Online environments.
  $spWebClientAppIds = @('57fb890c-0dab-4253-a5e0-7188c88b2bb4', '08e18876-6177-487e-b8b5-cf950c1e598c')
  Write-Host "  $_wrn Could not resolve SP by display name — falling back to known app IDs: $($spWebClientAppIds -join ', ')" -ForegroundColor Yellow
}

$app = Get-MgApplication -Filter "appId eq '$FunctionAppClientId'" -ErrorAction Stop
if (-not $app) {
  throw "Could not find App Registration with client ID '$FunctionAppClientId'. Verify the -FunctionAppClientId parameter."
}

if ($app.SignInAudience -ne 'AzureADMyOrg') {
  throw "App Registration '$FunctionAppClientId' is not single-tenant (SignInAudience=$($app.SignInAudience)). Set it to AzureADMyOrg before continuing."
}

# Ensure the identifier URI is set — required for the api:// audience used by EasyAuth.
$expectedUri = "api://guest-sponsor-info-proxy/$FunctionAppClientId"
if ($app.IdentifierUris -notcontains $expectedUri) {
  Write-Host "  Setting identifier URI to $expectedUri ..." -ForegroundColor Cyan
  Update-MgApplication -ApplicationId $app.Id -IdentifierUris @($expectedUri) -ErrorAction Stop
  Write-Host "  $_chk Identifier URI set." -ForegroundColor Green
}
else {
  Write-Host "  $_chk Identifier URI already set." -ForegroundColor Yellow
}

# Expose a 'user_impersonation' OAuth2 scope if not already present.
$existingScope = $app.Api.Oauth2PermissionScopes | Where-Object { $_.Value -eq 'user_impersonation' }
if (-not $existingScope) {
  Write-Host "  Adding 'user_impersonation' scope ..." -ForegroundColor Cyan
  $scopeId = [System.Guid]::NewGuid().ToString()
  $newScope = @{
    Id                      = $scopeId
    Value                   = 'user_impersonation'
    Type                    = 'User'
    AdminConsentDisplayName = 'Access Guest Sponsor Info web part proxy as the signed-in user'
    AdminConsentDescription = 'Allows the SharePoint web part to call the Azure Function proxy on behalf of the signed-in user.'
    UserConsentDisplayName  = 'Access Guest Sponsor Info web part proxy'
    UserConsentDescription  = 'Allows the app to call the Azure Function proxy on your behalf.'
    IsEnabled               = $true
  }
  $updatedScopes = @($newScope)
  Update-MgApplication -ApplicationId $app.Id -Api @{ Oauth2PermissionScopes = $updatedScopes } -ErrorAction Stop
  # Re-fetch to get the assigned scope ID (may differ from what we sent).
  $app = Get-MgApplication -Filter "appId eq '$FunctionAppClientId'" -ErrorAction Stop
  $existingScope = $app.Api.Oauth2PermissionScopes | Where-Object { $_.Value -eq 'user_impersonation' }
  Write-Host "  $_chk 'user_impersonation' scope added (id: $($existingScope.Id))." -ForegroundColor Green
}
else {
  Write-Host "  $_chk 'user_impersonation' scope already exists (id: $($existingScope.Id))." -ForegroundColor Yellow
}

# Pre-authorize the SharePoint Online Web Client Extensibility app(s) to call the scope.
# This is what makes token acquisition silent — no per-user consent prompt, no page redirect.
foreach ($spAppId in $spWebClientAppIds) {
  $alreadyPreAuthorized = $app.Api.PreAuthorizedApplications | Where-Object {
    $_.AppId -eq $spAppId -and
    $_.DelegatedPermissionIds -contains $existingScope.Id
  }
  if (-not $alreadyPreAuthorized) {
    Write-Host "  Pre-authorizing $spAppId ..." -ForegroundColor Cyan
    # Re-fetch the current state before each update to avoid overwriting parallel changes.
    $app = Get-MgApplication -Filter "appId eq '$FunctionAppClientId'" -ErrorAction Stop
    $otherPreAuthorized = $app.Api.PreAuthorizedApplications | Where-Object { $_.AppId -ne $spAppId }
    $newPreAuth = @{
      AppId                  = $spAppId
      DelegatedPermissionIds = @($existingScope.Id)
    }
    $updatedPreAuthorized = @($otherPreAuthorized) + @($newPreAuth)
    try {
      Update-MgApplication -ApplicationId $app.Id -Api @{ PreAuthorizedApplications = $updatedPreAuthorized } -ErrorAction Stop
      Write-Host "  $_chk $spAppId pre-authorized." -ForegroundColor Green
    }
    catch {
      if ($_.Exception.Message -like "*cannot be found*") {
        Write-Host "  $_wrn $spAppId not found in Microsoft's app registry — skipping." -ForegroundColor Yellow
      }
      else {
        throw
      }
    }
  }
  else {
    Write-Host "  $_chk $spAppId already pre-authorized." -ForegroundColor Yellow
  }
}

# Ensure appRoleAssignmentRequired is false on the Service Principal (Enterprise App).
# Normally created on first user sign-in, but since we run this script before any user
# has consented, we create it explicitly here.
$sp = Get-MgServicePrincipal -Filter "appId eq '$FunctionAppClientId'" -ErrorAction SilentlyContinue
if (-not $sp) {
  Write-Host "  Service Principal not found — creating it now (no user has signed in yet)..." -ForegroundColor Cyan
  $sp = New-MgServicePrincipal -AppId $FunctionAppClientId -ErrorAction Stop
  Write-Host "  $_chk Service Principal created (Object ID: $($sp.Id))." -ForegroundColor Green
}
else {
  Write-Host "  $_chk Service Principal already exists (Object ID: $($sp.Id))." -ForegroundColor Yellow
}

# appRoleAssignmentRequired=false: all users (including guests) can acquire tokens without
# individual assignment — even with pre-authorization in place.
if ($sp.AppRoleAssignmentRequired) {
  Write-Host "  Disabling appRoleAssignmentRequired on the Enterprise App (was: true) ..." -ForegroundColor Cyan
  Update-MgServicePrincipal -ServicePrincipalId $sp.Id -AppRoleAssignmentRequired:$false -ErrorAction Stop
  Write-Host "  $_chk appRoleAssignmentRequired set to false." -ForegroundColor Green
}
else {
  Write-Host "  $_chk appRoleAssignmentRequired is already false — no user assignment needed." -ForegroundColor Yellow
}

# Hide from My Apps portal (tags: HideApp). This is a backend auth proxy — it should not
# appear as a launchable app in users' My Apps page.
$hasHideApp = $sp.Tags -contains 'HideApp'
if (-not $hasHideApp) {
  Write-Host "  Hiding Enterprise App from My Apps portal (visible to users: No) ..." -ForegroundColor Cyan
  $updatedTags = @($sp.Tags) + @('HideApp')
  Update-MgServicePrincipal -ServicePrincipalId $sp.Id -Tags $updatedTags -ErrorAction Stop
  Write-Host "  $_chk Enterprise App hidden from My Apps portal." -ForegroundColor Green
}
else {
  Write-Host "  $_chk Enterprise App is already hidden from My Apps portal." -ForegroundColor Yellow
}

# Service Principal description — mirrors the App Registration description so Ops
# teams see the purpose in both the "App registrations" and "Enterprise applications" blades.
$spDescription = @(
  'EasyAuth identity provider for the "Guest Sponsor Info"',
  'SharePoint Online web part (SPFx). Authenticates requests from the',
  'web part to the Azure Function proxy, which calls Microsoft Graph on',
  'behalf of signed-in guest users to retrieve their Entra sponsor',
  'information. Tokens are acquired silently via pre-authorized',
  'SharePoint Online Web Client Extensibility.',
  'Source: https://github.com/workoho/spfx-guest-sponsor-info'
) -join ' '

# Notes field — visible under Enterprise App → Properties. Ideal for Ops runbook hints.
$spNotes = @(
  'Managed by: Workoho GmbH.',
  'Do not delete — the "Guest Sponsor Info" SharePoint web part depends',
  'on this for guest sponsor lookups via Microsoft Graph.',
  'This app should remain hidden from My Apps (HideApp tag).',
  'The associated Azure Function uses a system-assigned Managed',
  'Identity for Graph API calls (User.Read.All, Presence.Read.All,',
  'MailboxSettings.Read, TeamMember.Read.All).',
  'Source & docs: https://github.com/workoho/spfx-guest-sponsor-info'
) -join ' '

if ($sp.Description -ne $spDescription) {
  Write-Host "  Setting Enterprise App description ..." -ForegroundColor Cyan
  Update-MgServicePrincipal -ServicePrincipalId $sp.Id `
    -Description $spDescription -ErrorAction Stop
  Write-Host "  $_chk Description set." -ForegroundColor Green
}
else {
  Write-Host "  $_chk Enterprise App description already set." -ForegroundColor Yellow
}

if ($sp.Notes -ne $spNotes) {
  Write-Host "  Setting Enterprise App notes ..." -ForegroundColor Cyan
  Update-MgServicePrincipal -ServicePrincipalId $sp.Id `
    -Notes $spNotes -ErrorAction Stop
  Write-Host "  $_chk Notes set." -ForegroundColor Green
}
else {
  Write-Host "  $_chk Enterprise App notes already set." -ForegroundColor Yellow
}
# Service Management Reference — shown under Enterprise App → Properties.
# Points to the GitHub Issues tracker so Ops teams know where to file tickets.
$desiredSmRef = 'https://github.com/workoho/spfx-guest-sponsor-info/issues'
if ($sp.ServiceManagementReference -ne $desiredSmRef) {
  Write-Host "  Setting Service Management Reference ..." -ForegroundColor Cyan
  Update-MgServicePrincipal -ServicePrincipalId $sp.Id `
    -ServiceManagementReference $desiredSmRef -ErrorAction Stop
  Write-Host "  $_chk Service Management Reference set." -ForegroundColor Green
}
else {
  Write-Host "  $_chk Service Management Reference already set." -ForegroundColor Yellow
}

# Homepage URL — visible under Enterprise App → Properties.
$desiredHomepage = 'https://github.com/workoho/spfx-guest-sponsor-info'
if ($sp.Homepage -ne $desiredHomepage) {
  Write-Host "  Setting Enterprise App homepage URL ..." -ForegroundColor Cyan
  Update-MgServicePrincipal -ServicePrincipalId $sp.Id `
    -Homepage $desiredHomepage -ErrorAction Stop
  Write-Host "  $_chk Homepage URL set." -ForegroundColor Green
}
else {
  Write-Host "  $_chk Enterprise App homepage URL already set." -ForegroundColor Yellow
}
# Build a summary of assigned and skipped roles for the callout box.
$summaryLines = @('The Managed Identity can now call Microsoft Graph with:')
foreach ($r in $assignedRoles) {
  $summaryLines += "  - $r"
}
if ($skippedRoles.Count -gt 0) {
  $summaryLines += ''
  $summaryLines += 'Skipped (not available in this tenant):'
  foreach ($r in $skippedRoles) {
    $summaryLines += "  - $r"
  }
}
$summaryLines += ''
$summaryLines += 'The App Registration is configured for silent token acquisition:'
$summaryLines += "  - Identifier URI: $expectedUri"
$summaryLines += "  - Scope 'user_impersonation' exposed and SharePoint pre-authorized."
$summaryLines += ''
$summaryLines += 'The SharePoint web part can now acquire tokens silently.'
$summaryLines += 'No page reloads or consent prompts.'
Write-NextStep @summaryLines
