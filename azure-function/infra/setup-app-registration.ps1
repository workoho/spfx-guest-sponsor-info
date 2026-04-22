#!/usr/bin/env -S pwsh -NoLogo -NoProfile

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

.PARAMETER Confirm
    -Confirm:$false skips all interactive confirmations and runs unattended.
    When all parameters are supplied on the command line the script shows a
    summary of planned changes and asks once before connecting; pass
    -Confirm:$false together with all required parameters to bypass that
    prompt completely.

.PARAMETER WhatIf
    Shows what the script would do without making any changes. All write
    operations (POST / PATCH / PUT) are skipped; read operations still run
    so resolved values are shown.

.EXAMPLE
    ./setup-app-registration.ps1 -TenantId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy"

.EXAMPLE
    # Fully unattended — no prompts, no confirmation summary
    ./setup-app-registration.ps1 `
      -TenantId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
      -Confirm:$false

.NOTES
    Copyright 2026 Workoho GmbH <https://workoho.com>
    Author: Julian Pawlowski <https://github.com/jpawlowski>
    Licensed under PolyForm Shield License 1.0.0
    <https://polyformproject.org/licenses/shield/1.0.0>
#>

#region Parameters
#Requires -Version 5.1
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
  [string]$TenantId,
  [string]$DisplayName = 'Guest Sponsor Info - SharePoint Web Part Auth'
)

$ErrorActionPreference = 'Stop'

# Track whether any interactive prompt was shown. When all parameters were
# pre-supplied (via the command line or the session cache) we show a
# confirmation summary so the operator can verify before the script runs.
$_promptsShown = $false
# Convenience bool used throughout for WhatIf-aware fallbacks.
$_whatIf = $WhatIfPreference -eq [System.Management.Automation.SwitchParameter]$true
#endregion

#region Terminal initialization
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

# ── OSC 8 hyperlink capability ─────────────────────────────────────────────────
# OSC 8 clickable hyperlinks are supported by Windows Terminal, VS Code,
# iTerm2, WezTerm, Kitty, Foot, GNOME Terminal, Konsole, and most modern
# Linux terminals that advertise 24-bit colour support.
# Disabled automatically when stdout is redirected (no attached console).
$_osc8 = $false
if (-not [Console]::IsOutputRedirected) {
  $_osc8 = (
    $env:WT_SESSION -or # Windows Terminal
    $env:TERM_PROGRAM -eq 'vscode' -or # VS Code integrated terminal
    $env:TERM_PROGRAM -eq 'iTerm.app' -or # iTerm2
    $env:TERM_PROGRAM -eq 'WezTerm' -or # WezTerm
    $env:TERM -eq 'xterm-kitty' -or # Kitty
    $env:TERM -eq 'foot' -or # Foot (Wayland)
    $env:COLORTERM -eq 'truecolor' -or # Most modern Linux/macOS terminals
    $env:COLORTERM -eq '24bit' -or # Alternative truecolor flag
    $env:VTE_VERSION -or # GNOME Terminal / VTE-based
    $env:KONSOLE_VERSION                         # Konsole (KDE)
  ) -as [bool]
}
#endregion

#region Output helpers
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
    $H = [string][char]0x2500; $TL = [string][char]0x256D; $V = [string][char]0x2502; $BL = [string][char]0x2570
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
function Write-Failure {
  param([Parameter(ValueFromRemainingArguments)][string[]]$Lines)
  Write-Box -Title 'ERROR' -Color Red @Lines
}
function Write-Link {
  # Print a deep link to a URL. In terminals that support OSC 8 escape
  # sequences the link text is rendered as a clickable hyperlink, prefixed
  # with a ↗ arrow (U+2197) in Cyan so the click intent is obvious even
  # to users unfamiliar with terminal hyperlinks. In all other hosts the
  # label and URL are printed on two lines so nothing is lost.
  param(
    [Parameter(Mandatory)][string]$Url,
    [string]$Text,
    [string]$Indent = '    '
  )
  if ([string]::IsNullOrEmpty($Text)) { $Text = $Url }
  # ↗ (U+2197) signals "navigate / open link"; '>' on legacy hosts.
  $linkArrow = if ($_u) { [string][char]0x2197 } else { '>' }
  if ($_osc8) {
    $esc = [char]27
    Write-Host "$Indent$linkArrow " -NoNewline -ForegroundColor Cyan
    Write-Host "$($esc)]8;;$($Url)$($esc)\$($Text)$($esc)]8;;$($esc)\" -ForegroundColor DarkCyan
  }
  else {
    Write-Host "$Indent$linkArrow $Text"
    Write-Host "$Indent  $Url" -ForegroundColor DarkCyan
  }
}
#endregion

#region Error handler
# Script-level trap: on Graph authorization errors (401/403), print role
# guidance instead of a raw HTTP exception. Other errors re-throw normally.
trap {
  $_httpCode = $null
  if ($_.Exception -and $null -ne $_.Exception.Response) {
    try { $_httpCode = [int]$_.Exception.Response.StatusCode }
    catch { $null = $_ <# cast may fail if StatusCode is not numeric — ignore #> }
  }
  $_errMsg = if ($_.Exception -and $_.Exception.Message) { $_.Exception.Message } else { [string]$_ }
  # In WhatIf mode every error from a read/validate operation is non-fatal.
  # PowerShell resumes after the failing statement with $null as its value;
  # downstream ShouldProcess calls then print the "What if:" messages instead.
  if ($_whatIf) {
    $_shortMsg = ($_errMsg -replace '\r?\n.*', '').Trim()
    if ($_shortMsg) {
      Write-Host "  [WhatIf] $_shortMsg — treating object as non-existent." -ForegroundColor Yellow
    }
    continue
  }
  if ($_httpCode -in @(401, 403) -or
    $_errMsg -match 'Authorization_RequestDenied|Forbidden|Unauthorized|insufficient.privilege') {
    Write-Failure -Lines @(
      'The request was denied — your account lacks the required permissions.'
      ''
      'Required Entra role:  Cloud Application Administrator'
      '  (or Application Administrator, or Global Administrator)'
      ''
      'If your role is eligible (PIM): activate it, then re-run.'
      'If you do not have the role yet: request it from your admin.'
    )
    Write-Link -Url 'https://entra.microsoft.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/aadRoles' `
      -Text 'PIM → My roles → Entra roles  (activate eligible role)'
    Write-Link -Url 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/RolesManagementMenuBlade/~/AllRoles' `
      -Text 'Entra admin center → Roles and administrators'
    break
  }
  # Not a permission error — let PowerShell display the raw error and exit.
}
#endregion

#region Module management
# ── Module prerequisite helper ────────────────────────────────────────────────
# Checks whether a required PowerShell module is installed and, if not, offers
# to install it from the PowerShell Gallery. Handles:
#   - Explicit user confirmation before any installation attempt
#   - Windows PowerShell 5.1 and PowerShell 7+ on Windows, Linux, and macOS
#   - Running as admin: offer AllUsers (default) or CurrentUser
#   - Not running as admin on Windows:
#       [1] Elevate temporarily via UAC to install AllUsers (recommended)
#       [2] Install CurrentUser — with OneDrive KFM warning if applicable
#       [3] Abort — install manually and re-run this script
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
  Write-Host '  This module is required — it cannot be skipped.'
  Write-Host ''
  Write-Host '  To install manually and then re-run this script:'
  Write-Host "    Install-Module -Name '$Name' -Scope CurrentUser" -ForegroundColor Cyan
  Write-Host ''
  $answer = (Read-Host "  Proceed with automatic installation? [Y/n]").Trim()
  if ($answer -ne '' -and $answer -notmatch '^[Yy]') {
    Write-Host "  Aborted. '$Name' is required — cannot continue." -ForegroundColor Red
    exit 1
  }

  # ── Choose installation scope ─────────────────────────────────────────────
  # PS 5.1 has no $IsWindows; major version < 6 implies Windows-only runtime.
  $onWindows = ($PSVersionTable.PSVersion.Major -lt 6) -or $IsWindows
  $scope = 'CurrentUser'

  if ($onWindows) {
    $isAdmin = ([Security.Principal.WindowsPrincipal] `
        [Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if ($isAdmin) {
      # Running as administrator: offer AllUsers (default) or CurrentUser.
      Write-Host ''
      Write-Host '  Installation scope:' -ForegroundColor Cyan
      Write-Host '    [1] AllUsers     – installed system-wide              (recommended)'
      Write-Host '    [2] CurrentUser  – installed in your profile folder'
      Write-Host ''
      do {
        $scopeChoice = (Read-Host '  Scope [1/2, default: 1]').Trim()
        if ($scopeChoice -eq '') { $scopeChoice = '1' }
      } while ($scopeChoice -notin @('1', '2'))
      $scope = if ($scopeChoice -eq '2') { 'CurrentUser' } else { 'AllUsers' }
    }
    else {
      # Not running as admin. Check whether CurrentUser installs would land
      # inside an OneDrive-synced folder (Known Folder Move / KFM).
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

      Write-Host ''
      Write-Host '  You are not running as administrator.' -ForegroundColor DarkGray
      if ($kfmActive) {
        Write-Host ''
        Write-Host "  $_wrn OneDrive Folder Backup (KFM) is active." -ForegroundColor Yellow
        Write-Host "       CurrentUser modules would be stored in:" -ForegroundColor Yellow
        Write-Host "       $docsPath" -ForegroundColor DarkCyan
        Write-Host '       This folder syncs to OneDrive — DLL-lock conflicts are possible.' `
          -ForegroundColor Yellow
        Write-Host ''
        Write-Host '  Choose an option:' -ForegroundColor Cyan
        Write-Host '    [1] Elevate to install system-wide   (recommended — opens a UAC prompt)'
        Write-Host '    [2] Install in profile anyway        (stored in OneDrive, not recommended)'
        Write-Host '    [3] Abort — install manually and re-run this script'
      }
      else {
        Write-Host ''
        Write-Host '  Choose an option:' -ForegroundColor Cyan
        Write-Host '    [1] Elevate to install system-wide   (recommended — opens a UAC prompt)'
        Write-Host '    [2] Install in profile folder        (CurrentUser — no admin required)'
        Write-Host '    [3] Abort — install manually and re-run this script'
      }
      Write-Host ''
      do {
        $installChoice = (Read-Host '  Option [1/2/3, default: 1]').Trim()
        if ($installChoice -eq '') { $installChoice = '1' }
      } while ($installChoice -notin @('1', '2', '3'))

      switch ($installChoice) {
        '1' {
          # Elevate: run Install-Module as a local admin in a new elevated window.
          # Start-Process -Verb RunAs triggers UAC; -Wait blocks until it finishes.
          # $psExe uses the same PowerShell major version that is currently running.
          $psExe = if ($PSVersionTable.PSVersion.Major -ge 7) { 'pwsh' } else { 'powershell' }
          # The inner command wraps Install-Module in try/catch and exits with code 1
          # on failure so the caller can detect the error after the window closes.
          $innerCmd = "try { Install-Module -Name '$Name' -Scope AllUsers " +
          "-Force -AllowClobber -ErrorAction Stop } catch { exit 1 }"
          Write-Host ''
          Write-Host "  Launching elevated installer for '$Name' (AllUsers) …" -ForegroundColor Cyan
          Write-Host '  A UAC prompt will appear — approve it to continue.' -ForegroundColor DarkGray
          try {
            $proc = Start-Process -FilePath $psExe `
              -ArgumentList "-NoProfile -NonInteractive -Command `"$innerCmd`"" `
              -Verb RunAs -Wait -PassThru -ErrorAction Stop
            if ($null -ne $proc.ExitCode -and $proc.ExitCode -ne 0) {
              throw "Elevated installer exited with code $($proc.ExitCode)."
            }
            Write-Host "  $_chk '$Name' installed system-wide." -ForegroundColor Green
          }
          catch {
            Write-Host ''
            Write-Host "  $_wrn Elevation failed or was cancelled: $($_.Exception.Message)" `
              -ForegroundColor Red
            Write-Host ''
            Write-Host '  Install manually and then re-run this script:' -ForegroundColor Yellow
            Write-Host "    Install-Module -Name '$Name' -Scope AllUsers" -ForegroundColor Cyan
            Write-Host ''
            exit 1
          }
          Import-Module -Name $Name
          Write-Host "  $_chk '$Name' imported." -ForegroundColor Green
          Write-Host ''
          return
        }
        '2' {
          $scope = 'CurrentUser'
        }
        '3' {
          Write-Host ''
          Write-Host '  Install manually and then re-run this script:' -ForegroundColor Yellow
          Write-Host "    Install-Module -Name '$Name' -Scope AllUsers" -ForegroundColor Cyan
          Write-Host "    Install-Module -Name '$Name' -Scope CurrentUser  # no admin required" `
            -ForegroundColor DarkGray
          Write-Host ''
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
  try {
    if ($PSVersionTable.PSVersion.Major -ge 7 -and
      (Get-Command Install-PSResource -ErrorAction SilentlyContinue)) {
      Install-PSResource -Name $Name -Scope $scope -TrustRepository -Quiet `
        -ErrorAction Stop
    }
    else {
      # PowerShell 5.1 and PS 7 without PSResourceGet: use Install-Module.
      Install-Module -Name $Name -Scope $scope -Force -AllowClobber -ErrorAction Stop
    }
  }
  catch {
    Write-Host ''
    Write-Host "  $_wrn Auto-installation failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ''
    Write-Host '  Install the module manually and then re-run this script:' -ForegroundColor Yellow
    Write-Host "    Install-Module -Name '$Name' -Scope CurrentUser" -ForegroundColor Cyan
    Write-Host ''
    exit 1
  }

  Import-Module -Name $Name
  Write-Host "  $_chk '$Name' installed and imported." -ForegroundColor Green
  Write-Host ''
}

# ── PowerShell version check ──────────────────────────────────────────────────
# PowerShell 7+ is recommended for stable JSON deserialization from
# Invoke-MgGraphRequest. Windows PowerShell 5.1 is supported but may behave
# differently with complex nested objects.
if ($PSVersionTable.PSVersion.Major -lt 7) {
  Write-Host ''
  Write-Host "  $_wrn Running on Windows PowerShell $($PSVersionTable.PSVersion) — PowerShell 7 is recommended." `
    -ForegroundColor Yellow
  Write-Host '       Download: https://aka.ms/install-powershell' -ForegroundColor DarkCyan
  Write-Host ''
}

# ── Module bootstrap ──────────────────────────────────────────────────────────
# Only Microsoft.Graph.Authentication is needed — all Graph API calls use
# Invoke-MgGraphRequest directly (no higher-level SDK module required).
Install-RequiredModule -Name 'Microsoft.Graph.Authentication'
#endregion

#region Parameter collection
# ── Interactive parameter prompts ─────────────────────────────────────────────
# Each prompt shows a title, a short description, and where to find the value,
# then re-prompts until a valid GUID is entered.
$_guidPattern = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'

if (-not $TenantId) {
  # Session cache: skip the prompt if this script or setup-graph-permissions.ps1
  # already stored a value in the $Global:GsiSetup_* namespace during this session.
  if ($Global:GsiSetup_TenantId) {
    $TenantId = $Global:GsiSetup_TenantId
    Write-Host "  Using cached Tenant ID from this session: $TenantId" -ForegroundColor DarkGray
  }
  else {
    # Check for an existing Graph session — offer the connected tenant as default
    # so the user can just press Enter instead of looking the GUID up.
    $_existingCtx = $null
    try { $_existingCtx = Get-MgContext -ErrorAction SilentlyContinue } catch { $null = $_ }
    $_defaultTenant = if ($_existingCtx -and $_existingCtx.TenantId) { $_existingCtx.TenantId } else { '' }
    Write-Host ''
    Write-Host '  Required: Entra Tenant ID' -ForegroundColor Cyan
    Write-Host $_sep -ForegroundColor DarkGray
    Write-Host '  Your Microsoft Entra tenant ID (a GUID).'
    if ($_defaultTenant) {
      Write-Host "  Active Graph session: $($_existingCtx.Account)" -ForegroundColor DarkGray
      Write-Host "  Tenant: $_defaultTenant" -ForegroundColor DarkGray
      Write-Host '  Press Enter to use this tenant.' -ForegroundColor DarkGray
    }
    Write-Host '  Where to find it:'
    Write-Link -Url 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/TenantOverview.ReactView' `
      -Text "Microsoft Entra admin center $_arr Overview $_arr Tenant ID"
    Write-Host ''
    do {
      $_prompt = if ($_defaultTenant) { "  Tenant ID [$_defaultTenant]" } else { '  Tenant ID' }
      $TenantId = (Read-Host $_prompt).Trim()
      if (-not $TenantId -and $_defaultTenant) {
        $TenantId = $_defaultTenant
      }
      if (-not $TenantId) {
        Write-Host "  $_wrn Value is required." -ForegroundColor Yellow
      }
      elseif ($TenantId -notmatch $_guidPattern) {
        Write-Host "  $_wrn Expected a GUID: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -ForegroundColor Yellow
        $TenantId = ''
      }
    } while (-not $TenantId)
    Write-Host ''
    $_promptsShown = $true
  }
}
# Save the resolved TenantId to the session cache — available to repeated runs
# and to setup-graph-permissions.ps1 when run in the same PowerShell session.
$Global:GsiSetup_TenantId = $TenantId

Write-Hint -Lines @(
  'Required Entra role:  Cloud Application Administrator'
  '  (or Application Administrator, or Global Administrator)'
  ''
  'If your role is eligible (PIM): activate it, then re-run.'
  'If you do not have the role yet: request it from your admin.'
)
Write-Link -Url "https://entra.microsoft.com/$TenantId/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/aadRoles" `
  -Text 'PIM → My roles → Entra roles  (activate eligible role)'
Write-Link -Url "https://entra.microsoft.com/$TenantId/#view/Microsoft_AAD_IAM/RolesManagementMenuBlade/~/AllRoles" `
  -Text 'Entra admin center → Roles and administrators'

# When no interactive prompts were shown (params from CLI or session cache),
# show a summary of what the script is about to do and ask for confirmation
# — unless the caller already passed -Confirm:$false or -WhatIf.
if (-not $_promptsShown -and
  $WhatIfPreference -ne [System.Management.Automation.SwitchParameter]$true -and
  $ConfirmPreference -ne 'None') {
  Write-Host ''
  Write-Host '  Planned operations' -ForegroundColor Cyan
  Write-Host $_sep -ForegroundColor DarkGray
  Write-Host "  Tenant ID  : $TenantId"
  Write-Host "  App Name   : $DisplayName"
  Write-Host ''
  Write-Host '  The script will create or update the App Registration and' -ForegroundColor DarkGray
  Write-Host '  upload the logo.  All operations are idempotent.' -ForegroundColor DarkGray
  Write-Host ''
  $reply = (Read-Host '  Proceed? [Y/n]').Trim()
  if ($reply -and $reply -notmatch '^[Yy]') {
    Write-Host 'Aborted.' -ForegroundColor Yellow
    exit 0
  }
  Write-Host ''
}
#endregion

#region Graph connection
# ── Connect to Microsoft Graph ────────────────────────────────────────────────────────
# Skip Connect-MgGraph when the current session already covers the required
# scopes for the right tenant — avoids unnecessary MFA / browser prompts.
$_requiredScopes = if ($_whatIf) { @('Application.Read.All') } else { @('Application.ReadWrite.All') }
$_mgCtx = $null
try { $_mgCtx = Get-MgContext -ErrorAction SilentlyContinue } catch { $null = $_ }
$_scopesOk = $false
if ($_mgCtx -and $_mgCtx.TenantId -eq $TenantId) {
  $_scopesOk = $true
  foreach ($_s in $_requiredScopes) {
    if ($_mgCtx.Scopes -notcontains $_s) { $_scopesOk = $false; break }
  }
}

if ($_whatIf) {
  if ($_scopesOk) {
    Write-Host "  [WhatIf] Reusing existing Graph session — scopes sufficient." -ForegroundColor DarkGray
  }
  else {
    # In WhatIf mode read-only scope is sufficient — write permissions are not
    # needed for a dry run.  Fall back to offline simulation if sign-in fails.
    Write-Host "  [WhatIf] Connecting with read-only scope (Application.Read.All)..." -ForegroundColor DarkGray
    try {
      Connect-MgGraph -TenantId $TenantId -Scopes 'Application.Read.All' -ErrorAction Stop
      Write-Host "  [WhatIf] Connected — current state will be reflected where readable." -ForegroundColor DarkGray
    }
    catch {
      Write-Host "  [WhatIf] Sign-in failed — simulating all operations as 'would create/update'." -ForegroundColor Yellow
    }
  }
}
else {
  if ($_scopesOk) {
    Write-Host "Reusing existing Graph session." -ForegroundColor DarkGray
  }
  else {
    Connect-MgGraph -TenantId $TenantId -Scopes 'Application.ReadWrite.All'
  }
}

# ── Verify active Entra role assignments ─────────────────────────────────────
# Opportunistic: requires Directory.Read.All in the current token.
# Silently skipped when the scope is absent or the session is a service
# principal (workload identity).  The check is informational only — the
# script will still fail later with a clear error if the role is missing.
$_checkRoles = @(
  'Cloud Application Administrator'
  'Application Administrator'
  'Global Administrator'
)
$_activeRoles = @()
$_roleCheckOk = $false
try {
  $_roleResp = Invoke-MgGraphRequest -Method GET `
    -Uri 'https://graph.microsoft.com/v1.0/me/transitiveMemberOf/microsoft.graph.directoryRole?$select=displayName' `
    -OutputType PSObject -ErrorAction Stop
  if ($_roleResp.value) {
    # Keep only the roles relevant to this script so the output stays focused.
    $_activeRoles = @($($_roleResp.value | Where-Object { $_checkRoles -contains $_.displayName } | ForEach-Object { $_.displayName }))
  }
  $_roleCheckOk = $true
}
catch { $null = $_ }

if ($_roleCheckOk) {
  if ($_activeRoles.Count -gt 0) {
    $_chk = if ($_u) { [string][char]0x2713 } else { 'OK' }
    Write-Host "  $_chk Active role(s): $($_activeRoles -join ', ')" -ForegroundColor Green
  }
  else {
    # Connected successfully but none of the required roles are active.
    Write-Host "  $_wrn No required Entra role is active for your account." -ForegroundColor Yellow
    Write-Host '  Activate your role via PIM or request it from your admin.' -ForegroundColor Yellow
  }
}
#endregion

#region Desired state
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
  'Docs: https://guest-sponsor-info.workoho.cloud'
) -join ' '

# Internal notes for the App Registration (visible on the Overview blade).
$appNotes = @(
  'Do not delete — the "Guest Sponsor Info" SharePoint web part depends',
  'on this for guest sponsor lookups via Microsoft Graph.',
  'The associated Azure Function uses a system-assigned Managed Identity',
  'for Graph API calls (User.Read.All, Presence.Read.All,',
  'MailboxSettings.Read, TeamMember.Read.All).',
  'The Managed Identity Object ID and its Graph app role assignments are',
  'configured by setup-graph-permissions.ps1.',
  'Docs: https://guest-sponsor-info.workoho.cloud'
) -join ' '

# Info URLs shown on the App Registration "Branding & properties" blade.
# Keys are camelCase to match the Graph REST API JSON property names used in
# Invoke-MgGraphRequest PATCH bodies.
$desiredInfo = @{
  termsOfServiceUrl   = "$repoUrl/blob/main/docs/terms-of-use.md"
  privacyStatementUrl = "$repoUrl/blob/main/docs/privacy-policy.md"
  supportUrl          = "$repoUrl/issues"
  marketingUrl        = 'https://guest-sponsor-info.workoho.cloud'
}

# Logo file — resolved from the local repo clone when available; otherwise
# downloaded from GitHub so the script works when run directly via iwr.
$_localLogoPath = if ($PSScriptRoot) {
  Join-Path $PSScriptRoot '../../sharepoint/images/icon-300.png'
}
else { $null }
$logoPath = if ($_localLogoPath -and (Test-Path $_localLogoPath)) {
  $_localLogoPath
}
else { $null }
$logoRawUrl = "$repoUrl" -replace 'https://github\.com', 'https://raw.githubusercontent.com'
$logoRawUrl = "$logoRawUrl/main/sharepoint/images/icon-300.png"
#endregion

#region Find or create
# ── Find or create ────────────────────────────────────────────────────────────
Write-Host "Checking for existing App Registration '$DisplayName'..." `
  -ForegroundColor Cyan
$_appListResponse = Invoke-MgGraphRequest -Method GET `
  -Uri "https://graph.microsoft.com/v1.0/applications?`$filter=displayName eq '$DisplayName'&`$top=1" `
  -ErrorAction Stop
$app = if ($_appListResponse.value.Count -gt 0) { $_appListResponse.value[0] } else { $null }

if ($app) {
  $clientId = $app.appId
  $objectId = $app.id
  Write-Host "App Registration already exists. Client ID: $clientId" `
    -ForegroundColor Yellow
}
else {
  Write-Host "Creating App Registration '$DisplayName'..." `
    -ForegroundColor Cyan
  if ($PSCmdlet.ShouldProcess($DisplayName, 'Create App Registration')) {
    $app = Invoke-MgGraphRequest -Method POST `
      -Uri 'https://graph.microsoft.com/v1.0/applications' `
      -Body @{
      displayName    = $DisplayName
      signInAudience = 'AzureADMyOrg'
      description    = $appDescription
      notes          = $appNotes
      info           = $desiredInfo
      web            = @{ homePageUrl = $repoUrl }
      api            = @{ requestedAccessTokenVersion = $desiredTokenVersion }
    } `
      -ErrorAction Stop
    $clientId = $app.appId
    $objectId = $app.id
    Write-Host "  Created — Client ID: $clientId" -ForegroundColor Green
  }
  else {
    $clientId = '<not-created-in-WhatIf-mode>'
    $objectId = '<not-created-in-WhatIf-mode>'
  }
}

# In WhatIf mode, if read access was unavailable the GET returned null and
# creation was only simulated.  Inject a stub so every downstream PATCH check
# evaluates and emits its own "What if:" message.
if ($_whatIf -and -not $app) {
  $app = [PSCustomObject]@{
    description    = $null
    notes          = $null
    info           = [PSCustomObject]@{
      termsOfServiceUrl   = $null
      privacyStatementUrl = $null
      supportUrl          = $null
      marketingUrl        = $null
    }
    web            = [PSCustomObject]@{ homePageUrl = $null }
    signInAudience = 'AzureADMyOrg'
    identifierUris = @()
    api            = [PSCustomObject]@{ requestedAccessTokenVersion = $null }
  }
}
#endregion

#region Converge to desired state
# ── Converge to desired state (idempotent) ────────────────────────────────────
$changes = @()

# 1. Description
if ($app.description -ne $appDescription) {
  Write-Host "  Fixing Description" -ForegroundColor Yellow
  if ($PSCmdlet.ShouldProcess($DisplayName, 'PATCH description')) {
    Invoke-MgGraphRequest -Method PATCH `
      -Uri "https://graph.microsoft.com/v1.0/applications/$objectId" `
      -Body @{ description = $appDescription } -ErrorAction Stop
  }
  $changes += 'Description'
}

# 2. Notes
if ($app.notes -ne $appNotes) {
  Write-Host "  Fixing Notes" -ForegroundColor Yellow
  if ($PSCmdlet.ShouldProcess($DisplayName, 'PATCH notes')) {
    Invoke-MgGraphRequest -Method PATCH `
      -Uri "https://graph.microsoft.com/v1.0/applications/$objectId" `
      -Body @{ notes = $appNotes } -ErrorAction Stop
  }
  $changes += 'Notes'
}

# 3. Info URLs (terms, privacy, support, marketing/homepage)
$infoChanged = $false
if ($app.info.termsOfServiceUrl -ne $desiredInfo.termsOfServiceUrl) { $infoChanged = $true }
if ($app.info.privacyStatementUrl -ne $desiredInfo.privacyStatementUrl) { $infoChanged = $true }
if ($app.info.supportUrl -ne $desiredInfo.supportUrl) { $infoChanged = $true }
if ($app.info.marketingUrl -ne $desiredInfo.marketingUrl) { $infoChanged = $true }
if ($infoChanged) {
  Write-Host "  Fixing Info URLs" -ForegroundColor Yellow
  if ($PSCmdlet.ShouldProcess($DisplayName, 'PATCH info URLs')) {
    Invoke-MgGraphRequest -Method PATCH `
      -Uri "https://graph.microsoft.com/v1.0/applications/$objectId" `
      -Body @{ info = $desiredInfo } -ErrorAction Stop
  }
  $changes += 'Info'
}

# 4. Homepage URL (Web section)
if ($app.web.homePageUrl -ne $repoUrl) {
  Write-Host "  Fixing HomePageUrl" -ForegroundColor Yellow
  if ($PSCmdlet.ShouldProcess($DisplayName, 'PATCH web.homePageUrl')) {
    Invoke-MgGraphRequest -Method PATCH `
      -Uri "https://graph.microsoft.com/v1.0/applications/$objectId" `
      -Body @{ web = @{ homePageUrl = $repoUrl } } -ErrorAction Stop
  }
  $changes += 'HomePageUrl'
}

# 5. SignInAudience
if ($app.signInAudience -ne 'AzureADMyOrg') {
  throw ("Existing App Registration '$DisplayName' is not single-tenant " +
    "(SignInAudience=$($app.signInAudience)). " +
    "Configure it to AzureADMyOrg before using this solution.")
}

# 6. Identifier URI
$expectedUri = "api://guest-sponsor-info-proxy/$clientId"
# The null-coalescing operator ?? is PS7+ only; use if/else for PS5.1 compat.
$currentUris = if ($null -eq $app.identifierUris) { @() } else { @($app.identifierUris) }
if ($currentUris -notcontains $expectedUri) {
  Write-Host "  Fixing IdentifierUris: adding $expectedUri" `
    -ForegroundColor Yellow
  if ($PSCmdlet.ShouldProcess($DisplayName, "PATCH identifierUris: $expectedUri")) {
    Invoke-MgGraphRequest -Method PATCH `
      -Uri "https://graph.microsoft.com/v1.0/applications/$objectId" `
      -Body @{ identifierUris = @($expectedUri) } -ErrorAction Stop
  }
  $changes += 'IdentifierUris'
}

# 7. Access-token version
$currentVersion = $app.api.requestedAccessTokenVersion
if ($currentVersion -ne $desiredTokenVersion) {
  Write-Host ("  Fixing RequestedAccessTokenVersion: " +
    "$currentVersion -> $desiredTokenVersion") -ForegroundColor Yellow
  if ($PSCmdlet.ShouldProcess($DisplayName, "PATCH api.requestedAccessTokenVersion: $desiredTokenVersion")) {
    Invoke-MgGraphRequest -Method PATCH `
      -Uri "https://graph.microsoft.com/v1.0/applications/$objectId" `
      -Body @{ api = @{ requestedAccessTokenVersion = $desiredTokenVersion } } -ErrorAction Stop
  }
  $changes += 'RequestedAccessTokenVersion'
}
#endregion

#region Summary and output
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
if ($logoPath) {
  Write-Host "  Uploading logo from $logoPath ..." -ForegroundColor Cyan
  if ($PSCmdlet.ShouldProcess($DisplayName, 'PUT logo')) {
    $logoBytes = [System.IO.File]::ReadAllBytes((Resolve-Path $logoPath))
    Invoke-MgGraphRequest -Method PUT `
      -Uri "https://graph.microsoft.com/v1.0/applications/$objectId/logo" `
      -ContentType 'image/png' `
      -Body $logoBytes `
      -ErrorAction Stop
    Write-Host "  $_chk Logo uploaded." -ForegroundColor Green
  }
}
else {
  # Local file not available (e.g. script run via iwr) — download from GitHub.
  Write-Host "  Downloading logo from $logoRawUrl ..." -ForegroundColor Cyan
  try {
    $logoBytes = (Invoke-WebRequest -Uri $logoRawUrl -UseBasicParsing -ErrorAction Stop).Content
    if ($PSCmdlet.ShouldProcess($DisplayName, 'PUT logo (downloaded from GitHub)')) {
      Invoke-MgGraphRequest -Method PUT `
        -Uri "https://graph.microsoft.com/v1.0/applications/$objectId/logo" `
        -ContentType 'image/png' `
        -Body $logoBytes `
        -ErrorAction Stop
      Write-Host "  $_chk Logo uploaded." -ForegroundColor Green
    }
  }
  catch {
    Write-Host "  $_wrn Logo download failed — skipping. ($_)" `
      -ForegroundColor Yellow
  }
}

# Cache the Client ID so setup-graph-permissions.ps1 can reuse it without
# prompting when run later in the same PowerShell session.
if ($clientId -ne '<not-created-in-WhatIf-mode>') {
  $Global:GsiSetup_WebPartClientId = $clientId
}

# Detect whether setup-graph-permissions.ps1 lives next to this script.
# $PSScriptRoot is empty when the script was run via iwr (scriptblock
# execution) and non-empty when run from a saved local file — use this to
# show the right command to the operator.
$_graphPermScript = $null
if ($PSScriptRoot) {
  $_candidate = Join-Path $PSScriptRoot 'setup-graph-permissions.ps1'
  if (Test-Path $_candidate) {
    $_graphPermScript = $_candidate
  }
}
if ($_graphPermScript) {
  # Local file found — run it directly.
  $_graphPermCmd = "& '$_graphPermScript'"
}
else {
  # Not available locally — provide the iwr one-liner from GitHub.
  $_graphPermCmd = "& ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/setup-graph-permissions.ps1').Content))"
}

# The Deploy to Azure portal URL does not support pre-filled parameters —
# the standard ARM portal form always shows all input fields regardless of
# what is appended to the URL.  Show the Client ID prominently so the user
# can copy-paste it into the correct field during deployment.
# The portal auto-generates the field label from the camelCase parameter name:
# webPartClientId → "Web Part Client Id" (no createUIDefinition needed).
$_deployUrl = 'https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2Fworkoho%2Fspfx-guest-sponsor-info%2Fmain%2Fazure-function%2Finfra%2Fazuredeploy.json'

$_importantLines = @(
  'Copy this Client ID — you will need it in the next step:'
  ''
  "  Web Part Client Id:  $clientId"
)
if ($clientId -ne '<not-created-in-WhatIf-mode>') {
  $_importantLines += ''
  $_importantLines += 'It is also cached for this PowerShell session —'
  $_importantLines += 'setup-graph-permissions.ps1 will reuse it automatically.'
}
Write-Important -Lines $_importantLines

Write-NextStep @(
  'Step 1 — Deploy the Function App to Azure (link below).'
  "         Paste the Client ID above into the 'Web Part Client Id' field."
  ''
  'Step 2 — After deployment completes, return here and run:'
  ''
  "  $_graphPermCmd"
)
Write-Link -Url $_deployUrl -Text '[Deploy to Azure] Guest Sponsor API'
#endregion
