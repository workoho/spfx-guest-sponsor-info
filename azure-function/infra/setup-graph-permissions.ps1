#!/usr/bin/env -S pwsh -NoLogo -NoProfile

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

.PARAMETER Confirm
    -Confirm:$false skips all interactive confirmations and runs unattended.
    When all three parameters are supplied on the command line the script shows
    a summary of planned changes and asks once before connecting; pass
    -Confirm:$false together with all required parameters to bypass that
    prompt completely.

.PARAMETER WhatIf
    Shows what the script would do without making any changes. All write
    operations (POST / PATCH) are skipped; read operations still run so
    resolved values are shown.

.EXAMPLE
    ./setup-graph-permissions.ps1 `
      -ManagedIdentityObjectId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
      -TenantId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
      -FunctionAppClientId "zzzzzzzz-zzzz-zzzz-zzzz-zzzzzzzzzzzz"

.EXAMPLE
    # Fully unattended — no prompts, no confirmation summary
    ./setup-graph-permissions.ps1 `
      -ManagedIdentityObjectId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
      -TenantId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
      -FunctionAppClientId "zzzzzzzz-zzzz-zzzz-zzzz-zzzzzzzzzzzz" `
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
  [string]$ManagedIdentityObjectId,
  [string]$TenantId,
  [string]$FunctionAppClientId
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
      'Required Entra roles:'
      '  - Privileged Role Administrator      (to assign Graph app roles to the Managed Identity)'
      '  - Cloud Application Administrator    (to configure the App Registration)'
      '    (or Application Administrator, or Global Administrator)'
      ''
      'If your roles are eligible (PIM): activate them, then re-run.'
      'If you do not have the roles yet: request them from your admin.'
    )
    Write-Link -Url 'https://entra.microsoft.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/aadRoles' `
      -Text 'PIM → My roles → Entra roles  (activate eligible roles)'
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

# ── Module bootstrap ───────────────────────────────────────────────────
# Only Microsoft.Graph.Authentication is required — all Graph operations
# go through Invoke-MgGraphRequest, which ships in that module.
Install-RequiredModule -Name 'Microsoft.Graph.Authentication'
#endregion

#region Parameter collection
# ── Interactive parameter prompts ─────────────────────────────────────────────
# Each prompt shows a title, a short description, and where to find the value,
# then re-prompts until a valid GUID is entered.
$_guidPattern = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'

if (-not $TenantId) {
  # Session cache: skip the prompt if this script or setup-app-registration.ps1
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
# Save to session cache — available to repeated runs and to
# setup-app-registration.ps1 when run in the same PowerShell session.
$Global:GsiSetup_TenantId = $TenantId

if (-not $ManagedIdentityObjectId) {
  if ($Global:GsiSetup_ManagedIdentityObjectId) {
    $ManagedIdentityObjectId = $Global:GsiSetup_ManagedIdentityObjectId
    Write-Host "  Using cached Managed Identity Object ID from this session: $ManagedIdentityObjectId" -ForegroundColor DarkGray
  }
  else {
    Write-Host '  Required: Function App Managed Identity — Object ID' -ForegroundColor Cyan
    Write-Host $_sep -ForegroundColor DarkGray
    Write-Host '  The Object ID of the system-assigned Managed Identity of the'
    Write-Host '  Azure Function App (a GUID). This is NOT the App Client ID.'
    Write-Host '  Where to find it:'
    Write-Link -Url "https://portal.azure.com/#@$TenantId/blade/HubsExtension/BrowseResource/resourceType/Microsoft.Web%2Fsites" `
      -Text "Azure Portal $_arr Function Apps $_arr [your app] $_arr Settings $_arr Identity"
    Write-Host "    or: deployment outputs $_arr managedIdentityObjectId"
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
    $_promptsShown = $true
  }
}
$Global:GsiSetup_ManagedIdentityObjectId = $ManagedIdentityObjectId

if (-not $FunctionAppClientId) {
  if ($Global:GsiSetup_FunctionAppClientId) {
    $FunctionAppClientId = $Global:GsiSetup_FunctionAppClientId
    Write-Host "  Using cached App Registration Client ID from this session: $FunctionAppClientId" -ForegroundColor DarkGray
  }
  else {
    Write-Host '  Required: App Registration Client ID' -ForegroundColor Cyan
    Write-Host $_sep -ForegroundColor DarkGray
    Write-Host '  The Client ID (Application ID) of the App Registration created'
    Write-Host '  in the previous step (setup-app-registration.ps1). It was'
    Write-Host '  printed at the end of that script.'
    Write-Host '  Where to find it:'
    Write-Link -Url "https://entra.microsoft.com/$TenantId/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade" `
      -Text "Entra admin center $_arr App registrations"
    Write-Host "    'Guest Sponsor Info - SharePoint Web Part Auth' $_arr Application (client) ID"
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
    $_promptsShown = $true
  }
}
$Global:GsiSetup_FunctionAppClientId = $FunctionAppClientId

Write-Hint -Lines @(
  'Required Entra roles:'
  '  - Privileged Role Administrator      (to assign Graph app roles to the Managed Identity)'
  '  - Cloud Application Administrator    (to configure the App Registration)'
  '    (or Application Administrator, or Global Administrator)'
  ''
  'If your roles are eligible (PIM): activate them, then re-run.'
  'If you do not have the roles yet: request them from your admin.'
)
Write-Link -Url "https://entra.microsoft.com/$TenantId/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/aadRoles" `
  -Text 'PIM → My roles → Entra roles  (activate eligible roles)'
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
  Write-Host "  Tenant ID                  : $TenantId"
  Write-Host "  Managed Identity Object ID : $ManagedIdentityObjectId"
  Write-Host "  App Registration Client ID : $FunctionAppClientId"
  Write-Host ''
  Write-Host '  The script will assign Graph app roles to the Managed Identity' -ForegroundColor DarkGray
  Write-Host '  and configure the App Registration for silent token acquisition.' -ForegroundColor DarkGray
  Write-Host '  All operations are idempotent.' -ForegroundColor DarkGray
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
$_requiredScopes = if ($_whatIf) {
  @('AppRoleAssignment.Read.All', 'Application.Read.All')
}
else {
  @('AppRoleAssignment.ReadWrite.All', 'Application.ReadWrite.All')
}
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
    # In WhatIf mode read-only scopes are sufficient — write permissions are not
    # needed for a dry run.  Fall back to offline simulation if sign-in fails.
    Write-Host "  [WhatIf] Connecting with read-only scopes..." -ForegroundColor DarkGray
    try {
      Connect-MgGraph -TenantId $TenantId -Scopes 'AppRoleAssignment.Read.All', 'Application.Read.All' -ErrorAction Stop
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
    Connect-MgGraph -TenantId $TenantId -Scopes 'AppRoleAssignment.ReadWrite.All', 'Application.ReadWrite.All'
  }
}

# ── Verify active Entra role assignments ─────────────────────────────────────
# Opportunistic: requires Directory.Read.All in the current token.
# Silently skipped when the scope is absent or the session is a service
# principal (workload identity).  The check is informational only — the
# script will still fail later with a clear error if the role is missing.
$_checkRoles = @(
  'Privileged Role Administrator'
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
    Write-Host '  Activate your roles via PIM or request them from your admin.' -ForegroundColor Yellow
  }
}
#endregion

#region Graph app role assignments
Write-Host "Resolving Microsoft Graph service principal..." -ForegroundColor Cyan
$graphSpResp = Invoke-MgGraphRequest -Method GET `
  -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '00000003-0000-0000-c000-000000000000'&`$select=id,appRoles" `
  -OutputType PSObject -ErrorAction Stop
$graphSp = if ($graphSpResp.value) { $graphSpResp.value[0] } else { $null }
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

# Fetch all existing app role assignments for the Managed Identity once so
# the loop can skip roles that are already present without posting a duplicate
# and receiving a 400 error.  In WhatIf/offline mode the GET may return $null
# (trap continues) — the hashtable stays empty and ShouldProcess prints the
# "What if:" message for every role, which is the correct dry-run behaviour.
Write-Host "Reading existing app role assignments for the Managed Identity..." -ForegroundColor Cyan
$_existingAssignmentsResp = Invoke-MgGraphRequest -Method GET `
  -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$ManagedIdentityObjectId/appRoleAssignments?`$select=appRoleId" `
  -OutputType PSObject -ErrorAction Stop
# $_existingRoleIds: appRoleId GUID → $true — O(1) look-up in the loop below.
$_existingRoleIds = @{}
if ($_existingAssignmentsResp -and $_existingAssignmentsResp.value) {
  foreach ($_ea in $_existingAssignmentsResp.value) {
    $_existingRoleIds[$_ea.appRoleId] = $true
  }
}

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

  # Skip POST if the role is already assigned — avoids a 400 "already exists"
  # error and counts the role as done rather than skipped.
  if ($_existingRoleIds.ContainsKey($appRole.id)) {
    Write-Host "  $_chk $($role.Name) already assigned — skipping." -ForegroundColor Yellow
    $assignedRoles += $role.Name
    continue
  }

  if ($PSCmdlet.ShouldProcess("Managed Identity $ManagedIdentityObjectId", "POST appRoleAssignment: $($role.Name)")) {
    $null = Invoke-MgGraphRequest -Method POST `
      -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$ManagedIdentityObjectId/appRoleAssignments" `
      -Body @{
      principalId = $ManagedIdentityObjectId
      resourceId  = $graphSp.id
      appRoleId   = $appRole.id
    } -ErrorAction Stop
  }
  Write-Host "  $_chk $($role.Name) assigned." -ForegroundColor Green
  $assignedRoles += $role.Name
}
#endregion

#region App Registration configuration
Write-Host "`nConfiguring App Registration for silent token acquisition by the SharePoint web part..." -ForegroundColor Cyan

# The SharePoint Online Web Client Extensibility app is the MSAL client that SPFx uses
# internally to acquire tokens on behalf of the signed-in user. Pre-authorizing it on the
# EasyAuth App Registration allows silent token acquisition without user consent prompts
# or full-page redirects.
#
# The actual app IDs vary across SharePoint Online environments. We resolve them dynamically
# from the tenant rather than hardcoding, then fall back to the two known canonical IDs.
Write-Host "  Resolving SharePoint Online Web Client Extensibility service principal(s)..." -ForegroundColor Cyan
$spWebClientResp = Invoke-MgGraphRequest -Method GET `
  -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=displayName eq 'SharePoint Online Web Client Extensibility'&`$select=appId" `
  -OutputType PSObject -ErrorAction SilentlyContinue
$spWebClientSps = if ($spWebClientResp -and $spWebClientResp.value) { $spWebClientResp.value } else { $null }
if ($spWebClientSps) {
  $spWebClientAppIds = @($spWebClientSps | Select-Object -ExpandProperty appId)
  Write-Host "  Found $($spWebClientAppIds.Count) SP(s): $($spWebClientAppIds -join ', ')" -ForegroundColor Cyan
}
else {
  # Fall back to the two well-known first-party Microsoft app IDs used across SharePoint Online environments.
  $spWebClientAppIds = @('57fb890c-0dab-4253-a5e0-7188c88b2bb4', '08e18876-6177-487e-b8b5-cf950c1e598c')
  Write-Host "  $_wrn Could not resolve SP by display name — falling back to known app IDs: $($spWebClientAppIds -join ', ')" -ForegroundColor Yellow
}

$appResp = Invoke-MgGraphRequest -Method GET `
  -Uri "https://graph.microsoft.com/v1.0/applications?`$filter=appId eq '$FunctionAppClientId'" `
  -OutputType PSObject -ErrorAction Stop
$app = if ($appResp.value) { $appResp.value[0] } else { $null }
if (-not $app) {
  throw "Could not find App Registration with client ID '$FunctionAppClientId'. Verify the -FunctionAppClientId parameter."
}

# WhatIf stub: if $app is still null (offline simulation), inject a minimal object
# so all downstream checks evaluate and emit their own "What if:" messages.
if ($_whatIf -and -not $app) {
  $app = [PSCustomObject]@{
    id             = '<simulated-object-id>'
    displayName    = 'Guest Sponsor Info - SharePoint Web Part Auth'
    signInAudience = 'AzureADMyOrg'
    identifierUris = @()
    api            = [PSCustomObject]@{
      oauth2PermissionScopes    = @()
      preAuthorizedApplications = @()
    }
    web            = [PSCustomObject]@{ homePageUrl = $null }
  }
}

if ($app.signInAudience -ne 'AzureADMyOrg') {
  throw "App Registration '$FunctionAppClientId' is not single-tenant (SignInAudience=$($app.signInAudience)). Set it to AzureADMyOrg before continuing."
}

# Ensure the identifier URI is set — required for the api:// audience used by EasyAuth.
$expectedUri = "api://guest-sponsor-info-proxy/$FunctionAppClientId"
if ($app.identifierUris -notcontains $expectedUri) {
  Write-Host "  Setting identifier URI to $expectedUri ..." -ForegroundColor Cyan
  if ($PSCmdlet.ShouldProcess($FunctionAppClientId, "PATCH identifierUris: $expectedUri")) {
    Invoke-MgGraphRequest -Method PATCH `
      -Uri "https://graph.microsoft.com/v1.0/applications/$($app.id)" `
      -Body @{ identifierUris = @($expectedUri) } -ErrorAction Stop
  }
  Write-Host "  $_chk Identifier URI set." -ForegroundColor Green
}
else {
  Write-Host "  $_chk Identifier URI already set." -ForegroundColor Yellow
}

# Expose a 'user_impersonation' OAuth2 scope if not already present.
$existingScope = $app.api.oauth2PermissionScopes | Where-Object { $_.value -eq 'user_impersonation' }
if (-not $existingScope) {
  Write-Host "  Adding 'user_impersonation' scope ..." -ForegroundColor Cyan
  $scopeId = [System.Guid]::NewGuid().ToString()
  $newScope = @{
    id                      = $scopeId
    value                   = 'user_impersonation'
    type                    = 'User'
    adminConsentDisplayName = 'Access Guest Sponsor Info web part proxy as the signed-in user'
    adminConsentDescription = 'Allows the SharePoint web part to call the Azure Function proxy on behalf of the signed-in user.'
    userConsentDisplayName  = 'Access Guest Sponsor Info web part proxy'
    userConsentDescription  = 'Allows the app to call the Azure Function proxy on your behalf.'
    isEnabled               = $true
  }
  if ($PSCmdlet.ShouldProcess($FunctionAppClientId, "PATCH api.oauth2PermissionScopes: add user_impersonation")) {
    Invoke-MgGraphRequest -Method PATCH `
      -Uri "https://graph.microsoft.com/v1.0/applications/$($app.id)" `
      -Body @{ api = @{ oauth2PermissionScopes = @($newScope) } } -ErrorAction Stop
    # Re-fetch to get the assigned scope ID (may differ from what we sent).
    $app = Invoke-MgGraphRequest -Method GET `
      -Uri "https://graph.microsoft.com/v1.0/applications/$($app.id)" `
      -OutputType PSObject -ErrorAction Stop
    $existingScope = $app.api.oauth2PermissionScopes | Where-Object { $_.value -eq 'user_impersonation' }
  }
  Write-Host "  $_chk 'user_impersonation' scope added (id: $($existingScope.id))." -ForegroundColor Green
}
else {
  Write-Host "  $_chk 'user_impersonation' scope already exists (id: $($existingScope.id))." -ForegroundColor Yellow
}

# Pre-authorize the SharePoint Online Web Client Extensibility app(s) to call the scope.
# This is what makes token acquisition silent — no per-user consent prompt, no page redirect.
foreach ($spAppId in $spWebClientAppIds) {
  $alreadyPreAuthorized = $app.api.preAuthorizedApplications | Where-Object {
    $_.appId -eq $spAppId -and
    $_.delegatedPermissionIds -contains $existingScope.id
  }
  if (-not $alreadyPreAuthorized) {
    Write-Host "  Pre-authorizing $spAppId ..." -ForegroundColor Cyan
    # Re-fetch the current state before each update to avoid overwriting parallel changes.
    $app = Invoke-MgGraphRequest -Method GET `
      -Uri "https://graph.microsoft.com/v1.0/applications/$($app.id)" `
      -OutputType PSObject -ErrorAction Stop
    $otherPreAuthorized = $app.api.preAuthorizedApplications | Where-Object { $_.appId -ne $spAppId }
    $newPreAuth = @{
      appId                  = $spAppId
      delegatedPermissionIds = @($existingScope.id)
    }
    $updatedPreAuthorized = @($otherPreAuthorized) + @($newPreAuth)
    try {
      if ($PSCmdlet.ShouldProcess($FunctionAppClientId, "PATCH api.preAuthorizedApplications: $spAppId")) {
        Invoke-MgGraphRequest -Method PATCH `
          -Uri "https://graph.microsoft.com/v1.0/applications/$($app.id)" `
          -Body @{ api = @{ preAuthorizedApplications = $updatedPreAuthorized } } -ErrorAction Stop
      }
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
#endregion

#region Enterprise App (Service Principal)
# Ensure appRoleAssignmentRequired is false on the Service Principal (Enterprise App).
# Normally created on first user sign-in, but since we run this script before any user
# has consented, we create it explicitly here.
$spResp = Invoke-MgGraphRequest -Method GET `
  -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$FunctionAppClientId'&`$select=id,appRoleAssignmentRequired,tags,description,notes" `
  -OutputType PSObject -ErrorAction SilentlyContinue
$sp = if ($spResp -and $spResp.value) { $spResp.value[0] } else { $null }
if (-not $sp) {
  Write-Host "  Service Principal not found — creating it now (no user has signed in yet)..." -ForegroundColor Cyan
  if ($PSCmdlet.ShouldProcess($FunctionAppClientId, 'POST servicePrincipal')) {
    $sp = Invoke-MgGraphRequest -Method POST `
      -Uri "https://graph.microsoft.com/v1.0/servicePrincipals" `
      -Body @{ appId = $FunctionAppClientId } `
      -OutputType PSObject -ErrorAction Stop
    Write-Host "  $_chk Service Principal created (Object ID: $($sp.id))." -ForegroundColor Green
  }
}
else {
  Write-Host "  $_chk Service Principal already exists (Object ID: $($sp.id))." -ForegroundColor Yellow
}

# WhatIf stub: SP creation was simulated so $sp is still null.  Inject a stub
# that assumes worst-case defaults so all downstream PATCH checks emit their
# "What if:" messages (appRoleAssignmentRequired=true, no HideApp tag, etc.).
if ($_whatIf -and -not $sp) {
  $sp = [PSCustomObject]@{
    id                        = '<simulated-sp-id>'
    appRoleAssignmentRequired = $true
    tags                      = @()
    description               = $null
    notes                     = $null
  }
}

# appRoleAssignmentRequired=false: all users (including guests) can acquire tokens without
# individual assignment — even with pre-authorization in place.
if ($sp.appRoleAssignmentRequired) {
  Write-Host "  Disabling appRoleAssignmentRequired on the Enterprise App (was: true) ..." -ForegroundColor Cyan
  if ($PSCmdlet.ShouldProcess($FunctionAppClientId, 'PATCH servicePrincipal appRoleAssignmentRequired=false')) {
    Invoke-MgGraphRequest -Method PATCH `
      -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.id)" `
      -Body @{ appRoleAssignmentRequired = $false } -ErrorAction Stop
  }
  Write-Host "  $_chk appRoleAssignmentRequired set to false." -ForegroundColor Green
}
else {
  Write-Host "  $_chk appRoleAssignmentRequired is already false — no user assignment needed." -ForegroundColor Yellow
}

# Hide from My Apps portal (tags: HideApp). This is a backend auth proxy — it should not
# appear as a launchable app in users' My Apps page.
$hasHideApp = $sp.tags -contains 'HideApp'
if (-not $hasHideApp) {
  Write-Host "  Hiding Enterprise App from My Apps portal (visible to users: No) ..." -ForegroundColor Cyan
  $updatedTags = @($sp.tags) + @('HideApp')
  if ($PSCmdlet.ShouldProcess($FunctionAppClientId, 'PATCH servicePrincipal tags: add HideApp')) {
    Invoke-MgGraphRequest -Method PATCH `
      -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.id)" `
      -Body @{ tags = $updatedTags } -ErrorAction Stop
  }
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
  "Paired with Managed Identity: $ManagedIdentityObjectId.",
  'Docs: https://guest-sponsor-info.workoho.cloud'
) -join ' '

# Notes field — visible under Enterprise App → Properties. Ideal for Ops runbook hints.
$spNotes = @(
  'Do not delete — the "Guest Sponsor Info" SharePoint web part depends',
  'on this for guest sponsor lookups via Microsoft Graph.',
  'This app should remain hidden from My Apps (HideApp tag).',
  'The associated Azure Function uses a system-assigned Managed',
  'Identity for Graph API calls (User.Read.All, Presence.Read.All,',
  'MailboxSettings.Read, TeamMember.Read.All).',
  "EasyAuth App Registration: $($app.displayName) (Client ID: $FunctionAppClientId).",
  "Managed Identity Object ID: $ManagedIdentityObjectId.",
  'Docs: https://guest-sponsor-info.workoho.cloud'
) -join ' '

if ($sp.description -ne $spDescription) {
  Write-Host "  Setting Enterprise App description ..." -ForegroundColor Cyan
  if ($PSCmdlet.ShouldProcess($FunctionAppClientId, 'PATCH servicePrincipal description')) {
    Invoke-MgGraphRequest -Method PATCH `
      -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.id)" `
      -Body @{ description = $spDescription } -ErrorAction Stop
  }
  Write-Host "  $_chk Description set." -ForegroundColor Green
}
else {
  Write-Host "  $_chk Enterprise App description already set." -ForegroundColor Yellow
}

if ($sp.notes -ne $spNotes) {
  Write-Host "  Setting Enterprise App notes ..." -ForegroundColor Cyan
  if ($PSCmdlet.ShouldProcess($FunctionAppClientId, 'PATCH servicePrincipal notes')) {
    Invoke-MgGraphRequest -Method PATCH `
      -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.id)" `
      -Body @{ notes = $spNotes } -ErrorAction Stop
  }
  Write-Host "  $_chk Notes set." -ForegroundColor Green
}
else {
  Write-Host "  $_chk Enterprise App notes already set." -ForegroundColor Yellow
}
# Service Management Reference — shown under App Registration → Properties.
# Points to the GitHub Issues tracker so Ops teams know where to file tickets.
# Note: serviceManagementReference is a property of the Application object (not the
# ServicePrincipal). The Graph API silently rejects PATCH on the SP endpoint with 404,
# so we read and write via the /applications/ endpoint using $app.id.
$desiredSmRef = 'https://github.com/workoho/spfx-guest-sponsor-info/issues'
$currentSmRef = (Invoke-MgGraphRequest -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/applications/$($app.id)?`$select=serviceManagementReference" `
    -ErrorAction Stop).serviceManagementReference
if ($currentSmRef -ne $desiredSmRef) {
  Write-Host "  Setting Service Management Reference ..." -ForegroundColor Cyan
  if ($PSCmdlet.ShouldProcess($FunctionAppClientId, 'PATCH application serviceManagementReference')) {
    Invoke-MgGraphRequest -Method PATCH `
      -Uri "https://graph.microsoft.com/v1.0/applications/$($app.id)" `
      -Body @{ serviceManagementReference = $desiredSmRef } -ErrorAction Stop
  }
  Write-Host "  $_chk Service Management Reference set." -ForegroundColor Green
}
else {
  Write-Host "  $_chk Service Management Reference already set." -ForegroundColor Yellow
}

# Homepage URL — visible under Enterprise App → Properties.
# Graph requires homePageUrl to be set on the Application object (web.homePageUrl);
# the ServicePrincipal mirrors it automatically. Setting it directly on the SP fails
# with "does not match the application object" (400). Read current value via $app.
$desiredHomepage = 'https://github.com/workoho/spfx-guest-sponsor-info'
if ($app.web.homePageUrl -ne $desiredHomepage) {
  Write-Host "  Setting App Registration homepage URL ..." -ForegroundColor Cyan
  if ($PSCmdlet.ShouldProcess($FunctionAppClientId, 'PATCH application web.homePageUrl')) {
    Invoke-MgGraphRequest -Method PATCH `
      -Uri "https://graph.microsoft.com/v1.0/applications/$($app.id)" `
      -Body @{ web = @{ homePageUrl = $desiredHomepage } } -ErrorAction Stop
  }
  Write-Host "  $_chk Homepage URL set." -ForegroundColor Green
}
else {
  Write-Host "  $_chk App Registration homepage URL already set." -ForegroundColor Yellow
}
#endregion

#region Summary
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
#endregion
