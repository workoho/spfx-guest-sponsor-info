#!/usr/bin/env -S pwsh -NoLogo -NoProfile

<#
.SYNOPSIS
    Guided console deployment for the Guest Sponsor Info Azure infrastructure.

.DESCRIPTION
    Detects which deployment tools are already available locally and then offers
    three console-based deployment paths:

      - azd provision  (preferred when azd and Azure CLI are already available)
      - Bicep via Azure CLI
      - ARM JSON via Azure CLI

    The default suggestion balances the preferred modern workflow with the
    tools that are already installed, so users can avoid unnecessary setup.

    When the script is run via iwr, required repository assets are downloaded
    temporarily from GitHub and removed again when the script finishes.

.PARAMETER Mode
    Auto (default), Azd, Bicep, or ArmJson.

.PARAMETER ResourceGroupName
    Target resource group for the direct Bicep and ARM JSON paths.

.PARAMETER TenantName
    SharePoint tenant name without the .sharepoint.com suffix.

.PARAMETER FunctionAppName
    Globally unique Function App name.

.PARAMETER WebPartClientId
    Existing EasyAuth App Registration client ID. If omitted, the script can
    create or reuse the registration by calling setup-app-registration.ps1.

.PARAMETER HostingPlan
    Consumption (default) or FlexConsumption.

.PARAMETER DeployAzureMaps
    Include Azure Maps in the deployment.

.PARAMETER AppVersion
    Function package version. Defaults to latest.

.PARAMETER EnableMonitoring
  Deploy Log Analytics, Application Insights, managed Failure Anomalies rule,
  and the repository's KQL alert resources. Defaults to true.

.PARAMETER EnableFailureAnomaliesAlert
  Enable the Application Insights Failure Anomalies smart detector alert
  rule. Defaults to false.

.PARAMETER MaximumFlexInstances
    Hard scale-out cap for Flex Consumption. Defaults to 10.

.EXAMPLE
    ./azure-function/infra/deploy-azure.ps1

.EXAMPLE
    & ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/deploy-azure.ps1').Content))

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
  [ValidateSet('Auto', 'Azd', 'Bicep', 'ArmJson')]
  [string]$Mode = 'Auto',
  [string]$ResourceGroupName,
  [string]$TenantName,
  [string]$FunctionAppName,
  [string]$WebPartClientId,
  [ValidateSet('Consumption', 'FlexConsumption')]
  [string]$HostingPlan = 'Consumption',
  [bool]$DeployAzureMaps = $true,
  [string]$AppVersion = 'latest',
  [bool]$EnableMonitoring = $true,
  [bool]$EnableFailureAnomaliesAlert = $false,
  [int]$MaximumFlexInstances = 10
)

$ErrorActionPreference = 'Stop'

# Track whether any interactive prompt was shown. When all parameters were
# pre-supplied (via the command line) we show a confirmation summary so the
# operator can verify before the script runs.
$_promptsShown = $false
# Convenience bool used throughout for WhatIf-aware fallbacks.
$_whatIf = $WhatIfPreference -eq [System.Management.Automation.SwitchParameter]$true

$script:RepoRawBaseUrl = 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main'
$script:AppRegistrationDisplayName = 'Guest Sponsor Info - SharePoint Web Part Auth'
$script:TempPaths = [System.Collections.Generic.List[string]]::new()
$script:StagedRepoRoot = $null
$script:AzPath = $null
$script:AzdPath = $null
$script:SubscriptionName = ''
$script:SubscriptionId = ''
$script:TenantId = ''
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
#                  hand data back to its caller.
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
# Script-level trap: on Azure CLI authorization errors (403), print role
# guidance instead of a raw exception. Other errors re-throw normally.
trap {
  $_errMsg = if ($_.Exception -and $_.Exception.Message) { $_.Exception.Message } else { [string]$_ }
  if ($_errMsg -match '(?i)AuthorizationFailed|does not have authorization|403|Forbidden') {
    Write-Failure -Lines @(
      'The request was denied — your account lacks the required permissions.'
      ''
      'Required Azure RBAC role:  Contributor  (on the target subscription or resource group)'
      '  Owner also works. For provider registration: Contributor or higher at subscription level.'
      ''
      'If your role is eligible (PIM): activate it, then re-run.'
      'If you do not have the role yet: request it from your Azure admin.'
    )
    Write-Link -Url 'https://portal.azure.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/azurerbac' `
      -Text 'PIM → My roles → Azure resources  (activate eligible role)'
    Write-Link -Url 'https://portal.azure.com/#view/Microsoft_Azure_AD_IAM/ActiveDirectoryMenuBlade/~/RolesAndAdministrators' `
      -Text 'Azure portal → Subscriptions → Access control (IAM)'
    return
  }
  # Not a permission error — let PowerShell display the raw error and exit.
}
#endregion

#region Tool detection helpers
function Test-WindowsHost {
  return ($PSVersionTable.PSVersion.Major -lt 6) -or ($env:OS -eq 'Windows_NT')
}

function Update-ProcessPathFromSystem {
  [CmdletBinding(SupportsShouldProcess)]
  param()

  if (-not (Test-WindowsHost)) {
    return
  }

  $pathParts = @(
    [System.Environment]::GetEnvironmentVariable('Path', 'Machine'),
    [System.Environment]::GetEnvironmentVariable('Path', 'User')
  ) | Where-Object { $_ }

  if (($pathParts.Count -gt 0) -and $PSCmdlet.ShouldProcess('process PATH', 'refresh from system environment')) {
    $env:Path = $pathParts -join ';'
  }
}

function Get-AzureCliPath {
  Update-ProcessPathFromSystem

  if (Test-WindowsHost) {
    $programFiles = $env:ProgramFiles
    if ($programFiles) {
      $candidate = Join-Path -Path $programFiles -ChildPath 'Microsoft SDKs\Azure\CLI2\wbin\az.cmd'
      if (Test-Path -Path $candidate) {
        return $candidate
      }
    }
  }

  $azCommand = Get-Command -Name az -ErrorAction SilentlyContinue
  if ($azCommand) {
    return $azCommand.Source
  }

  return $null
}

function Get-AzdPath {
  Update-ProcessPathFromSystem

  $azdCommand = Get-Command -Name azd -ErrorAction SilentlyContinue
  if ($azdCommand) {
    return $azdCommand.Source
  }

  return $null
}

function Invoke-AzureCli {
  param([Parameter(ValueFromRemainingArguments)][string[]]$Arguments)

  & $script:AzPath @Arguments
  if ($LASTEXITCODE -ne 0) {
    throw "Azure CLI failed (exit code $LASTEXITCODE): az $($Arguments -join ' ')"
  }
}
#endregion

function Invoke-Azd {
  param([Parameter(ValueFromRemainingArguments)][string[]]$Arguments)

  & $script:AzdPath @Arguments
  if ($LASTEXITCODE -ne 0) {
    throw "Azure Developer CLI failed (exit code $LASTEXITCODE): azd $($Arguments -join ' ')"
  }
}

function Test-BicepReady {
  if (-not $script:AzPath) {
    return $false
  }

  try {
    Invoke-AzureCli -Arguments @('bicep', 'version') | Out-Null
    return $true
  }
  catch {
    return $false
  }
}

function Install-AzureCliIfNeeded {
  $script:AzPath = Get-AzureCliPath
  if ($script:AzPath) {
    return
  }

  if (-not (Test-WindowsHost)) {
    throw @(
      'Azure CLI (az) is not installed.',
      'Install it first and re-run this script.',
      'Docs: https://learn.microsoft.com/cli/azure/install-azure-cli'
    ) -join ' '
  }

  $winget = Get-Command -Name winget -ErrorAction SilentlyContinue
  if (-not $winget) {
    throw @(
      'Azure CLI (az) is not installed and winget is not available.',
      'Install it manually and re-run this script.',
      'Recommended Windows path: winget install --exact --id Microsoft.AzureCLI'
    ) -join ' '
  }

  $answer = (Read-Host -Prompt 'Azure CLI is required. Install it now via winget? [Y/n]').Trim()
  if ($answer -ne '' -and $answer -notmatch '^[Yy]') {
    throw 'Azure CLI is required for all supported deployment methods in this repository.'
  }

  Write-Host "  $_arr Installing Azure CLI via winget..." -ForegroundColor Cyan
  # --scope user avoids requiring administrator elevation.
  winget install --exact --id Microsoft.AzureCLI --scope user --accept-source-agreements --accept-package-agreements | Out-Host
  $script:AzPath = Get-AzureCliPath

  if (-not $script:AzPath) {
    throw 'Azure CLI was installed, but the current session still cannot find az. Open a new terminal and re-run the script.'
  }

  Write-Host "  $_chk Azure CLI is available." -ForegroundColor Green
}

function Install-AzdIfNeeded {
  $script:AzdPath = Get-AzdPath
  if ($script:AzdPath) {
    return
  }

  Write-Host "  $_wrn Azure Developer CLI (azd) is not installed." -ForegroundColor Yellow

  if (Test-WindowsHost) {
    # ── Windows: use winget ────────────────────────────────────────────────
    $winget = Get-Command -Name winget -ErrorAction SilentlyContinue
    if (-not $winget) {
      throw @(
        'Azure Developer CLI (azd) is not installed and winget is not available.',
        'Install it manually and re-run this script.',
        'Docs: https://learn.microsoft.com/azure/developer/azure-developer-cli/install-azd'
      ) -join ' '
    }

    $answer = (Read-Host '  Install it now via winget? [Y/n]').Trim()
    if ($answer -ne '' -and $answer -notmatch '^[Yy]') {
      throw 'The azd deployment path requires Azure Developer CLI.'
    }

    Write-Host "  $_arr Installing Azure Developer CLI via winget..." -ForegroundColor Cyan
    # --scope user avoids requiring administrator elevation.
    winget install microsoft.azd --scope user --accept-source-agreements --accept-package-agreements | Out-Host
  }
  else {
    # ── Linux / macOS ─────────────────────────────────────────────────────
    # Prefer Homebrew when available (works on both macOS and Linux).
    # Fall back to the official Microsoft install script distributed via
    # https://aka.ms/install-azd.sh — curl pipes the script into bash, which
    # installs azd to /usr/local/bin on most distributions.
    $brew = Get-Command -Name brew -ErrorAction SilentlyContinue
    $curl = Get-Command -Name curl -ErrorAction SilentlyContinue

    if ($brew) {
      $answer = (Read-Host '  Install it now via Homebrew? [Y/n]').Trim()
      if ($answer -ne '' -and $answer -notmatch '^[Yy]') {
        throw 'The azd deployment path requires Azure Developer CLI.'
      }
      Write-Host "  $_arr Installing Azure Developer CLI via Homebrew..." -ForegroundColor Cyan
      & brew tap azure/azd
      if ($LASTEXITCODE -ne 0) { throw "brew tap azure/azd failed (exit $LASTEXITCODE)." }
      & brew install azd
      if ($LASTEXITCODE -ne 0) { throw "brew install azd failed (exit $LASTEXITCODE)." }
    }
    elseif ($curl) {
      $answer = (Read-Host '  Install it now via the official Microsoft install script? [Y/n]').Trim()
      if ($answer -ne '' -and $answer -notmatch '^[Yy]') {
        throw 'The azd deployment path requires Azure Developer CLI.'
      }
      Write-Host "  $_arr Installing Azure Developer CLI via install script (curl | bash)..." -ForegroundColor Cyan
      # Official Microsoft install script for azd on Linux and macOS.
      # See: https://learn.microsoft.com/azure/developer/azure-developer-cli/install-azd
      & bash -c 'curl -fsSL https://aka.ms/install-azd.sh | bash'
      if ($LASTEXITCODE -ne 0) { throw "azd installation script failed (exit $LASTEXITCODE)." }

      # The install script may place azd in ~/.local/bin on some distros
      # (e.g. when run without root). Extend PATH for this session so that
      # Get-AzdPath can find the binary immediately without a new terminal.
      foreach ($_extraDir in @("$env:HOME/.local/bin", "$env:HOME/bin")) {
        if ((Test-Path $_extraDir) -and $env:PATH -notlike "*$_extraDir*") {
          $env:PATH = $env:PATH + ':' + $_extraDir
        }
      }
    }
    else {
      throw @(
        'Azure Developer CLI (azd) is not installed and neither Homebrew nor curl is available.',
        'Install it manually: https://learn.microsoft.com/azure/developer/azure-developer-cli/install-azd'
      ) -join ' '
    }
  }

  $script:AzdPath = Get-AzdPath
  if (-not $script:AzdPath) {
    throw 'Azure Developer CLI was installed, but the current session still cannot find azd. Open a new terminal and re-run the script.'
  }

  Write-Host "  $_chk Azure Developer CLI is available." -ForegroundColor Green
}

function Install-BicepCliIfNeeded {
  if (Test-BicepReady) {
    return
  }

  $answer = (Read-Host -Prompt 'Bicep is not available yet. Install it now via az bicep install? [Y/n]').Trim()
  if ($answer -ne '' -and $answer -notmatch '^[Yy]') {
    throw 'The Bicep deployment path requires Bicep CLI.'
  }

  Write-Host "  $_arr Installing Bicep via Azure CLI..." -ForegroundColor Cyan
  Invoke-AzureCli -Arguments @('bicep', 'install') | Out-Null
  Write-Host "  $_chk Bicep CLI is available." -ForegroundColor Green
}

function Connect-AzureCliIfNeeded {
  try {
    Invoke-AzureCli -Arguments @('account', 'show', '--output', 'none') | Out-Null
  }
  catch {
    Write-Host "  $_arr No active Azure CLI session found. Starting az login..." -ForegroundColor Cyan
    # If setup-app-registration.ps1 already ran in this session it will have
    # stored the Entra tenant ID. Pass it as a hint so az login lands on the
    # right tenant without asking the operator to pick one manually.
    if ($Global:GsiSetup_TenantId) {
      Invoke-AzureCli -Arguments @('login', '--tenant', $Global:GsiSetup_TenantId) | Out-Null
    }
    else {
      Invoke-AzureCli -Arguments @('login') | Out-Null
    }
  }

  $script:SubscriptionName = (Invoke-AzureCli -Arguments @('account', 'show', '--query', 'name', '-o', 'tsv')).Trim()
  $script:SubscriptionId = (Invoke-AzureCli -Arguments @('account', 'show', '--query', 'id', '-o', 'tsv')).Trim()
  $script:TenantId = (Invoke-AzureCli -Arguments @('account', 'show', '--query', 'tenantId', '-o', 'tsv')).Trim()

  # Publish the tenant ID into the shared session cache so downstream scripts
  # (e.g. setup-graph-permissions.ps1) can skip their own login prompts.
  $Global:GsiSetup_TenantId = $script:TenantId
}

function Get-ToolState {
  $script:AzPath = Get-AzureCliPath
  $script:AzdPath = Get-AzdPath

  $azReady = [bool]$script:AzPath
  $bicepReady = $false
  if ($azReady) {
    $bicepReady = Test-BicepReady
  }

  return [pscustomobject]@{
    AzureCliReady = $azReady
    AzdReady      = [bool]$script:AzdPath -and $azReady
    AzdInstalled  = [bool]$script:AzdPath
    BicepReady    = $bicepReady
  }
}

function Select-DefaultMode {
  param([Parameter(Mandatory)][pscustomobject]$ToolState)

  # If setup-app-registration.ps1 (Step 1) was already executed in this
  # session the Entra App Registration is done. Recommending azd would cause
  # the pre-provision hook to repeat Step 1 unnecessarily — prefer a direct
  # deployment path instead.
  # NOTE: intentionally using AppRegistrationDone, not GsiSetup_TenantId.
  # GsiSetup_TenantId is also written by Connect-AzureCliIfNeeded in this
  # script, so it is not a reliable indicator that Step 1 actually ran.
  $_step1Done = [bool]$Global:GsiSetup_AppRegistrationDone

  if ($ToolState.AzdReady -and -not $_step1Done) {
    return 'Azd'
  }

  if ($ToolState.BicepReady) {
    return 'Bicep'
  }

  if ($ToolState.AzureCliReady) {
    return 'ArmJson'
  }

  # Fallback: Bicep (az bicep install will be offered during deployment).
  return 'Bicep'
}

function Get-DefaultModeReason {
  param(
    [Parameter(Mandatory)][string]$SelectedMode,
    [Parameter(Mandatory)][pscustomobject]$ToolState
  )

  $_step1Done = [bool]$Global:GsiSetup_AppRegistrationDone

  switch ($SelectedMode) {
    'Azd' {
      return 'azd and Azure CLI are already available, and this repository contains an azd workflow with pre- and post-provision hooks that handle all 3 setup steps automatically.'
    }
    'Bicep' {
      if ($_step1Done -and $ToolState.BicepReady) {
        return 'Step 1 (App Registration) was already completed in this session — a direct Bicep deployment skips the pre-provision hook and avoids repeating it.'
      }
      if ($_step1Done) {
        return 'Step 1 (App Registration) was already completed in this session — azd would re-run it via the pre-provision hook, so a direct path is preferred.'
      }
      if ($ToolState.BicepReady) {
        return 'Azure CLI and Bicep are already available, so the preferred direct CLI path is ready immediately.'
      }
      return 'Azure CLI plus az bicep install is the smallest modern install path when nothing is ready yet.'
    }
    'ArmJson' {
      if ($_step1Done) {
        return 'Step 1 (App Registration) was already completed in this session — ARM JSON is the fastest direct path on this machine without Bicep.'
      }
      return 'Azure CLI is already available, so ARM JSON is the fastest no-install fallback on this machine.'
    }
    default {
      return 'Automatic selection could not determine a better default.'
    }
  }
}

function Read-DefaultValue {
  param(
    [Parameter(Mandatory)][string]$Prompt,
    [Parameter(Mandatory)][string]$DefaultValue
  )

  $value = (Read-Host -Prompt "$Prompt [$DefaultValue]").Trim()
  if ($value) {
    return $value
  }

  return $DefaultValue
}

function Get-DetectedTenantName {
  try {
    $derivedName = Invoke-AzureCli -Arguments @(
      'rest',
      '--method', 'GET',
      '--url', 'https://graph.microsoft.com/v1.0/organization?$select=verifiedDomains',
      '--query', 'value[0].verifiedDomains[?isInitial].name | [0]',
      '-o', 'tsv'
    )
    return ($derivedName.Trim() -replace '\.onmicrosoft\.com$', '')
  }
  catch {
    return ''
  }
}

function Get-LocalRepoRoot {
  if (-not $PSScriptRoot) {
    return $null
  }

  $candidate = Resolve-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath '../..') -ErrorAction SilentlyContinue
  if (-not $candidate) {
    return $null
  }

  $azureYamlPath = Join-Path -Path $candidate.Path -ChildPath 'azure.yaml'
  $mainBicepPath = Join-Path -Path $candidate.Path -ChildPath 'azure-function/infra/main.bicep'
  if ((Test-Path -Path $azureYamlPath) -and (Test-Path -Path $mainBicepPath)) {
    return $candidate.Path
  }

  return $null
}

function Save-RepoFile {
  param(
    [Parameter(Mandatory)][string]$RepoRoot,
    [Parameter(Mandatory)][string]$RelativePath
  )

  $destinationPath = Join-Path -Path $RepoRoot -ChildPath $RelativePath
  $destinationDirectory = Split-Path -Path $destinationPath -Parent
  if (-not (Test-Path -Path $destinationDirectory)) {
    New-Item -ItemType Directory -Path $destinationDirectory -Force | Out-Null
  }

  $downloadPath = $RelativePath -replace '\\', '/'
  $downloadUrl = "$($script:RepoRawBaseUrl)/$downloadPath"
  Write-Host "  $_arr Downloading $RelativePath from GitHub..." -ForegroundColor Cyan
  Invoke-WebRequest -Uri $downloadUrl -OutFile $destinationPath -UseBasicParsing -ErrorAction Stop
}

function Get-RepoRoot {
  $localRepoRoot = Get-LocalRepoRoot
  if ($localRepoRoot) {
    return $localRepoRoot
  }

  if ($script:StagedRepoRoot) {
    return $script:StagedRepoRoot
  }

  $script:StagedRepoRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("gsi-infra-" + [guid]::NewGuid().ToString('n'))
  New-Item -ItemType Directory -Path $script:StagedRepoRoot -Force | Out-Null
  $script:TempPaths.Add($script:StagedRepoRoot)

  $requiredFiles = @(
    'azure.yaml',
    'azure-function/infra/azuredeploy.json',
    'azure-function/infra/deploy-azure.ps1',
    'azure-function/infra/hooks/post-provision.ps1',
    'azure-function/infra/hooks/post-provision.sh',
    'azure-function/infra/hooks/pre-provision.ps1',
    'azure-function/infra/hooks/pre-provision.sh',
    'azure-function/infra/main.bicep',
    'azure-function/infra/modules/monitoring.bicep',
    'azure-function/infra/setup-app-registration.ps1'
  )

  foreach ($relativePath in $requiredFiles) {
    Save-RepoFile -RepoRoot $script:StagedRepoRoot -RelativePath $relativePath
  }

  Write-Host "  $_chk Temporary repo assets staged at $($script:StagedRepoRoot)." -ForegroundColor Green
  return $script:StagedRepoRoot
}

function Get-RepoFilePath {
  param([Parameter(Mandatory)][string]$RelativePath)

  $repoRoot = Get-RepoRoot
  $resolvedPath = Join-Path -Path $repoRoot -ChildPath $RelativePath
  if (-not (Test-Path -Path $resolvedPath)) {
    throw "Required file not found: $RelativePath"
  }

  return (Resolve-Path -Path $resolvedPath).Path
}

function Select-AzureSubscription {
  # List all subscriptions accessible to the signed-in account. If more than
  # one is available let the operator confirm or switch before deployment
  # proceeds — relying on whatever az happens to have set as default is not
  # acceptable when deploying production infrastructure.
  $_subs = $null
  try {
    $_rawJson = Invoke-AzureCli -Arguments @(
      'account', 'list',
      '--query', '[].{name:name,id:id,tenantId:tenantId,isDefault:isDefault}',
      '-o', 'json'
    )
    $_subs = $_rawJson | ConvertFrom-Json
  }
  catch {
    # Listing failed — carry on with whatever account is currently active.
    return
  }

  if (-not $_subs -or $_subs.Count -eq 0) { return }

  if ($_subs.Count -eq 1) {
    # Only one subscription — auto-select without prompting.
    $script:SubscriptionName = $_subs[0].name
    $script:SubscriptionId = $_subs[0].id
    $script:TenantId = $_subs[0].tenantId
    return
  }

  Write-Host ''
  Write-Host '  Azure Subscription' -ForegroundColor Cyan
  Write-Host $_sep -ForegroundColor DarkGray
  Write-Host '  Multiple subscriptions are accessible. Choose which one to deploy into.'
  Write-Host ''

  # Limit the visible list to avoid flooding the console on large tenants.
  # Microsoft tools (az, azd) show at most ~10 entries and ask for an ID.
  $_maxShown = 10
  $_showCount = [Math]::Min($_subs.Count, $_maxShown)

  for ($_i = 0; $_i -lt $_showCount; $_i++) {
    $_s = $_subs[$_i]
    # Mark the currently active subscription for easy identification.
    $_tag = if ($_s.isDefault) { '  (current)' } else { '' }
    Write-Host "    [$($_i + 1)] $($_s.name)$_tag" `
      -ForegroundColor $(if ($_s.isDefault) { 'Green' } else { 'White' })
    Write-Host "         $($_s.id)" -ForegroundColor DarkGray
  }

  if ($_subs.Count -gt $_maxShown) {
    $_hidden = $_subs.Count - $_maxShown
    Write-Host ''
    Write-Host "    ... and $_hidden more. Paste the subscription ID directly to use one not listed." `
      -ForegroundColor DarkGray
  }
  Write-Host ''
  # Pre-select whichever subscription is currently marked as default.
  $_defaultIdx = 1
  for ($_i = 0; $_i -lt $_subs.Count; $_i++) {
    if ($_subs[$_i].isDefault) { $_defaultIdx = $_i + 1; break }
  }

  do {
    $_choice = (Read-Host "  Subscription [default: $_defaultIdx, or paste a subscription ID]").Trim()
    if ($_choice -eq '') { $_choice = [string]$_defaultIdx }

    # Accept either a list number or a raw subscription GUID/ID.
    if ($_choice -match '^\d+$') {
      # Numeric input — must be within the displayed range.
      if ([int]$_choice -lt 1 -or [int]$_choice -gt $_showCount) {
        Write-Host "  $_wrn Enter a number between 1 and $_showCount, or paste a subscription ID." `
          -ForegroundColor Yellow
        $_choice = ''
      }
    }
    elseif ($_choice -match '^[0-9a-fA-F-]{36}$') {
      # GUID-shaped input — look it up in the list. Subscriptions not in the
      # displayed list (e.g. from a very large tenant) are accepted as-is; az
      # will reject the call if the ID is actually invalid.
      $_matchIdx = -1
      for ($_i = 0; $_i -lt $_subs.Count; $_i++) {
        if ($_subs[$_i].id -eq $_choice) { $_matchIdx = $_i; break }
      }
      if ($_matchIdx -ge 0) {
        # Convert to a list number so the shared resolution logic below works.
        $_choice = [string]($_matchIdx + 1)
      }
      # else: leave $_choice as the raw GUID — handled after the loop.
    }
    else {
      Write-Host "  $_wrn Enter a list number or paste a valid subscription ID (GUID)." `
        -ForegroundColor Yellow
      $_choice = ''
    }
  } while (-not $_choice)
  Write-Host ''

  # Resolve the chosen entry. $_choice is either a list number or a raw GUID
  # for a subscription that was not in the displayed list.
  if ($_choice -match '^\d+$') {
    $_selected = $_subs[[int]$_choice - 1]
    # Only call az account set if the selection differs from the current default.
    if ($_selected.id -ne $script:SubscriptionId) {
      Invoke-AzureCli -Arguments @('account', 'set', '--subscription', $_selected.id) | Out-Null
    }
    $script:SubscriptionName = $_selected.name
    $script:SubscriptionId = $_selected.id
    $script:TenantId = $_selected.tenantId
  }
  else {
    # Raw GUID — the operator typed a subscription ID not visible in the list.
    Invoke-AzureCli -Arguments @('account', 'set', '--subscription', $_choice) | Out-Null
    $script:SubscriptionId = $_choice
    # Re-read name and tenant from az now that the account is switched.
    $script:SubscriptionName = (Invoke-AzureCli -Arguments @(
        'account', 'show', '--query', 'name', '-o', 'tsv')).Trim()
    $script:TenantId = (Invoke-AzureCli -Arguments @(
        'account', 'show', '--query', 'tenantId', '-o', 'tsv')).Trim()
  }
}

function Initialize-DeploymentMode {
  param([Parameter(Mandatory)][string]$SelectedMode)

  switch ($SelectedMode) {
    'Azd' {
      Install-AzureCliIfNeeded
      Install-AzdIfNeeded
      Connect-AzureCliIfNeeded
      # Tell azd to reuse the Azure CLI token so the user is not prompted to
      # log in a second time via a separate azd browser window.
      Invoke-Azd -Arguments @('config', 'set', 'auth.useAzureCliCredentials', 'true')
    }
    'Bicep' {
      Install-AzureCliIfNeeded
      Connect-AzureCliIfNeeded
      Install-BicepCliIfNeeded
    }
    'ArmJson' {
      Install-AzureCliIfNeeded
      Connect-AzureCliIfNeeded
    }
  }
}

function Initialize-ResourceGroup {
  param([Parameter(Mandatory)][string]$Name)

  $exists = (Invoke-AzureCli -Arguments @('group', 'exists', '--name', $Name)).Trim()
  if ($exists -eq 'true') {
    Write-Host "  $_chk Resource group $Name already exists." -ForegroundColor Green
    return
  }

  if ($_whatIf) {
    Write-Host "  [WhatIf] Would create resource group $Name." -ForegroundColor DarkGray
    return
  }

  $location = Read-DefaultValue -Prompt 'Azure location for the new resource group' -DefaultValue 'westeurope'
  Write-Host "  $_arr Creating resource group $Name in $location..." -ForegroundColor Cyan
  Invoke-AzureCli -Arguments @('group', 'create', '--name', $Name, '--location', $location) | Out-Null
  Write-Host "  $_chk Resource group $Name created." -ForegroundColor Green
}

function Invoke-ProviderPreflight {
  param(
    [Parameter(Mandatory)][string]$SelectedHostingPlan,
    [Parameter(Mandatory)][bool]$DeployMaps,
    [Parameter(Mandatory)][bool]$MonitoringEnabled
  )

  $requiredProviders = @(
    'Microsoft.Authorization',
    'Microsoft.ManagedIdentity',
    'Microsoft.Resources',
    'Microsoft.Storage',
    'Microsoft.Web'
  )

  if ($MonitoringEnabled) {
    $requiredProviders += @(
      'Microsoft.AlertsManagement',
      'Microsoft.Insights',
      'Microsoft.OperationalInsights'
    )
  }

  if ($SelectedHostingPlan -eq 'FlexConsumption') {
    $requiredProviders += 'Microsoft.ContainerInstance'
  }

  if ($DeployMaps) {
    $requiredProviders += 'Microsoft.Maps'
  }

  $requiredProviders = $requiredProviders | Sort-Object -Unique
  $missingProviders = @()

  Write-Host ''
  Write-Host '  Provider preflight' -ForegroundColor Cyan
  Write-Host $_sep -ForegroundColor DarkGray
  foreach ($provider in $requiredProviders) {
    $state = ''
    try {
      $state = (Invoke-AzureCli -Arguments @('provider', 'show', '--namespace', $provider, '--query', 'registrationState', '-o', 'tsv')).Trim()
    }
    catch {
      $state = ''
    }

    switch ($state) {
      'Registered' {
        Write-Host "  $_chk $provider is registered." -ForegroundColor Green
      }
      'Registering' {
        Write-Host "  $_wrn $provider is still registering. Deployment can usually continue." -ForegroundColor Yellow
      }
      'NotRegistered' {
        Write-Host "  $_wrn $provider is not registered." -ForegroundColor Yellow
        $missingProviders += $provider
      }
      'Unregistered' {
        Write-Host "  $_wrn $provider is not registered." -ForegroundColor Yellow
        $missingProviders += $provider
      }
      default {
        Write-Host "  $_wrn $provider returned state: $state" -ForegroundColor Yellow
        $missingProviders += $provider
      }
    }
  }

  if ($missingProviders.Count -eq 0) {
    Write-Host "  $_chk All required resource providers are ready." -ForegroundColor Green
    return
  }

  foreach ($provider in $missingProviders) {
    if ($_whatIf) {
      Write-Host "  [WhatIf] Would register $provider." -ForegroundColor DarkGray
      continue
    }
    Write-Host "  $_arr Registering $provider..." -ForegroundColor Cyan
    try {
      Invoke-AzureCli -Arguments @('provider', 'register', '--namespace', $provider, '--wait') | Out-Null
      Write-Host "  $_chk $provider registered." -ForegroundColor Green
    }
    catch {
      Write-Failure -Lines @(
        "Could not register $provider."
        'This usually means your account lacks subscription-level register permission.'
        'Minimum built-in role: Contributor. Owner also works.'
      )
      throw "Provider registration failed for $provider."
    }
  }
}

function Invoke-AzdProvision {
  $repoRoot = Get-RepoRoot

  Write-Host ''
  Write-Host '  azd provision' -ForegroundColor Cyan
  Write-Host $_sep -ForegroundColor DarkGray
  if ($_whatIf) {
    Write-Host '  [WhatIf] Would run: azd provision' -ForegroundColor DarkGray
    Write-Host "           Working directory: $repoRoot" -ForegroundColor DarkGray
    return
  }

  # ── azd environment name ──────────────────────────────────────────────────
  # azd uses an "environment" as a named workspace that stores deployment
  # configuration in .azure/<name>/.env (subscription, location, parameters).
  # We prompt for it here with a sensible default so the user does not see an
  # unexplained prompt mid-run.
  Write-Host '  azd stores your deployment configuration in a named environment'
  Write-Host '  (a folder under .azure/ in this repo). Use the default or enter'
  Write-Host '  a short name that identifies this deployment (e.g. "contoso-gsi").'
  Write-Host ''
  do {
    $_envName = (Read-Host '  Environment name [guest-sponsor-info]').Trim()
    if ($_envName -eq '') { $_envName = 'guest-sponsor-info' }
    # azd allows letters, digits, and hyphens; must start with a letter or digit.
    if ($_envName -notmatch '^[a-zA-Z0-9][a-zA-Z0-9\-]{0,62}$') {
      Write-Host "  $_wrn Name must start with a letter or digit, contain only letters, digits," -ForegroundColor Yellow
      Write-Host '        and hyphens, and be between 1 and 64 characters.' -ForegroundColor Yellow
      $_envName = ''
    }
  } while (-not $_envName)
  Write-Host ''

  # ── Resource group ────────────────────────────────────────────────────────
  # azd creates the resource group itself (as the Bicep deployment target).
  # The default name follows the Azure CAF pattern rg-<workload>. The operator
  # can override — common when deploying multiple environments side by side.
  $_rgDefault = "rg-$_envName"
  Write-Host '  Resource Group' -ForegroundColor Cyan
  Write-Host $_sep -ForegroundColor DarkGray
  Write-Host '  The resource group that azd will create (or reuse) for this deployment.'
  Write-Host '  Azure naming best practice: rg-<workload>  or  rg-<workload>-<environment>'
  Write-Host "  Suggested: $_rgDefault" -ForegroundColor DarkGray
  Write-Host ''
  $_rgName = (Read-Host "  Resource group [$_rgDefault]").Trim()
  if ($_rgName -eq '') { $_rgName = $_rgDefault }
  Write-Host ''

  # ── Azure location ────────────────────────────────────────────────────────
  # Set AZURE_LOCATION so azd does not open an extra prompt mid-run.
  Write-Host '  Azure Location' -ForegroundColor Cyan
  Write-Host $_sep -ForegroundColor DarkGray
  Write-Host '  The Azure region where all resources will be deployed.'
  Write-Link -Url 'https://azure.microsoft.com/explore/global-infrastructure/geographies/' `
    -Text 'Azure regions overview'
  Write-Host ''
  do {
    $_location = (Read-Host '  Azure location [westeurope]').Trim()
    if ($_location -eq '') { $_location = 'westeurope' }
    # Basic sanity check: Azure location names are lowercase letters and digits only.
    if ($_location -notmatch '^[a-z][a-z0-9]+$') {
      Write-Host "  $_wrn Enter a valid Azure location name (e.g. westeurope, eastus2)." -ForegroundColor Yellow
      $_location = ''
    }
  } while (-not $_location)
  Write-Host ''

  # Pre-populate all azd environment variables so azd does not open any
  # additional interactive prompts during provision.
  #
  # IMPORTANT: setting only process env vars is NOT sufficient — azd reads
  # resource group and location from its own environment store
  # (.azure/<name>/.env), not from the calling process environment.
  # We therefore create (or select) the azd environment first, then write
  # all values via 'azd env set' so that 'azd provision' finds them and
  # skips its own prompts completely.
  Push-Location -Path $repoRoot
  try {
    $_azdEnvDir = Join-Path $repoRoot ".azure/$_envName"
    if (Test-Path $_azdEnvDir) {
      Invoke-Azd -Arguments @('env', 'select', $_envName)
    }
    else {
      Invoke-Azd -Arguments @('env', 'new', $_envName)
    }
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_SUBSCRIPTION_ID', $script:SubscriptionId)
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_RESOURCE_GROUP', $_rgName)
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_LOCATION', $_location)
  }
  finally {
    Pop-Location
  }
  # Keep process env vars too — the NEXT STEPS block reads them after provision.
  $env:AZURE_ENV_NAME = $_envName
  $env:AZURE_SUBSCRIPTION_ID = $script:SubscriptionId
  $env:AZURE_LOCATION = $_location
  $env:AZURE_RESOURCE_GROUP = $_rgName

  Write-Host "  $_arr Running azd provision." -ForegroundColor Cyan
  Write-Host '       The pre-provision hook will create the Entra App Registration (Step 1),' -ForegroundColor DarkGray
  Write-Host '       Bicep will deploy the Azure infrastructure, and the post-provision hook' -ForegroundColor DarkGray
  Write-Host '       will grant Graph permissions to the Managed Identity (Step 3).' -ForegroundColor DarkGray
  Write-Host '       All three setup steps run automatically in a single command.' -ForegroundColor DarkGray
  if ($PSScriptRoot) {
    # Only relevant when running from a local clone where the Function source
    # code is present — azd deploy / azd up require the source tree.
    Write-Host '       Use azd up later if you also want azd to handle future code redeploy cycles.' -ForegroundColor DarkGray
  }

  Push-Location -Path $repoRoot
  try {
    # --no-prompt: azd v1.24+ still shows a resource group picker even when
    # AZURE_RESOURCE_GROUP is written to the env file via 'azd env set'.
    # --no-prompt tells azd to accept the stored value as the default and
    # skip all interactive pickers.  This is safe because we pre-set every
    # required value (AZURE_SUBSCRIPTION_ID, AZURE_LOCATION, AZURE_RESOURCE_GROUP)
    # in the azd env immediately above, so no required value is missing.
    Invoke-Azd -Arguments @('provision', '--no-prompt')
  }
  finally {
    Pop-Location
  }
}

function Invoke-BicepDeployment {
  param(
    [Parameter(Mandatory)][string]$GroupName,
    [Parameter(Mandatory)][string]$TenantShortName,
    [Parameter(Mandatory)][string]$AppName,
    [Parameter(Mandatory)][string]$ClientId,
    [Parameter(Mandatory)][string]$SelectedHostingPlan,
    [Parameter(Mandatory)][bool]$DeployMaps,
    [Parameter(Mandatory)][string]$SelectedAppVersion,
    [Parameter(Mandatory)][bool]$MonitoringEnabled,
    [Parameter(Mandatory)][bool]$EnableFailureAlert,
    [Parameter(Mandatory)][int]$FlexScaleLimit
  )

  $templateFile = Get-RepoFilePath -RelativePath 'azure-function/infra/main.bicep'

  Write-Host ''
  Write-Host '  Bicep deployment' -ForegroundColor Cyan
  Write-Host $_sep -ForegroundColor DarkGray
  # Use 'what-if' instead of 'create' when running in WhatIf mode — shows
  # planned resource changes without deploying anything.
  $_subCmd = if ($_whatIf) { 'what-if' } else { 'create' }
  Invoke-AzureCli -Arguments @(
    'deployment', 'group', $_subCmd,
    '--resource-group', $GroupName,
    '--template-file', $templateFile,
    '--parameters',
    "tenantId=$($script:TenantId)",
    "tenantName=$TenantShortName",
    "functionAppName=$AppName",
    "webPartClientId=$ClientId",
    "enableMonitoring=$MonitoringEnabled",
    "deployAzureMaps=$DeployMaps",
    "hostingPlan=$SelectedHostingPlan",
    "appVersion=$SelectedAppVersion",
    "enableFailureAnomaliesAlert=$EnableFailureAlert",
    "maximumFlexInstances=$FlexScaleLimit"
  )
}

function Invoke-ArmJsonDeployment {
  param(
    [Parameter(Mandatory)][string]$GroupName,
    [Parameter(Mandatory)][string]$TenantShortName,
    [Parameter(Mandatory)][string]$AppName,
    [Parameter(Mandatory)][string]$ClientId,
    [Parameter(Mandatory)][string]$SelectedHostingPlan,
    [Parameter(Mandatory)][bool]$DeployMaps,
    [Parameter(Mandatory)][string]$SelectedAppVersion,
    [Parameter(Mandatory)][bool]$MonitoringEnabled,
    [Parameter(Mandatory)][bool]$EnableFailureAlert,
    [Parameter(Mandatory)][int]$FlexScaleLimit
  )

  $templateFile = Get-RepoFilePath -RelativePath 'azure-function/infra/azuredeploy.json'

  Write-Host ''
  Write-Host '  ARM JSON deployment' -ForegroundColor Cyan
  Write-Host $_sep -ForegroundColor DarkGray
  # Use 'what-if' instead of 'create' when running in WhatIf mode — shows
  # planned resource changes without deploying anything.
  $_subCmd = if ($_whatIf) { 'what-if' } else { 'create' }
  Invoke-AzureCli -Arguments @(
    'deployment', 'group', $_subCmd,
    '--resource-group', $GroupName,
    '--template-file', $templateFile,
    '--parameters',
    "tenantId=$($script:TenantId)",
    "tenantName=$TenantShortName",
    "functionAppName=$AppName",
    "webPartClientId=$ClientId",
    "enableMonitoring=$MonitoringEnabled",
    "deployAzureMaps=$DeployMaps",
    "hostingPlan=$SelectedHostingPlan",
    "appVersion=$SelectedAppVersion",
    "enableFailureAnomaliesAlert=$EnableFailureAlert",
    "maximumFlexInstances=$FlexScaleLimit"
  )
}

#region Main
try {
  Write-Host ''
  # Show a step indicator only when Step 1 was confirmed to have already
  # run in this session. Otherwise we do not yet know if the operator will
  # pick azd (all-in-one) or Bicep/ARM JSON (still needs step 3 after).
  $_step1Done = [bool]$Global:GsiSetup_AppRegistrationDone
  $_stepLabel = if ($_step1Done) { 'Step 2 of 3: Azure Deployment' } else { 'Azure Deployment' }
  Write-Host "  Guest Sponsor Info  $(if ($_u) { [string][char]0x00B7 } else { '|' })  $_stepLabel" -ForegroundColor DarkCyan
  Write-Host $_sep -ForegroundColor DarkGray

  # ── Deployment method ─────────────────────────────────────────────────────
  $toolState = Get-ToolState
  $defaultMode = Select-DefaultMode -ToolState $toolState
  $defaultReason = Get-DefaultModeReason -SelectedMode $defaultMode -ToolState $toolState

  $selectedMode = $Mode
  if ($selectedMode -eq 'Auto') {
    Write-Host ''
    Write-Host '  Deployment method' -ForegroundColor Cyan
    Write-Host $_sep -ForegroundColor DarkGray

    if (-not $_step1Done) {
      # ── Step 1 not yet run: only azd makes sense ───────────────────────────
      # azd provision runs all three setup steps in one go via its pre- and
      # post-provision hooks. Offering Bicep or ARM JSON here would leave the
      # operator without an App Registration and without Graph permissions.
      Write-Host '  The Entra App Registration (Step 1) has not been created yet in this session.'
      Write-Host '  azd provision is the only option here — it runs all three setup steps'
      Write-Host '  automatically: App Registration, Azure infrastructure, and Graph permissions.'
      Write-Host ''
      Write-Host "    az       : $(if ($toolState.AzureCliReady) { 'ready' } else { 'not installed' })" `
        -ForegroundColor $(if ($toolState.AzureCliReady) { 'Green' } else { 'Yellow' })
      Write-Host "    azd      : $(if ($toolState.AzdReady) { 'ready' } elseif ($toolState.AzdInstalled) { 'installed, but Azure CLI is still missing' } else { 'not installed — will be installed automatically' })" `
        -ForegroundColor $(if ($toolState.AzdReady) { 'Green' } else { 'Yellow' })
      Write-Host ''
      Write-Host '  Proceeding with azd provision.' -ForegroundColor DarkGray
      $selectedMode = 'Azd'
    }
    else {
      # ── Step 1 already done: only direct deployment paths ─────────────────
      # The App Registration was created in this session (or passed via param).
      # azd would re-run the pre-provision hook and repeat Step 1 unnecessarily,
      # so it is not offered here.
      Write-Host '  The App Registration (Step 1) was already completed. Choose a direct'
      Write-Host '  deployment path for the Azure infrastructure — Step 3 (Graph permissions)'
      Write-Host '  will follow as a separate script.'
      Write-Host ''
      Write-Host "    az       : $(if ($toolState.AzureCliReady) { 'ready' } else { 'not installed' })" `
        -ForegroundColor $(if ($toolState.AzureCliReady) { 'Green' } else { 'Yellow' })
      Write-Host "    az bicep : $(if ($toolState.BicepReady) { 'ready' } elseif ($toolState.AzureCliReady) { 'available after az bicep install' } else { 'Azure CLI missing' })" `
        -ForegroundColor $(if ($toolState.BicepReady) { 'Green' } else { 'Yellow' })
      Write-Host ''
      Write-Host "  Suggested: $defaultMode" -ForegroundColor DarkGray
      Write-Host "  Reason   : $defaultReason" -ForegroundColor DarkGray
      Write-Host ''
      Write-Host '  Options:'
      Write-Host '    [1] Bicep     (preferred direct Azure CLI path)'
      Write-Host '    [2] ARM JSON  (direct compatibility fallback)'
      Write-Host ''
      $_defaultOption = if ($defaultMode -eq 'ArmJson') { '2' } else { '1' }
      do {
        $_choice = (Read-Host "  Deployment method [default: $_defaultOption]").Trim()
        if ($_choice -eq '') { $_choice = $_defaultOption }
        if ($_choice -notin @('1', '2')) {
          Write-Host "  $_wrn Enter 1 or 2." -ForegroundColor Yellow
        }
      } while ($_choice -notin @('1', '2'))
      $selectedMode = switch ($_choice) {
        '1' { 'Bicep' }
        '2' { 'ArmJson' }
      }
    }
    Write-Host ''
    $_promptsShown = $true
  }

  # ── Initialize tooling and Azure CLI session ──────────────────────────────
  Initialize-DeploymentMode -SelectedMode $selectedMode
  # Allow the operator to confirm or switch the target subscription before any
  # resource operations begin. An incorrect subscription would deploy into the
  # wrong environment and is hard to undo.
  Select-AzureSubscription
  Write-Host ''
  Write-Host "  $_chk Active subscription : $($script:SubscriptionName) ($($script:SubscriptionId))" -ForegroundColor Green
  Write-Host "  $_chk Tenant ID           : $($script:TenantId)" -ForegroundColor Green

  if ($selectedMode -eq 'Azd') {
    Invoke-AzdProvision

    # ── NEXT STEPS (azd path) ─────────────────────────────────────────────
    # azd provision ran all three setup steps automatically via its hooks.
    # Try to resolve the Function App URL so the operator can finish the
    # web part configuration without switching to the Azure portal.
    $_azdFunctionUrl = $null
    if (-not $_whatIf -and $env:AZURE_RESOURCE_GROUP) {
      try {
        $_azdHostname = (Invoke-AzureCli -Arguments @(
            'functionapp', 'list',
            '--resource-group', $env:AZURE_RESOURCE_GROUP,
            '--query', '[0].defaultHostName',
            '-o', 'tsv'
          )).Trim()
        if ($_azdHostname) { $_azdFunctionUrl = "https://$_azdHostname" }
      }
      catch {
        # Non-fatal — the URL can be found in the Azure portal.
        Write-Verbose "Could not resolve Function App URL after azd provision: $_"
      }
    }
    $_ns = [System.Collections.Generic.List[string]]::new()
    $_ns.Add('All three setup steps completed successfully:')
    $_ns.Add('')
    $_ns.Add('  Step 1 — Entra App Registration   (pre-provision hook)')
    $_ns.Add('  Step 2 — Azure infrastructure     (azd provision / Bicep)')
    $_ns.Add('  Step 3 — Graph permissions        (post-provision hook)')
    $_ns.Add('')
    $_ns.Add('Configure the web part (SharePoint property pane → Guest Sponsor API):')
    if ($_azdFunctionUrl) {
      $_ns.Add("  Base URL               : $_azdFunctionUrl")
    }
    else {
      $_ns.Add('  Base URL               : see Function App hostname in the Azure portal')
      if ($env:AZURE_RESOURCE_GROUP) {
        $_ns.Add("  Resource group         : $($env:AZURE_RESOURCE_GROUP)")
      }
    }
    $_ns.Add('  Application (client) ID: see the pre-provision hook output above')
    Write-NextStep @($_ns)
    return
  }

  # ── Parameter collection ──────────────────────────────────────────────────
  if (-not $ResourceGroupName) {
    Write-Host ''
    Write-Host '  Required: Azure Resource Group' -ForegroundColor Cyan
    Write-Host $_sep -ForegroundColor DarkGray
    Write-Host '  The Azure resource group to deploy the Guest Sponsor Info resources into.'
    Write-Host '  The group will be created in the location you specify if it does not exist yet.'
    Write-Host '  Azure naming best practice: rg-<workload>  or  rg-<workload>-<environment>'
    Write-Host '  Suggested: rg-guest-sponsor-info' -ForegroundColor DarkGray
    Write-Host '  Where to find existing groups:'
    Write-Link -Url 'https://portal.azure.com/#view/HubsExtension/BrowseResourceGroups' `
      -Text "Azure Portal $_arr Resource groups"
    Write-Host ''
    do {
      $ResourceGroupName = (Read-Host '  Resource group name [rg-guest-sponsor-info]').Trim()
      if ($ResourceGroupName -eq '') { $ResourceGroupName = 'rg-guest-sponsor-info' }
    } while (-not $ResourceGroupName)
    Write-Host ''
    $_promptsShown = $true
  }

  if (-not $TenantName) {
    $_detectedTenantName = Get-DetectedTenantName
    Write-Host ''
    Write-Host '  Required: SharePoint Tenant Name' -ForegroundColor Cyan
    Write-Host $_sep -ForegroundColor DarkGray
    Write-Host '  The short name of your SharePoint Online tenant — the part before'
    Write-Host '  .sharepoint.com  (e.g. "contoso" for contoso.sharepoint.com).'
    if ($_detectedTenantName) {
      Write-Host "  Detected from the tenant's verified domains: $_detectedTenantName" -ForegroundColor DarkGray
      Write-Host '  Press Enter to accept.' -ForegroundColor DarkGray
    }
    Write-Host ''
    do {
      $_prompt = if ($_detectedTenantName) { "  SharePoint tenant name [$_detectedTenantName]" } else { '  SharePoint tenant name' }
      $TenantName = (Read-Host $_prompt).Trim()
      if (-not $TenantName -and $_detectedTenantName) { $TenantName = $_detectedTenantName }
      if (-not $TenantName) { Write-Host "  $_wrn Value is required." -ForegroundColor Yellow }
    } while (-not $TenantName)
    Write-Host ''
    $_promptsShown = $true
  }

  if (-not $FunctionAppName) {
    Write-Host ''
    Write-Host '  Required: Function App Name' -ForegroundColor Cyan
    Write-Host $_sep -ForegroundColor DarkGray
    Write-Host '  A globally unique name for the Azure Function App.'
    Write-Host '  The name becomes part of the default hostname: <name>.azurewebsites.net'
    Write-Host '  Allowed characters: letters, numbers, and hyphens. Max 60 characters.'
    Write-Host ''
    do {
      $FunctionAppName = (Read-Host '  Function App name').Trim()
      if (-not $FunctionAppName) {
        Write-Host "  $_wrn Value is required." -ForegroundColor Yellow
      }
    } while (-not $FunctionAppName)
    Write-Host ''
    $_promptsShown = $true
  }

  if (-not $PSBoundParameters.ContainsKey('HostingPlan')) {
    Write-Host ''
    Write-Host '  Hosting Plan' -ForegroundColor Cyan
    Write-Host $_sep -ForegroundColor DarkGray
    Write-Host '  Choose the Azure Functions hosting plan.'
    Write-Host '    Consumption       — pay-per-use, cold starts possible, no VNet'
    Write-Host '    FlexConsumption   — scale-to-zero with faster starts, VNet-ready'
    Write-Host '  Most deployments start with Consumption.'
    Write-Link -Url 'https://learn.microsoft.com/azure/azure-functions/functions-scale' `
      -Text "Azure Docs $_arr Azure Functions hosting options"
    Write-Host ''
    do {
      $HostingPlan = (Read-Host '  Hosting plan [Consumption]').Trim()
      if ($HostingPlan -eq '') { $HostingPlan = 'Consumption' }
      if ($HostingPlan -notin @('Consumption', 'FlexConsumption')) {
        Write-Host "  $_wrn Enter Consumption or FlexConsumption." -ForegroundColor Yellow
        $HostingPlan = ''
      }
    } while (-not $HostingPlan)
    Write-Host ''
    $_promptsShown = $true
  }

  if (-not $PSBoundParameters.ContainsKey('DeployAzureMaps')) {
    Write-Host ''
    Write-Host '  Azure Maps' -ForegroundColor Cyan
    Write-Host $_sep -ForegroundColor DarkGray
    Write-Host '  Deploy an Azure Maps account for rendering sponsor address maps in the web part.'
    Write-Host '  Set to false to skip — the web part shows an external map link instead.'
    Write-Host ''
    do {
      $_v = (Read-Host '  Deploy Azure Maps [true]').Trim().ToLowerInvariant()
      if ($_v -eq '') { $_v = 'true' }
      if ($_v -notin @('true', 'false')) {
        Write-Host "  $_wrn Enter true or false." -ForegroundColor Yellow
        $_v = ''
      }
    } while (-not $_v)
    $DeployAzureMaps = $_v -eq 'true'
    Write-Host ''
    $_promptsShown = $true
  }

  if (-not $PSBoundParameters.ContainsKey('AppVersion')) {
    Write-Host ''
    Write-Host '  Function Package Version' -ForegroundColor Cyan
    Write-Host $_sep -ForegroundColor DarkGray
    Write-Host '  The release tag of the Function App package to deploy.'
    Write-Host '  Use "latest" to always pull the most recent published release.'
    Write-Link -Url 'https://github.com/workoho/spfx-guest-sponsor-info/releases' `
      -Text "GitHub releases $_arr workoho/spfx-guest-sponsor-info"
    Write-Host ''
    $AppVersion = (Read-Host '  Function package version [latest]').Trim()
    if ($AppVersion -eq '') { $AppVersion = 'latest' }
    Write-Host ''
    $_promptsShown = $true
  }

  if (-not $PSBoundParameters.ContainsKey('EnableMonitoring')) {
    Write-Host ''
    Write-Host '  Monitoring Stack' -ForegroundColor Cyan
    Write-Host $_sep -ForegroundColor DarkGray
    Write-Host '  Deploy Log Analytics workspace, Application Insights, and alert resources.'
    Write-Host '  Strongly recommended for production — enables diagnostics and smart alerts.'
    Write-Host ''
    do {
      $_v = (Read-Host '  Enable monitoring [true]').Trim().ToLowerInvariant()
      if ($_v -eq '') { $_v = 'true' }
      if ($_v -notin @('true', 'false')) {
        Write-Host "  $_wrn Enter true or false." -ForegroundColor Yellow
        $_v = ''
      }
    } while (-not $_v)
    $EnableMonitoring = $_v -eq 'true'
    Write-Host ''
    $_promptsShown = $true
  }

  if ($EnableMonitoring) {
    if (-not $PSBoundParameters.ContainsKey('EnableFailureAnomaliesAlert')) {
      Write-Host ''
      Write-Host '  Failure Anomalies Alert' -ForegroundColor Cyan
      Write-Host $_sep -ForegroundColor DarkGray
      Write-Host '  Enable the Application Insights Failure Anomalies smart detector alert.'
      Write-Host '  Sends an email notification when the failure rate spikes unexpectedly.'
      Write-Host ''
      do {
        $_v = (Read-Host '  Enable Failure Anomalies alert [false]').Trim().ToLowerInvariant()
        if ($_v -eq '') { $_v = 'false' }
        if ($_v -notin @('true', 'false')) {
          Write-Host "  $_wrn Enter true or false." -ForegroundColor Yellow
          $_v = ''
        }
      } while (-not $_v)
      $EnableFailureAnomaliesAlert = $_v -eq 'true'
      Write-Host ''
      $_promptsShown = $true
    }
  }
  else {
    $EnableFailureAnomaliesAlert = $false
  }

  if ($HostingPlan -eq 'FlexConsumption' -and -not $PSBoundParameters.ContainsKey('MaximumFlexInstances')) {
    Write-Host ''
    Write-Host '  Maximum Flex Instances' -ForegroundColor Cyan
    Write-Host $_sep -ForegroundColor DarkGray
    Write-Host '  Hard scale-out cap for Flex Consumption — controls the maximum number of'
    Write-Host '  concurrent function instances allowed for this app. Default is 10.'
    Write-Host ''
    do {
      $_raw = (Read-Host '  Maximum Flex instances [10]').Trim()
      if ($_raw -eq '') { $_raw = '10' }
      if ($_raw -match '^\d+$') {
        $MaximumFlexInstances = [int]$_raw
      }
      else {
        Write-Host "  $_wrn Enter a positive integer." -ForegroundColor Yellow
        $_raw = ''
      }
    } while (-not $_raw)
    Write-Host ''
    $_promptsShown = $true
  }

  if (-not $WebPartClientId) {
    Write-Host ''
    Write-Host '  Required: Web Part Client ID' -ForegroundColor Cyan
    Write-Host $_sep -ForegroundColor DarkGray
    Write-Host '  The Application (client) ID of the Entra App Registration used for EasyAuth.'
    Write-Host '  Run setup-app-registration.ps1 first if you have not created it yet.'
    Write-Host '  Where to find it:'
    Write-Link -Url "https://entra.microsoft.com/$($script:TenantId)/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade" `
      -Text "Entra admin center $_arr App registrations"
    Write-Host "    'Guest Sponsor Info - SharePoint Web Part Auth' $_arr Application (client) ID"
    if ($Global:GsiSetup_WebPartClientId) {
      # Session cache: reuse the value set by setup-app-registration.ps1 in
      # this PowerShell session — no interactive prompt needed.
      $WebPartClientId = $Global:GsiSetup_WebPartClientId
      Write-Host "  Using cached App Registration Client ID from this session: $WebPartClientId" -ForegroundColor DarkGray
    }
    else {
      Write-Host ''
      $_answer = (Read-Host '  Run setup-app-registration.ps1 to create or reuse it now? [Y/n]').Trim()
      if ($_answer -eq '' -or $_answer -match '^[Yy]') {
        if ($_whatIf) {
          Write-Host "  [WhatIf] Would run: setup-app-registration.ps1 -TenantId $($script:TenantId)" -ForegroundColor DarkGray
          $WebPartClientId = '00000000-0000-0000-0000-000000000000'
        }
        else {
          $_setupScriptPath = Get-RepoFilePath -RelativePath 'azure-function/infra/setup-app-registration.ps1'
          & $_setupScriptPath -TenantId $script:TenantId -Confirm:$false
          # Mark Step 1 as done so the NEXT STEPS label is correct.
          $Global:GsiSetup_AppRegistrationDone = $true
          $WebPartClientId = (Invoke-AzureCli -Arguments @(
              'ad', 'app', 'list',
              '--display-name', $script:AppRegistrationDisplayName,
              '--query', '[0].appId',
              '-o', 'tsv'
            )).Trim()
          if (-not $WebPartClientId) {
            throw 'The App Registration script completed, but no client ID could be resolved afterwards.'
          }
        }
      }
      else {
        do {
          $WebPartClientId = (Read-Host '  App Registration Client ID').Trim()
          if (-not $WebPartClientId) {
            Write-Host "  $_wrn Value is required." -ForegroundColor Yellow
          }
        } while (-not $WebPartClientId)
      }
    }
    Write-Host ''
    $_promptsShown = $true
  }

  # Publish the resolved client ID into the session cache so
  # setup-graph-permissions.ps1 can skip its own prompt in Step 3.
  if ($WebPartClientId) {
    $Global:GsiSetup_WebPartClientId = $WebPartClientId
  }

  # ── Required Azure permissions ─────────────────────────────────────────────
  Write-Hint @(
    'Required Azure RBAC role:  Contributor  (on the target subscription or resource group)'
    '  Owner also works. For provider registration: Contributor or higher at subscription level.'
    ''
    'If your role is eligible (PIM): activate it, then re-run.'
    'If you do not have the role yet: request it from your Azure admin.'
  )
  Write-Link -Url 'https://portal.azure.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/azurerbac' `
    -Text 'PIM → My roles → Azure resources  (activate eligible role)'
  Write-Link -Url 'https://portal.azure.com/#view/Microsoft_Azure_AD_IAM/ActiveDirectoryMenuBlade/~/RolesAndAdministrators' `
    -Text 'Azure portal → Subscriptions → Access control (IAM)'

  # ── Confirmation summary ───────────────────────────────────────────────────
  # When all parameters were supplied on the command line (no interactive
  # prompts shown) we display a summary so the operator can verify before
  # the script commits any changes — unless -Confirm:$false or -WhatIf was passed.
  if (-not $_promptsShown -and
    $WhatIfPreference -ne [System.Management.Automation.SwitchParameter]$true -and
    $ConfirmPreference -ne 'None') {
    Write-Host ''
    Write-Host '  Planned operations' -ForegroundColor Cyan
    Write-Host $_sep -ForegroundColor DarkGray
    Write-Host "  Subscription    : $($script:SubscriptionName) ($($script:SubscriptionId))"
    Write-Host "  Resource group  : $ResourceGroupName"
    Write-Host "  SharePoint org  : $TenantName"
    Write-Host "  Function App    : $FunctionAppName"
    Write-Host "  Hosting plan    : $HostingPlan"
    Write-Host "  Azure Maps      : $DeployAzureMaps"
    Write-Host "  Monitoring      : $EnableMonitoring"
    Write-Host "  App version     : $AppVersion"
    Write-Host "  Client ID       : $WebPartClientId"
    Write-Host "  Method          : $selectedMode"
    Write-Host ''
    Write-Host '  The script will deploy Azure resources and register resource providers.' -ForegroundColor DarkGray
    Write-Host '  All deployment operations are idempotent — re-running is safe.' -ForegroundColor DarkGray
    Write-Host ''
    $reply = (Read-Host '  Proceed? [Y/n]').Trim()
    if ($reply -and $reply -notmatch '^[Yy]') {
      Write-Host 'Aborted.' -ForegroundColor Yellow
      exit 0
    }
    Write-Host ''
  }

  # ── Deployment ────────────────────────────────────────────────────────────
  Initialize-ResourceGroup -Name $ResourceGroupName
  Invoke-ProviderPreflight -SelectedHostingPlan $HostingPlan -DeployMaps $DeployAzureMaps -MonitoringEnabled:$EnableMonitoring

  if ($selectedMode -eq 'Bicep') {
    Invoke-BicepDeployment `
      -GroupName $ResourceGroupName `
      -TenantShortName $TenantName `
      -AppName $FunctionAppName `
      -ClientId $WebPartClientId `
      -SelectedHostingPlan $HostingPlan `
      -DeployMaps:$DeployAzureMaps `
      -SelectedAppVersion $AppVersion `
      -MonitoringEnabled:$EnableMonitoring `
      -EnableFailureAlert:$EnableFailureAnomaliesAlert `
      -FlexScaleLimit $MaximumFlexInstances
  }
  else {
    Invoke-ArmJsonDeployment `
      -GroupName $ResourceGroupName `
      -TenantShortName $TenantName `
      -AppName $FunctionAppName `
      -ClientId $WebPartClientId `
      -SelectedHostingPlan $HostingPlan `
      -DeployMaps:$DeployAzureMaps `
      -SelectedAppVersion $AppVersion `
      -MonitoringEnabled:$EnableMonitoring `
      -EnableFailureAlert:$EnableFailureAnomaliesAlert `
      -FlexScaleLimit $MaximumFlexInstances
  }

  # ── Publish deployment outputs to the session cache ───────────────────────
  # Read the Managed Identity Object ID directly from the deployed Function
  # App. This lets setup-graph-permissions.ps1 (Step 3) skip that prompt when
  # run in the same PowerShell session immediately after this script.
  if (-not $_whatIf) {
    try {
      $_principalId = (Invoke-AzureCli -Arguments @(
          'functionapp', 'identity', 'show',
          '--name', $FunctionAppName,
          '--resource-group', $ResourceGroupName,
          '--query', 'principalId',
          '-o', 'tsv'
        )).Trim()
      if ($_principalId) {
        $Global:GsiSetup_ManagedIdentityObjectId = $_principalId
        Write-Host "  $_chk Managed Identity Object ID cached for Step 3: $_principalId" -ForegroundColor Green
      }
    }
    catch {
      # Non-fatal — Step 3 will prompt for the value manually if it is missing.
      Write-Host "  $_wrn Could not read Managed Identity Object ID from the Function App." -ForegroundColor Yellow
      Write-Host '       You will be prompted for it in setup-graph-permissions.ps1.' -ForegroundColor DarkGray
    }
  }
  # $PSScriptRoot is empty when the script was run via iwr (scriptblock
  # execution) and non-empty when run from a saved local file — use this to
  # show the right command to the operator.
  $_graphPermScript = $null
  if ($PSScriptRoot) {
    $_candidate = Join-Path $PSScriptRoot 'setup-graph-permissions.ps1'
    if (Test-Path $_candidate) { $_graphPermScript = $_candidate }
  }
  if ($_graphPermScript) {
    # Local file found — run it directly.
    $_graphPermCmd = "& '$_graphPermScript'"
  }
  else {
    # Not available locally — provide the iwr one-liner from GitHub.
    $_graphPermCmd = "& ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/setup-graph-permissions.ps1').Content))"
  }

  # Re-read the step-1-done flag — it may have been set during this run
  # (e.g. the operator chose to run setup-app-registration.ps1 inline).
  $_step1DoneNow = [bool]$Global:GsiSetup_AppRegistrationDone
  # Determine the step label for "run setup-graph-permissions.ps1":
  #   - Step 1 done before this script → it is "Step 3" of the 3-step flow
  #   - Step 1 done inline during this run → same
  #   - Step 1 never done → this was step 1, graph permissions is step 2
  $_graphPermStepLabel = if ($_step1DoneNow) { 'Step 3' } else { 'Step 2' }
  Write-NextStep @(
    'Step 1 — Verify the deployment succeeded in the Azure portal.'
    ''
    "$_graphPermStepLabel — Run setup-graph-permissions.ps1 to assign Graph app roles"
    '         to the Function App Managed Identity and enable silent'
    '         token acquisition by the web part.'
    ''
    "  $_graphPermCmd"
    ''
    '  You will need the Managed Identity Object ID from the'
    '  deployment outputs (Azure portal → Resource group → Deployments).'
  )
}
finally {
  foreach ($path in $script:TempPaths) {
    if (Test-Path -Path $path) {
      Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
    }
  }
}
#endregion
