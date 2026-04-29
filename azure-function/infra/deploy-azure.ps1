#!/usr/bin/env -S pwsh -NoLogo -NoProfile

<#
.SYNOPSIS
    Interactive deployment wizard for Guest Sponsor Info Azure infrastructure.

.DESCRIPTION
    Guided console wizard that collects all required deployment parameters and
    deploys the Guest Sponsor Info Azure infrastructure via Azure Developer CLI
    (azd). All parameters can be provided on the command line for unattended
    operation.

    The script must be run from the azure-function/infra/ directory inside the
    repository or an extracted infra package. It uses azd exclusively and
    installs Azure CLI and Azure Developer CLI automatically when missing.

    To download the infra package and run this wizard without cloning the
    repository, use the installer wrapper:

      & ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.ps1').Content))

.PARAMETER AzdEnvironmentName
    azd environment name (stored under .azure/<name>/ in the repository root).
    Defaults to "guest-sponsor-info" when not specified.

.PARAMETER ResourceGroupName
    Azure resource group to deploy into. Created when it does not exist yet.
    Defaults to rg-<AzdEnvironmentName>.

.PARAMETER AzureLocation
    Azure region for all resources (e.g. "westeurope", "eastus2").
    Defaults to "westeurope" when not specified.

.PARAMETER AzureTenantId
    Optional Azure/Entra tenant ID used for az login. Use this when the account
    is a guest in other tenants that block tenant enumeration via Conditional
    Access.

.PARAMETER TenantName
    SharePoint tenant short name — the part before .sharepoint.com
    (e.g. "contoso" for contoso.sharepoint.com).

.PARAMETER FunctionAppName
    Globally unique Function App name (2-58 characters). Bicep auto-generates
    one (e.g. "gsi-a1b2c3d4") when left blank.

.PARAMETER HostingPlan
    Consumption (default) or FlexConsumption.

.PARAMETER DeployAzureMaps
    Deploy an Azure Maps account for address map rendering. Defaults to true.

.PARAMETER AppVersion
    Function package version tag. Defaults to "latest".

.PARAMETER Environment
    Optional workload environment tag. The wizard suggests "prod" by default.
    Enter an empty string on the command line or "-" in the wizard to omit it.

.PARAMETER Criticality
    Optional workload criticality tag. The wizard suggests "low" by default.
    Enter an empty string on the command line or "-" in the wizard to omit it.

.PARAMETER EnableMonitoring
    Deploy Log Analytics workspace, Application Insights, and alert resources.
    Defaults to true.

.PARAMETER EnableFailureAnomaliesAlert
    Enable the Application Insights Failure Anomalies smart detector alert rule.
    Defaults to false.

.PARAMETER MaximumFlexInstances
    Hard scale-out cap for Flex Consumption. Defaults to 10.

.PARAMETER AlwaysReadyInstances
  Number of always-ready (pre-warmed) instances for Flex Consumption. Defaults to 1.

.PARAMETER InstanceMemoryMB
  Memory size per Flex Consumption instance in MB. Defaults to 2048.

.PARAMETER SkipGraphRoleAssignments
    Defer Microsoft Graph app role assignments to setup-graph-permissions.ps1.
    Requires Privileged Role Administrator. Default: false (assign now).

.PARAMETER PreflightOnly
    Install/check required tools, sign in, collect deployment settings, and
    validate the visible Azure/Entra prerequisites without running azd provision.

.EXAMPLE
    ./deploy-azure.ps1

.EXAMPLE
    ./deploy-azure.ps1 -ResourceGroupName rg-gsi -TenantName contoso

.EXAMPLE
    ./deploy-azure.ps1 -SkipGraphRoleAssignments $true -AzureLocation eastus2

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
  [string]$AzdEnvironmentName,
  [string]$ResourceGroupName,
  [string]$AzureLocation,
  [string]$AzureTenantId,
  [string]$TenantName,
  [string]$FunctionAppName,
  [ValidateSet('Consumption', 'FlexConsumption')]
  [string]$HostingPlan = 'Consumption',
  [bool]$DeployAzureMaps = $true,
  [string]$AppVersion = 'latest',
  [AllowEmptyString()]
  [string]$Environment = '',
  [AllowEmptyString()]
  [string]$Criticality = '',
  [bool]$EnableMonitoring = $true,
  [bool]$EnableFailureAnomaliesAlert = $false,
  [int]$AlwaysReadyInstances = 1,
  [int]$MaximumFlexInstances = 10,
  [ValidateSet(512, 2048)]
  [int]$InstanceMemoryMB = 2048,
  [bool]$SkipGraphRoleAssignments = $false,
  [switch]$PreflightOnly
)

$ErrorActionPreference = 'Stop'

# Track whether any interactive prompt was shown. When all parameters were
# pre-supplied (via the command line or the session cache) we show a
# confirmation summary so the operator can verify before the script runs.
$_promptsShown = $false
# Convenience bool used throughout for WhatIf-aware fallbacks.
$_whatIf = $WhatIfPreference -eq [System.Management.Automation.SwitchParameter]$true

$script:AppRegistrationDisplayName = 'Guest Sponsor Info - SharePoint Web Part Auth'
$script:AzPath = $null
$script:AzdPath = $null
$script:SubscriptionName = ''
$script:SubscriptionId = ''
$script:TenantId = ''
$script:FunctionAppNameMinLength = 2
$script:FunctionAppNameMaxLength = 58
$script:DeploySessionCache = if ($Global:GsiDeploy_Cache -is [hashtable]) { $Global:GsiDeploy_Cache } else { $null }
$script:CachedDeployParameters = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$script:ExplicitDeployParameters = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$script:UsedDeploySessionCache = $false
$script:ReconfigureMode = $false
foreach ($_parameterName in $PSBoundParameters.Keys) {
  $null = $script:ExplicitDeployParameters.Add($_parameterName)
}
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

function Test-FunctionAppNameLength {
  param([string]$Value)

  return -not [string]::IsNullOrWhiteSpace($Value) -and
  $Value.Length -ge $script:FunctionAppNameMinLength -and
  $Value.Length -le $script:FunctionAppNameMaxLength
}

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

function Test-MacOSHost {
  return (Get-Variable -Name IsMacOS -ValueOnly -ErrorAction SilentlyContinue) -eq $true
}

function Test-DebianLinuxHost {
  if ((Test-WindowsHost) -or (Test-MacOSHost) -or (-not (Test-Path -Path '/etc/os-release'))) {
    return $false
  }

  $osRelease = Get-Content -Path '/etc/os-release' -ErrorAction SilentlyContinue
  return ($osRelease -match '^(ID|ID_LIKE)=.*(debian|ubuntu)').Count -gt 0
}

function Add-DirectoryToPath {
  param([Parameter(Mandatory)][string]$Path)

  if ((Test-Path -Path $Path) -and $env:PATH -notlike "*$Path*") {
    $env:PATH = "$Path`:$env:PATH"
  }
}

function Get-HomebrewPath {
  $brewCommand = Get-Command -Name brew -ErrorAction SilentlyContinue
  if ($brewCommand) {
    return $brewCommand.Source
  }

  foreach ($candidate in @('/opt/homebrew/bin/brew', '/usr/local/bin/brew')) {
    if (Test-Path -Path $candidate) {
      return $candidate
    }
  }

  return $null
}

function Update-ProcessPathFromHomebrew {
  [CmdletBinding(SupportsShouldProcess)]
  param()

  if ($PSCmdlet.ShouldProcess('process PATH', 'include Homebrew prefixes')) {
    Add-DirectoryToPath -Path '/opt/homebrew/bin'
    Add-DirectoryToPath -Path '/usr/local/bin'
  }
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

  $_previousPythonWarnings = $env:PYTHONWARNINGS
  try {
    $env:PYTHONWARNINGS = if ($_previousPythonWarnings) {
      "ignore::SyntaxWarning,$_previousPythonWarnings"
    }
    else {
      'ignore::SyntaxWarning'
    }

    & $script:AzPath @Arguments
    if ($LASTEXITCODE -ne 0) {
      throw "Azure CLI failed (exit code $LASTEXITCODE): az $($Arguments -join ' ')"
    }
  }
  finally {
    if ($null -eq $_previousPythonWarnings) {
      Remove-Item -Path Env:PYTHONWARNINGS -ErrorAction SilentlyContinue
    }
    else {
      $env:PYTHONWARNINGS = $_previousPythonWarnings
    }
  }
}

function Invoke-AzureCliQuiet {
  param([Parameter(ValueFromRemainingArguments)][string[]]$Arguments)

  $_previousPythonWarnings = $env:PYTHONWARNINGS
  try {
    $env:PYTHONWARNINGS = if ($_previousPythonWarnings) {
      "ignore::SyntaxWarning,$_previousPythonWarnings"
    }
    else {
      'ignore::SyntaxWarning'
    }

    $output = & $script:AzPath @Arguments 2>$null
    if ($LASTEXITCODE -ne 0) {
      throw "Azure CLI failed (exit code $LASTEXITCODE): az $($Arguments -join ' ')"
    }
  }
  finally {
    if ($null -eq $_previousPythonWarnings) {
      Remove-Item -Path Env:PYTHONWARNINGS -ErrorAction SilentlyContinue
    }
    else {
      $env:PYTHONWARNINGS = $_previousPythonWarnings
    }
  }

  return $output
}

function Test-AzureCliAccountAvailable {
  & $script:AzPath account show --output none 2>$null
  return $LASTEXITCODE -eq 0
}
#endregion

function Invoke-Azd {
  param([Parameter(ValueFromRemainingArguments)][string[]]$Arguments)

  & $script:AzdPath @Arguments
  if ($LASTEXITCODE -ne 0) {
    throw "Azure Developer CLI failed (exit code $LASTEXITCODE): azd $($Arguments -join ' ')"
  }
}

function Show-PreflightOverview {
  Write-Hint @(
    'Before deployment this wizard checks the local tools and signs in to Azure.'
    ''
    'It can install missing tools when needed:'
    '  PowerShell bootstrapper (install.sh): PowerShell 7+'
    '  This deployment wizard: Azure CLI (az) and Azure Developer CLI (azd)'
    '  macOS fallback: Homebrew when Azure CLI installation needs it'
    ''
    'Interactive steps you may see: tool installation prompts, sudo/admin password prompts,'
    'browser-based az login, subscription selection, and deployment parameter prompts.'
  )
}

function Get-CommandVersionText {
  param(
    [Parameter(Mandatory)][string]$Command,
    [Parameter(Mandatory)][string[]]$Arguments
  )

  try {
    $output = & $Command @Arguments 2>$null
    if ($LASTEXITCODE -ne 0) { return 'available (version check failed)' }
    $line = @($output | Where-Object { $_ } | Select-Object -First 1)[0]
    if ($line) { return $line.Trim() }
  }
  catch {
    return 'available (version check failed)'
  }

  return 'available'
}

function Show-ToolVersion {
  $pwshVersion = $PSVersionTable.PSVersion.ToString()
  $azVersion = Get-CommandVersionText -Command $script:AzPath -Arguments @('--version')
  $azdVersion = Get-CommandVersionText -Command $script:AzdPath -Arguments @('version')

  Write-Host ''
  Write-Host '  Tool preflight' -ForegroundColor Cyan
  Write-Host $_sep -ForegroundColor DarkGray
  Write-Host "  $_chk PowerShell            : $pwshVersion" -ForegroundColor Green
  Write-Host "  $_chk Azure CLI (az)       : $azVersion" -ForegroundColor Green
  Write-Host "  $_chk Azure Developer CLI  : $azdVersion" -ForegroundColor Green
  Write-Host '       azd uses its own scoped Bicep CLI during azd provision.' -ForegroundColor DarkGray
}

function Install-AzureCliIfNeeded {
  $script:AzPath = Get-AzureCliPath
  if ($script:AzPath) {
    return
  }

  Write-Host "  $_wrn Azure CLI (az) is not installed." -ForegroundColor Yellow

  if (Test-WindowsHost) {
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
  }
  elseif (Test-MacOSHost) {
    Update-ProcessPathFromHomebrew
    $brew = Get-HomebrewPath
    if (-not $brew) {
      $answer = (Read-Host '  Homebrew is required to install Azure CLI on macOS. Install Homebrew now? [Y/n]').Trim()
      if ($answer -ne '' -and $answer -notmatch '^[Yy]') {
        throw @(
          'Azure CLI is required for all supported deployment methods in this repository.',
          'Install Homebrew or Azure CLI manually and re-run this script.',
          'Homebrew: https://brew.sh',
          'Azure CLI: https://learn.microsoft.com/cli/azure/install-azure-cli-macos'
        ) -join ' '
      }

      Write-Host "  $_arr Installing Homebrew..." -ForegroundColor Cyan
      # Official Homebrew installer from https://brew.sh. Download to a temp
      # file first so network failures are explicit and cleanup is reliable.
      $homebrewInstaller = Join-Path ([System.IO.Path]::GetTempPath()) "homebrew-install-$([guid]::NewGuid().ToString('n')).sh"
      try {
        Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh' `
          -OutFile $homebrewInstaller `
          -UseBasicParsing
        & /bin/bash $homebrewInstaller
        if ($LASTEXITCODE -ne 0) { throw "Homebrew installation failed (exit $LASTEXITCODE)." }
      }
      finally {
        Remove-Item -Path $homebrewInstaller -Force -ErrorAction SilentlyContinue
      }

      Update-ProcessPathFromHomebrew
      $brew = Get-HomebrewPath
      if (-not $brew) {
        throw 'Homebrew was installed, but the current session still cannot find brew. Open a new terminal and re-run the script.'
      }
    }

    $answer = (Read-Host '  Install Azure CLI now via Homebrew? [Y/n]').Trim()
    if ($answer -ne '' -and $answer -notmatch '^[Yy]') {
      throw 'Azure CLI is required for all supported deployment methods in this repository.'
    }

    Write-Host "  $_arr Installing Azure CLI via Homebrew..." -ForegroundColor Cyan
    & $brew update
    if ($LASTEXITCODE -ne 0) { throw "brew update failed (exit $LASTEXITCODE)." }
    & $brew install azure-cli
    if ($LASTEXITCODE -ne 0) { throw "brew install azure-cli failed (exit $LASTEXITCODE)." }
  }
  else {
    $curl = Get-Command -Name curl -ErrorAction SilentlyContinue
    if (-not (Test-DebianLinuxHost)) {
      throw @(
        'Azure CLI (az) is not installed.',
        'Automatic installation is only supported on Windows, macOS, and Debian/Ubuntu Linux.',
        'Install it manually and re-run this script.',
        'Docs: https://learn.microsoft.com/cli/azure/install-azure-cli-linux'
      ) -join ' '
    }

    if (-not $curl) {
      throw @(
        'Azure CLI (az) is not installed and curl is not available.',
        'Install it manually and re-run this script.',
        'Docs: https://learn.microsoft.com/cli/azure/install-azure-cli-linux'
      ) -join ' '
    }

    $answer = (Read-Host '  Install Azure CLI now via the official Microsoft Linux install script? [Y/n]').Trim()
    if ($answer -ne '' -and $answer -notmatch '^[Yy]') {
      throw 'Azure CLI is required for all supported deployment methods in this repository.'
    }

    Write-Host "  $_arr Installing Azure CLI via Microsoft install script..." -ForegroundColor Cyan
    & bash -c 'curl -sL https://aka.ms/InstallAzureCLIDeb | sudo bash'
    if ($LASTEXITCODE -ne 0) { throw "Azure CLI installation script failed (exit $LASTEXITCODE)." }
  }

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

function Connect-AzureCliIfNeeded {
  if (-not (Test-AzureCliAccountAvailable)) {
    Write-Host "  $_arr No active Azure CLI session found. Starting az login..." -ForegroundColor Cyan

    $_loginTenantId = ''
    if ($AzureTenantId) {
      $_loginTenantId = $AzureTenantId.Trim()
    }
    elseif ($Global:GsiSetup_TenantId) {
      # GsiSetup_TenantId may have been set by setup-graph-permissions.ps1 or a
      # previous run of this script. Pass it as a hint so az login lands on the
      # right tenant without asking the operator to pick one manually.
      $_loginTenantId = ([string]$Global:GsiSetup_TenantId).Trim()
    }
    else {
      Write-Host ''
      Write-Host '  Azure/Entra Tenant' -ForegroundColor Cyan
      Write-Host $_sep -ForegroundColor DarkGray
      Write-Host '  Optional: paste your target tenant ID to sign in directly to that tenant.'
      Write-Host '  Leave blank to let Azure CLI discover all tenants for this account.'
      Write-Host '  This helps when guest tenants block enumeration via Conditional Access.'
      Write-Host ''
      $_loginTenantId = (Read-Host '  Tenant ID for az login [auto]').Trim()
      Write-Host ''
    }

    $_loginArgs = @('login')
    if ($_loginTenantId) {
      $_loginArgs += @('--tenant', $_loginTenantId)
    }

    & $script:AzPath @_loginArgs | Out-Null
    if ($LASTEXITCODE -ne 0) {
      throw "Azure CLI sign-in failed (exit code $LASTEXITCODE). Re-run with the target tenant ID if this account is a guest in other tenants."
    }

    if (-not (Test-AzureCliAccountAvailable)) {
      throw 'Azure CLI sign-in completed, but no active Azure subscription is available. Confirm your Azure role is active and re-run the script.'
    }

    Write-Host "  $_chk Azure CLI sign-in completed." -ForegroundColor Green
    Write-Host '     If Azure CLI reported tenant warnings above, they can be ignored as long as your target subscription is listed below.' -ForegroundColor DarkGray
  }

  $script:SubscriptionName = (Invoke-AzureCli -Arguments @('account', 'show', '--query', 'name', '-o', 'tsv')).Trim()
  $script:SubscriptionId = (Invoke-AzureCli -Arguments @('account', 'show', '--query', 'id', '-o', 'tsv')).Trim()
  $script:TenantId = (Invoke-AzureCli -Arguments @('account', 'show', '--query', 'tenantId', '-o', 'tsv')).Trim()

  # Publish the tenant ID into the shared session cache so downstream scripts
  # (e.g. setup-graph-permissions.ps1) can skip their own login prompts.
  $Global:GsiSetup_TenantId = $script:TenantId
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

function Use-DeployCachedValue {
  param(
    [Parameter(Mandatory)][string]$Key,
    [Parameter(Mandatory)][ref]$Target,
    [switch]$AllowEmptyString
  )

  if (-not $script:DeploySessionCache) { return }
  if ($script:ExplicitDeployParameters.Contains($Key)) { return }
  if (-not $script:DeploySessionCache.ContainsKey($Key)) { return }

  $_cachedValue = $script:DeploySessionCache[$Key]
  if ($null -eq $_cachedValue) { return }
  if (-not $AllowEmptyString -and $_cachedValue -is [string] -and $_cachedValue -eq '') { return }

  $Target.Value = $_cachedValue
  $null = $script:CachedDeployParameters.Add($Key)
  $script:UsedDeploySessionCache = $true
}

function Test-DeployParameterProvided {
  param([Parameter(Mandatory)][string]$Name)

  return $script:ExplicitDeployParameters.Contains($Name) -or $script:CachedDeployParameters.Contains($Name)
}

function Should-PromptDeployParameter {
  param([Parameter(Mandatory)][string]$Name)

  return $script:ReconfigureMode -or -not (Test-DeployParameterProvided -Name $Name)
}

function Get-PromptDefaultValue {
  param(
    [AllowEmptyString()][string]$CurrentValue,
    [Parameter(Mandatory)][string]$FallbackValue,
    [string]$EmptyDisplay = ''
  )

  if ($script:ReconfigureMode) {
    if ($CurrentValue -ne '') {
      return $CurrentValue
    }
    if ($EmptyDisplay -ne '') {
      return $EmptyDisplay
    }
  }

  return $FallbackValue
}

function Save-DeploySessionCache {
  param(
    [Parameter(Mandatory)][string]$SubscriptionName,
    [Parameter(Mandatory)][string]$SubscriptionId,
    [Parameter(Mandatory)][string]$TenantId,
    [Parameter(Mandatory)][string]$AzdEnvironmentName,
    [Parameter(Mandatory)][string]$ResourceGroupName,
    [Parameter(Mandatory)][string]$AzureLocation,
    [Parameter(Mandatory)][string]$TenantName,
    [Parameter(Mandatory)][AllowEmptyString()][string]$FunctionAppName,
    [Parameter(Mandatory)][string]$HostingPlan,
    [Parameter(Mandatory)][bool]$DeployAzureMaps,
    [Parameter(Mandatory)][string]$AppVersion,
    [Parameter(Mandatory)][AllowEmptyString()][string]$Environment,
    [Parameter(Mandatory)][AllowEmptyString()][string]$Criticality,
    [Parameter(Mandatory)][bool]$EnableMonitoring,
    [Parameter(Mandatory)][bool]$EnableFailureAnomaliesAlert,
    [Parameter(Mandatory)][int]$AlwaysReadyInstances,
    [Parameter(Mandatory)][int]$MaximumFlexInstances,
    [Parameter(Mandatory)][int]$InstanceMemoryMB,
    [Parameter(Mandatory)][bool]$SkipGraphRoleAssignments
  )

  $Global:GsiDeploy_Cache = @{
    SubscriptionName            = $SubscriptionName
    SubscriptionId              = $SubscriptionId
    TenantId                    = $TenantId
    AzdEnvironmentName          = $AzdEnvironmentName
    ResourceGroupName           = $ResourceGroupName
    AzureLocation               = $AzureLocation
    TenantName                  = $TenantName
    FunctionAppName             = $FunctionAppName
    HostingPlan                 = $HostingPlan
    DeployAzureMaps             = $DeployAzureMaps
    AppVersion                  = $AppVersion
    Environment                 = $Environment
    Criticality                 = $Criticality
    EnableMonitoring            = $EnableMonitoring
    EnableFailureAnomaliesAlert = $EnableFailureAnomaliesAlert
    AlwaysReadyInstances        = $AlwaysReadyInstances
    MaximumFlexInstances        = $MaximumFlexInstances
    InstanceMemoryMB            = $InstanceMemoryMB
    SkipGraphRoleAssignments    = $SkipGraphRoleAssignments
  }

  $script:DeploySessionCache = $Global:GsiDeploy_Cache
}

function Get-DetectedTenantName {
  try {
    $derivedName = Invoke-AzureCliQuiet -Arguments @(
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

function Confirm-SharePointTenantBelongsToAzureTenant {
  param(
    [Parameter(Mandatory)][string]$SharePointTenantName,
    [AllowEmptyString()][string]$DetectedTenantName
  )

  if (-not $DetectedTenantName) {
    Write-Host "  $_wrn Could not verify whether the SharePoint tenant belongs to the selected Azure tenant." -ForegroundColor Yellow
    Write-Host '     Make sure the Azure subscription is in the same Entra tenant as the SharePoint tenant.' -ForegroundColor DarkGray
    return
  }

  if ($SharePointTenantName -ieq $DetectedTenantName) {
    Write-Host "  $_chk SharePoint tenant matches the selected Azure/Entra tenant." -ForegroundColor Green
    return
  }

  Write-Important @(
    'The SharePoint tenant name does not match the selected Azure/Entra tenant.',
    '',
    "Selected Azure tenant ID : $($script:TenantId)",
    "Detected tenant name     : $DetectedTenantName.onmicrosoft.com",
    "Entered SharePoint host  : $SharePointTenantName.sharepoint.com",
    '',
    'The Azure Function, App Registration, Managed Identity, and SharePoint web part',
    'must belong to the same Entra tenant. Continue only if this SharePoint host is',
    'a renamed SharePoint domain in the selected Entra tenant.'
  )

  do {
    $_continue = (Read-Host '  Continue with this SharePoint tenant? [y/N]').Trim()
    if (-not $_continue) { $_continue = 'n' }
    if ($_continue -notmatch '^(?i:y|yes|n|no)$') {
      Write-Host "  $_wrn Enter Y or N." -ForegroundColor Yellow
      $_continue = ''
    }
  } while (-not $_continue)

  if ($_continue -notmatch '^(?i:y|yes)$') {
    throw "SharePoint tenant '$SharePointTenantName.sharepoint.com' does not match the selected Azure tenant. Select an Azure subscription in the SharePoint tenant or re-run with -AzureTenantId <tenant-id>."
  }
}

function Get-RepoRoot {
  # Supports two layouts:
  #
  # 1. Standalone infra package (extracted from guest-sponsor-info-infra.zip):
  #    deploy-azure.ps1 and azure.yaml are in the same directory.
  #    $PSScriptRoot itself is the "root" azd should run from.
  #
  # 2. Repository clone:
  #    deploy-azure.ps1 lives at <repo>/azure-function/infra/
  #    azure.yaml lives two levels up at <repo>/azure.yaml.
  if (-not $PSScriptRoot) {
    throw 'Cannot determine script location: $PSScriptRoot is empty. Use install.ps1 for remote invocation — deploy-azure.ps1 must be run from a local path.'
  }
  # Check layout 1: azure.yaml next to this script (standalone package).
  if (Test-Path (Join-Path -Path $PSScriptRoot -ChildPath 'azure.yaml')) {
    return $PSScriptRoot
  }
  # Check layout 2: azure.yaml two levels up (repository clone).
  $candidate = Resolve-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath '../..') -ErrorAction SilentlyContinue
  if ($candidate -and (Test-Path (Join-Path -Path $candidate.Path -ChildPath 'azure.yaml'))) {
    return $candidate.Path
  }
  throw "azure.yaml not found. Run deploy-azure.ps1 from the azure-function/infra/ directory (repo clone) or from an extracted infra package."
}

function Select-AzureSubscription {
  # List all subscriptions accessible to the signed-in account. If more than
  # one is available let the operator confirm or switch before deployment
  # proceeds — relying on whatever az happens to have set as default is not
  # acceptable when deploying production infrastructure.
  $_subs = $null
  try {
    $_rawJson = Invoke-AzureCliQuiet -Arguments @(
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

  if (-not $script:ReconfigureMode -and $script:DeploySessionCache -and $script:DeploySessionCache.ContainsKey('SubscriptionId')) {
    $_cachedSubscriptionId = [string]$script:DeploySessionCache.SubscriptionId
    if ($_cachedSubscriptionId) {
      $_cachedSub = $null
      foreach ($_subscription in $_subs) {
        if ($_subscription.id -eq $_cachedSubscriptionId) {
          $_cachedSub = $_subscription
          break
        }
      }

      if ($_cachedSub) {
        Write-Host ''
        Write-Host '  Azure Subscription' -ForegroundColor Cyan
        Write-Host $_sep -ForegroundColor DarkGray
        Write-Host '  A subscription from this PowerShell session is available:'
        Write-Host "    $($_cachedSub.name)" -ForegroundColor White
        Write-Host "    subscription: $($_cachedSub.id)" -ForegroundColor DarkGray
        Write-Host "    tenant      : $($_cachedSub.tenantId)" -ForegroundColor DarkGray
        Write-Host ''
        do {
          $_useCachedSubscription = (Read-Host '  Use this subscription? [Y/n]').Trim()
          if (-not $_useCachedSubscription) { $_useCachedSubscription = 'y' }
          if ($_useCachedSubscription -notmatch '^(?i:y|yes|n|no)$') {
            Write-Host "  $_wrn Enter Y or N." -ForegroundColor Yellow
            $_useCachedSubscription = ''
          }
        } while (-not $_useCachedSubscription)
        Write-Host ''

        if ($_useCachedSubscription -match '^(?i:y|yes)$') {
          if ($_cachedSub.id -ne $script:SubscriptionId) {
            Invoke-AzureCli -Arguments @('account', 'set', '--subscription', $_cachedSub.id) | Out-Null
          }
          $script:SubscriptionName = $_cachedSub.name
          $script:SubscriptionId = $_cachedSub.id
          $script:TenantId = $_cachedSub.tenantId
          $script:UsedDeploySessionCache = $true
          return
        }
      }
    }
  }

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
    Write-Host "         subscription: $($_s.id)" -ForegroundColor DarkGray
    Write-Host "         tenant      : $($_s.tenantId)" -ForegroundColor DarkGray
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

function Get-SignedInUserObjectId {
  try {
    return (Invoke-AzureCliQuiet -Arguments @('ad', 'signed-in-user', 'show', '--query', 'id', '-o', 'tsv')).Trim()
  }
  catch {
    return ''
  }
}

function Get-AzureRoleNamesForScope {
  param(
    [Parameter(Mandatory)][string]$Scope,
    [Parameter(Mandatory)][string]$AssigneeObjectId
  )

  try {
    $raw = Invoke-AzureCliQuiet -Arguments @(
      'role', 'assignment', 'list',
      '--scope', $Scope,
      '--assignee', $AssigneeObjectId,
      '--include-inherited',
      '--query', '[].roleDefinitionName',
      '-o', 'tsv'
    )
    return @($raw -split "`n" | Where-Object { $_ } | Select-Object -Unique)
  }
  catch {
    return @()
  }
}

function Get-EntraDirectoryRoleName {
  param([Parameter(Mandatory)][ref]$Succeeded)

  $Succeeded.Value = $false

  try {
    $raw = Invoke-AzureCliQuiet -Arguments @(
      'rest',
      '--method', 'GET',
      '--url', 'https://graph.microsoft.com/v1.0/me/transitiveMemberOf/microsoft.graph.directoryRole?$select=displayName',
      '--query', 'value[*].displayName',
      '-o', 'tsv'
    )
    $Succeeded.Value = $true
    return @($raw -split "`n" | Where-Object { $_ } | Select-Object -Unique)
  }
  catch {
    return @()
  }
}

function Show-MissingRoleRequest {
  param([AllowEmptyCollection()][string[]]$MissingItem = @())

  if ($MissingItem.Count -eq 0) { return }

  Write-Important @(
    'Some required permissions were not visible for the signed-in account.'
    ''
    'Copy/paste request for your Azure/Entra administrator:'
    ''
    'Please grant or activate the following roles for a Guest Sponsor Info deployment:'
    ($MissingItem | ForEach-Object { "  - $_" })
    ''
    "Target subscription: $($script:SubscriptionName) ($($script:SubscriptionId))"
    "Tenant ID: $($script:TenantId)"
    ''
    'After the roles are active, re-run the same install command.'
  )
}

function Test-DeploymentPrerequisite {
  param(
    [Parameter(Mandatory)][string]$ResourceGroupName,
    [Parameter(Mandatory)][bool]$SkipGraphRoleAssignments
  )

  Write-Host ''
  Write-Host '  Permission preflight' -ForegroundColor Cyan
  Write-Host $_sep -ForegroundColor DarkGray

  $script:LastPreflightMissingPermissions = $false
  $script:LastPreflightUnverifiedPermissions = $false
  $missing = [System.Collections.Generic.List[string]]::new()
  $subscriptionScope = "/subscriptions/$($script:SubscriptionId)"
  $resourceGroupScope = "$subscriptionScope/resourceGroups/$ResourceGroupName"
  $userObjectId = Get-SignedInUserObjectId

  if (-not $userObjectId) {
    Write-Host "  $_wrn Could not identify the signed-in user. Role checks are limited." -ForegroundColor Yellow
  }
  else {
    $subscriptionRoles = @(Get-AzureRoleNamesForScope -Scope $subscriptionScope -AssigneeObjectId $userObjectId)
    $resourceGroupExists = $false
    try {
      $resourceGroupExists = ((Invoke-AzureCliQuiet -Arguments @('group', 'exists', '--name', $ResourceGroupName)).Trim() -eq 'true')
    }
    catch {
      $resourceGroupExists = $false
    }
    $resourceGroupRoles = if ($resourceGroupExists) {
      @(Get-AzureRoleNamesForScope -Scope $resourceGroupScope -AssigneeObjectId $userObjectId)
    }
    else {
      @()
    }
    $allAzureRoles = @($subscriptionRoles + $resourceGroupRoles | Select-Object -Unique)
    $hasContributor = ($allAzureRoles | Where-Object { $_ -in @('Owner', 'Contributor') }).Count -gt 0
    $hasRoleAssignment = ($allAzureRoles | Where-Object { $_ -in @('Owner', 'User Access Administrator') }).Count -gt 0

    if ($hasContributor) {
      Write-Host "  $_chk Azure deployment role: Contributor/Owner visible." -ForegroundColor Green
    }
    else {
      Write-Host "  $_wrn Azure deployment role missing: Contributor or Owner." -ForegroundColor Yellow
      $missing.Add("Azure Contributor or Owner on $resourceGroupScope (or inherited from the subscription)")
    }

    if ($hasRoleAssignment) {
      Write-Host "  $_chk Azure role assignment permission: Owner/User Access Administrator visible." -ForegroundColor Green
    }
    else {
      Write-Host "  $_wrn Azure role assignment permission missing: Owner or User Access Administrator." -ForegroundColor Yellow
      $missing.Add("Azure Owner or User Access Administrator on $resourceGroupScope (or inherited from the subscription)")
    }

    if (-not $resourceGroupExists) {
      Write-Host '       Target resource group does not exist yet; inherited subscription roles were checked.' -ForegroundColor DarkGray
    }
  }

  $entraRoleCheckSucceeded = $false
  $entraRoles = @(Get-EntraDirectoryRoleName -Succeeded ([ref]$entraRoleCheckSucceeded))
  if (-not $entraRoleCheckSucceeded) {
    Write-Host "  $_wrn Could not verify active Entra directory roles for the signed-in account." -ForegroundColor Yellow
    Write-Host '       The deployment may still continue, but Bicep will fail if the Entra roles are not active.' -ForegroundColor DarkGray
    $script:LastPreflightUnverifiedPermissions = $true
  }
  else {
    $hasAppAdmin = ($entraRoles | Where-Object {
        $_ -in @('Global Administrator', 'Cloud Application Administrator', 'Application Administrator')
      }).Count -gt 0
    $hasPrivilegedRoleAdmin = ($entraRoles | Where-Object {
        $_ -in @('Global Administrator', 'Privileged Role Administrator')
      }).Count -gt 0

    if ($hasAppAdmin) {
      Write-Host "  $_chk Entra app registration role: available." -ForegroundColor Green
    }
    else {
      Write-Host "  $_wrn Entra app registration role missing: Cloud Application Administrator." -ForegroundColor Yellow
      $missing.Add('Entra Cloud Application Administrator or Application Administrator (Global Administrator also works)')
    }

    if ($SkipGraphRoleAssignments) {
      Write-Host '  Graph app-role assignment: deferred to setup-graph-permissions.ps1.' -ForegroundColor DarkGray
    }
    elseif ($hasPrivilegedRoleAdmin) {
      Write-Host "  $_chk Entra Graph app-role assignment role: available." -ForegroundColor Green
    }
    else {
      Write-Host "  $_wrn Entra Graph app-role assignment role missing: Privileged Role Administrator." -ForegroundColor Yellow
      $missing.Add('Entra Privileged Role Administrator for Microsoft Graph app-role assignments (Global Administrator also works)')
    }
  }

  $script:LastPreflightMissingPermissions = $missing.Count -gt 0
  Show-MissingRoleRequest -MissingItem @($missing)

  return $missing.Count -eq 0 -and -not $script:LastPreflightUnverifiedPermissions
}

function Invoke-AzdProvision {
  param(
    [Parameter(Mandatory)][string]$EnvName,
    [Parameter(Mandatory)][string]$ResourceGroup,
    [Parameter(Mandatory)][string]$Location,
    [Parameter(Mandatory)][string]$SharePointTenant,
    [Parameter(Mandatory)][AllowEmptyString()][string]$Environment,
    [Parameter(Mandatory)][AllowEmptyString()][string]$Criticality,
    [string]$AppName,
    [Parameter(Mandatory)][string]$Plan,
    [Parameter(Mandatory)][bool]$Maps,
    [Parameter(Mandatory)][string]$Version,
    [Parameter(Mandatory)][bool]$Monitoring,
    [Parameter(Mandatory)][bool]$FailureAlert,
    [Parameter(Mandatory)][int]$AlwaysReadyInstances,
    [Parameter(Mandatory)][int]$FlexInstances,
    [Parameter(Mandatory)][int]$InstanceMemoryMB,
    [Parameter(Mandatory)][bool]$SkipRoles
  )

  $repoRoot = Get-RepoRoot

  Write-Host ''
  Write-Host '  azd provision' -ForegroundColor Cyan
  Write-Host $_sep -ForegroundColor DarkGray

  # Tell azd to reuse the Azure CLI token so the user is not prompted to
  # log in a second time via a separate azd browser window.
  Invoke-Azd -Arguments @('config', 'set', 'auth.useAzureCliCredentials', 'true')

  # Create or select the azd environment, then pre-populate all required env
  # vars so azd does not open any additional interactive prompts during provision.
  #
  # IMPORTANT: setting only process env vars is NOT sufficient — azd reads
  # resource group and location from its own environment store
  # (.azure/<name>/.env), not from the calling process environment.
  Push-Location -Path $repoRoot
  try {
    $_azdEnvDir = Join-Path $repoRoot ".azure/$EnvName"
    if (Test-Path $_azdEnvDir) {
      Invoke-Azd -Arguments @('env', 'select', $EnvName)
    }
    else {
      Invoke-Azd -Arguments @('env', 'new', $EnvName)
    }
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_SUBSCRIPTION_ID', $script:SubscriptionId)
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_RESOURCE_GROUP', $ResourceGroup)
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_LOCATION', $Location)
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_SHAREPOINT_TENANT_NAME', $SharePointTenant)
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_TAG_ENVIRONMENT', $Environment)
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_TAG_CRITICALITY', $Criticality)
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_APP_VERSION', $Version)
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_ENABLE_MONITORING', $Monitoring.ToString().ToLower())
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_ENABLE_FAILURE_ANOMALIES_ALERT', $FailureAlert.ToString().ToLower())
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_ALWAYS_READY_INSTANCES', $AlwaysReadyInstances.ToString())
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_MAXIMUM_FLEX_INSTANCES', $FlexInstances.ToString())
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_INSTANCE_MEMORY_MB', $InstanceMemoryMB.ToString())
    # Store deployment params in azd env so the pre-provision hook can read
    # them for provider preflight (hosting plan → ContainerInstance, maps → Maps).
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_HOSTING_PLAN', $Plan)
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_DEPLOY_AZURE_MAPS', $Maps.ToString().ToLower())
    # Store the graph role assignment preference so the post-provision hook
    # can give the correct next-steps guidance.
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_SKIP_GRAPH_ROLE_ASSIGNMENTS', $SkipRoles.ToString().ToLower())
  }
  finally {
    Pop-Location
  }

  # Keep process env vars too — the post-provision hook and NEXT STEPS block read them.
  $env:AZURE_ENV_NAME = $EnvName
  $env:AZURE_SUBSCRIPTION_ID = $script:SubscriptionId
  $env:AZURE_LOCATION = $Location
  $env:AZURE_RESOURCE_GROUP = $ResourceGroup
  $env:AZURE_SHAREPOINT_TENANT_NAME = $SharePointTenant
  $env:AZURE_TAG_ENVIRONMENT = $Environment
  $env:AZURE_TAG_CRITICALITY = $Criticality
  $env:AZURE_APP_VERSION = $Version
  $env:AZURE_ENABLE_MONITORING = $Monitoring.ToString().ToLower()
  $env:AZURE_ENABLE_FAILURE_ANOMALIES_ALERT = $FailureAlert.ToString().ToLower()
  $env:AZURE_ALWAYS_READY_INSTANCES = $AlwaysReadyInstances.ToString()
  $env:AZURE_MAXIMUM_FLEX_INSTANCES = $FlexInstances.ToString()
  $env:AZURE_INSTANCE_MEMORY_MB = $InstanceMemoryMB.ToString()
  $env:AZURE_SKIP_GRAPH_ROLE_ASSIGNMENTS = $SkipRoles.ToString().ToLower()

  if ($_whatIf) {
    Write-Host "  $_arr Running azd provision --preview..." -ForegroundColor Cyan
    Write-Host '       azd asks ARM/Bicep to preview Azure resource changes without applying them.' -ForegroundColor DarkGray
    Write-Host '       The local azd environment values were updated so the preview uses' -ForegroundColor DarkGray
    Write-Host '       the same settings as a real deployment.' -ForegroundColor DarkGray
  }
  else {
    Write-Host "  $_arr Running azd provision..." -ForegroundColor Cyan
    Write-Host '       Bicep deploys the Azure infrastructure, creates the Entra App Registration,' -ForegroundColor DarkGray
    Write-Host '       configures EasyAuth, and (unless deferred) assigns Graph permissions' -ForegroundColor DarkGray
    Write-Host '       to the Managed Identity. The post-provision hook restarts the Function App.' -ForegroundColor DarkGray
  }
  Write-Host ''

  Push-Location -Path $repoRoot
  try {
    # --no-prompt: azd v1.24+ still shows a resource group picker even when
    # AZURE_RESOURCE_GROUP is written to the env file via 'azd env set'.
    # --no-prompt tells azd to accept the stored values and skip all pickers.
    if ($_whatIf) {
      Invoke-Azd -Arguments @('provision', '--preview', '--no-prompt')
    }
    else {
      Invoke-Azd -Arguments @('provision', '--no-prompt')
    }
  }
  finally {
    Pop-Location
  }
}

#region Main
try {
  Write-Host ''
  Write-Host "  Guest Sponsor Info  $(if ($_u) { [string][char]0x00B7 } else { '|' })  Azure Deployment" -ForegroundColor DarkCyan
  Write-Host $_sep -ForegroundColor DarkGray
  Show-PreflightOverview

  # ── Install tools and connect to Azure ────────────────────────────────────
  Install-AzureCliIfNeeded
  Install-AzdIfNeeded
  Show-ToolVersion
  Connect-AzureCliIfNeeded
  # Allow the operator to confirm or switch the target subscription before any
  # resource operations begin.
  Select-AzureSubscription
  Write-Host ''
  Write-Host "  $_chk Active subscription : $($script:SubscriptionName) ($($script:SubscriptionId))" -ForegroundColor Green
  Write-Host "  $_chk Tenant ID           : $($script:TenantId)" -ForegroundColor Green

  Use-DeployCachedValue -Key 'AzdEnvironmentName' -Target ([ref]$AzdEnvironmentName)
  Use-DeployCachedValue -Key 'ResourceGroupName' -Target ([ref]$ResourceGroupName)
  Use-DeployCachedValue -Key 'AzureLocation' -Target ([ref]$AzureLocation)
  Use-DeployCachedValue -Key 'TenantName' -Target ([ref]$TenantName)
  Use-DeployCachedValue -Key 'FunctionAppName' -Target ([ref]$FunctionAppName) -AllowEmptyString
  Use-DeployCachedValue -Key 'HostingPlan' -Target ([ref]$HostingPlan)
  Use-DeployCachedValue -Key 'DeployAzureMaps' -Target ([ref]$DeployAzureMaps)
  Use-DeployCachedValue -Key 'AppVersion' -Target ([ref]$AppVersion)
  Use-DeployCachedValue -Key 'Environment' -Target ([ref]$Environment) -AllowEmptyString
  Use-DeployCachedValue -Key 'Criticality' -Target ([ref]$Criticality) -AllowEmptyString
  Use-DeployCachedValue -Key 'EnableMonitoring' -Target ([ref]$EnableMonitoring)
  Use-DeployCachedValue -Key 'EnableFailureAnomaliesAlert' -Target ([ref]$EnableFailureAnomaliesAlert)
  Use-DeployCachedValue -Key 'AlwaysReadyInstances' -Target ([ref]$AlwaysReadyInstances)
  Use-DeployCachedValue -Key 'MaximumFlexInstances' -Target ([ref]$MaximumFlexInstances)
  Use-DeployCachedValue -Key 'InstanceMemoryMB' -Target ([ref]$InstanceMemoryMB)
  Use-DeployCachedValue -Key 'SkipGraphRoleAssignments' -Target ([ref]$SkipGraphRoleAssignments)

  if ($script:UsedDeploySessionCache) {
    Write-Host "  Using cached deployment settings from this PowerShell session." -ForegroundColor DarkGray
  }

  while ($true) {
    $_promptsShown = $false

    if ($script:ReconfigureMode) {
      Select-AzureSubscription
      Write-Host ''
      Write-Host "  $_chk Active subscription : $($script:SubscriptionName) ($($script:SubscriptionId))" -ForegroundColor Green
      Write-Host "  $_chk Tenant ID           : $($script:TenantId)" -ForegroundColor Green
    }

    # ── azd environment name ──────────────────────────────────────────────────
    if ($script:ReconfigureMode -or -not $AzdEnvironmentName) {
      $_azdEnvironmentDefault = Get-PromptDefaultValue -CurrentValue $AzdEnvironmentName -FallbackValue 'guest-sponsor-info'
      Write-Host ''
      Write-Host '  azd Environment Name' -ForegroundColor Cyan
      Write-Host $_sep -ForegroundColor DarkGray
      Write-Host '  azd stores your deployment configuration in a named environment'
      Write-Host '  (a folder under .azure/ in the repository root). Use the default or enter'
      Write-Host '  a short name that identifies this deployment (e.g. "contoso-gsi").'
      Write-Host ''
      do {
        $AzdEnvironmentName = (Read-Host "  Environment name [$_azdEnvironmentDefault]").Trim()
        if ($AzdEnvironmentName -eq '') { $AzdEnvironmentName = $_azdEnvironmentDefault }
        # azd allows letters, digits, and hyphens; must start with a letter or digit.
        if ($AzdEnvironmentName -notmatch '^[a-zA-Z0-9][a-zA-Z0-9\-]{0,62}$') {
          Write-Host "  $_wrn Name must start with a letter or digit, contain only letters, digits," -ForegroundColor Yellow
          Write-Host '        and hyphens, and be between 1 and 64 characters.' -ForegroundColor Yellow
          $AzdEnvironmentName = ''
        }
      } while (-not $AzdEnvironmentName)
      Write-Host ''
      $_promptsShown = $true
    }

    # ── Resource Group ────────────────────────────────────────────────────────
    if ($script:ReconfigureMode -or -not $ResourceGroupName) {
      $_rgDefault = if ($ResourceGroupName) { $ResourceGroupName } else { "rg-$AzdEnvironmentName" }
      Write-Host ''
      Write-Host '  Resource Group' -ForegroundColor Cyan
      Write-Host $_sep -ForegroundColor DarkGray
      Write-Host '  The Azure resource group that azd will create (or reuse) for this deployment.'
      Write-Host '  Azure naming best practice: rg-<workload>  or  rg-<workload>-<environment>'
      Write-Host "  Suggested: $_rgDefault" -ForegroundColor DarkGray
      Write-Host ''
      $ResourceGroupName = (Read-Host "  Resource group [$_rgDefault]").Trim()
      if ($ResourceGroupName -eq '') { $ResourceGroupName = $_rgDefault }
      Write-Host ''
      $_promptsShown = $true
    }

    # ── Azure Location ────────────────────────────────────────────────────────
    if ($script:ReconfigureMode -or -not $AzureLocation) {
      $_azureLocationDefault = Get-PromptDefaultValue -CurrentValue $AzureLocation -FallbackValue 'westeurope'
      Write-Host ''
      Write-Host '  Azure Location' -ForegroundColor Cyan
      Write-Host $_sep -ForegroundColor DarkGray
      Write-Host '  The Azure region where all resources will be deployed.'
      Write-Link -Url 'https://azure.microsoft.com/explore/global-infrastructure/geographies/' `
        -Text 'Azure regions overview'
      Write-Host ''
      do {
        $AzureLocation = (Read-Host "  Azure location [$_azureLocationDefault]").Trim()
        if ($AzureLocation -eq '') { $AzureLocation = $_azureLocationDefault }
        # Basic sanity check: Azure location names are lowercase letters and digits only.
        if ($AzureLocation -notmatch '^[a-z][a-z0-9]+$') {
          Write-Host "  $_wrn Enter a valid Azure location name (e.g. westeurope, eastus2)." -ForegroundColor Yellow
          $AzureLocation = ''
        }
      } while (-not $AzureLocation)
      Write-Host ''
      $_promptsShown = $true
    }

    # ── Environment Tag ───────────────────────────────────────────────────────
    if (Should-PromptDeployParameter -Name 'Environment') {
      $_environmentDefault = Get-PromptDefaultValue -CurrentValue $Environment -FallbackValue 'prod' -EmptyDisplay '-'
      Write-Host ''
      Write-Host '  Environment Tag' -ForegroundColor Cyan
      Write-Host $_sep -ForegroundColor DarkGray
      Write-Host '  Optional tag for the workload environment on the resource group and all resources.'
      Write-Host '  Press Enter to use the recommended default "prod".'
      Write-Host '  Enter - to omit the tag entirely, or enter any custom value.'
      Write-Host ''
      $_environment = (Read-Host "  Environment [$_environmentDefault]").Trim()
      if ($_environment -eq '') {
        $_environment = $_environmentDefault
      }
      if ($_environment -match '^(?:-|none)$') {
        $Environment = ''
      }
      else {
        $Environment = $_environment
      }
      Write-Host ''
      $_promptsShown = $true
    }

    # ── Criticality Tag ───────────────────────────────────────────────────────
    if (Should-PromptDeployParameter -Name 'Criticality') {
      $_criticalityDefault = Get-PromptDefaultValue -CurrentValue $Criticality -FallbackValue 'low' -EmptyDisplay '-'
      Write-Host ''
      Write-Host '  Criticality Tag' -ForegroundColor Cyan
      Write-Host $_sep -ForegroundColor DarkGray
      Write-Host '  Optional tag for the business criticality of this workload.'
      Write-Host '  Press Enter to use the recommended default "low".'
      Write-Host '  Enter - to omit the tag entirely, or enter any custom value.'
      Write-Host ''
      $_criticality = (Read-Host "  Criticality [$_criticalityDefault]").Trim()
      if ($_criticality -eq '') {
        $_criticality = $_criticalityDefault
      }
      if ($_criticality -match '^(?:-|none)$') {
        $Criticality = ''
      }
      else {
        $Criticality = $_criticality
      }
      Write-Host ''
      $_promptsShown = $true
    }

    # ── SharePoint Tenant Name ────────────────────────────────────────────────
    $_detectedTenantName = Get-DetectedTenantName
    if ($script:ReconfigureMode -or -not $TenantName) {
      $_tenantNameDefault = if ($TenantName) { $TenantName } elseif ($_detectedTenantName) { $_detectedTenantName } else { '' }
      Write-Host ''
      Write-Host '  SharePoint Tenant Name' -ForegroundColor Cyan
      Write-Host $_sep -ForegroundColor DarkGray
      Write-Host '  The short name of your SharePoint Online tenant — the part before'
      Write-Host '  .sharepoint.com  (e.g. "contoso" for contoso.sharepoint.com).'
      if ($_detectedTenantName) {
        Write-Host "  Detected from the tenant's verified domains: $_detectedTenantName" -ForegroundColor DarkGray
        Write-Host '  Press Enter to accept.' -ForegroundColor DarkGray
      }
      Write-Host ''
      do {
        $_prompt = if ($_tenantNameDefault) { "  SharePoint tenant name [$_tenantNameDefault]" } else { '  SharePoint tenant name' }
        $TenantName = (Read-Host $_prompt).Trim()
        if (-not $TenantName -and $_tenantNameDefault) { $TenantName = $_tenantNameDefault }
        if (-not $TenantName) { Write-Host "  $_wrn Value is required." -ForegroundColor Yellow }
      } while (-not $TenantName)
      Write-Host ''
      $_promptsShown = $true
    }
    Confirm-SharePointTenantBelongsToAzureTenant `
      -SharePointTenantName $TenantName `
      -DetectedTenantName $_detectedTenantName

    # ── Function App Name ─────────────────────────────────────────────────────
    if ($FunctionAppName -and -not (Test-FunctionAppNameLength -Value $FunctionAppName)) {
      throw "FunctionAppName must be between $($script:FunctionAppNameMinLength) and $($script:FunctionAppNameMaxLength) characters so Bicep can derive valid resource names."
    }

    if (Should-PromptDeployParameter -Name 'FunctionAppName') {
      $_functionAppNameDefault = Get-PromptDefaultValue -CurrentValue $FunctionAppName -FallbackValue 'auto-generate' -EmptyDisplay 'auto-generate'
      Write-Host ''
      Write-Host '  Function App Name  (optional)' -ForegroundColor Cyan
      Write-Host $_sep -ForegroundColor DarkGray
      Write-Host '  Globally unique name for the Azure Function App (2-58 chars).'
      Write-Host '  Leave blank to let Bicep auto-generate one (e.g., "gsi-a1b2c3d4").'
      Write-Host ''
      $_fnName = (Read-Host "  Function App Name [$_functionAppNameDefault]").Trim()
      if (-not $_fnName) { $_fnName = $_functionAppNameDefault }
      if ($_fnName -and $_fnName -ne 'auto-generate') {
        if (-not (Test-FunctionAppNameLength -Value $_fnName)) {
          Write-Host "  $_wrn Enter 2-58 characters so derived Bicep resource names stay valid." -ForegroundColor Yellow
          $_fnName = ''
        }
        else {
          $FunctionAppName = $_fnName
        }
      }
      elseif ($_fnName -eq 'auto-generate') {
        $FunctionAppName = ''
      }
      if (-not $FunctionAppName) {
        Write-Host '  Auto-generation enabled — Bicep will generate a short unique name.' -ForegroundColor DarkGray
      }
      Write-Host ''
      $_promptsShown = $true
    }

    # ── Hosting Plan ──────────────────────────────────────────────────────────
    if (Should-PromptDeployParameter -Name 'HostingPlan') {
      $_hostingPlanDefault = Get-PromptDefaultValue -CurrentValue $HostingPlan -FallbackValue 'Consumption'
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
        $_hostingPlan = (Read-Host "  Hosting plan [$_hostingPlanDefault]").Trim()
        if ($_hostingPlan -eq '') { $_hostingPlan = $_hostingPlanDefault }
        if ($_hostingPlan -notin @('Consumption', 'FlexConsumption')) {
          Write-Host "  $_wrn Enter Consumption or FlexConsumption." -ForegroundColor Yellow
          $_hostingPlan = ''
        }
        else {
          $HostingPlan = $_hostingPlan
        }
      } while (-not $HostingPlan)
      Write-Host ''
      $_promptsShown = $true
    }

    # ── Azure Maps ────────────────────────────────────────────────────────────
    if (Should-PromptDeployParameter -Name 'DeployAzureMaps') {
      $_deployAzureMapsDefault = if ($DeployAzureMaps) { 'true' } else { 'false' }
      Write-Host ''
      Write-Host '  Azure Maps' -ForegroundColor Cyan
      Write-Host $_sep -ForegroundColor DarkGray
      Write-Host '  Deploy an Azure Maps account for rendering sponsor address maps in the web part.'
      Write-Host '  Set to false to skip — the web part shows an external map link instead.'
      Write-Host ''
      do {
        $_v = (Read-Host "  Deploy Azure Maps [$_deployAzureMapsDefault]").Trim().ToLowerInvariant()
        if ($_v -eq '') { $_v = $_deployAzureMapsDefault }
        if ($_v -notin @('true', 'false')) {
          Write-Host "  $_wrn Enter true or false." -ForegroundColor Yellow
          $_v = ''
        }
      } while (-not $_v)
      $DeployAzureMaps = $_v -eq 'true'
      Write-Host ''
      $_promptsShown = $true
    }

    # ── App Version ───────────────────────────────────────────────────────────
    if (Should-PromptDeployParameter -Name 'AppVersion') {
      $_appVersionDefault = Get-PromptDefaultValue -CurrentValue $AppVersion -FallbackValue 'latest'
      Write-Host ''
      Write-Host '  Function Package Version' -ForegroundColor Cyan
      Write-Host $_sep -ForegroundColor DarkGray
      Write-Host '  The release tag of the Function App package to deploy.'
      Write-Host '  Use "latest" to always pull the most recent published release.'
      Write-Link -Url 'https://github.com/workoho/spfx-guest-sponsor-info/releases' `
        -Text "GitHub releases $_arr workoho/spfx-guest-sponsor-info"
      Write-Host ''
      $AppVersion = (Read-Host "  Function package version [$_appVersionDefault]").Trim()
      if ($AppVersion -eq '') { $AppVersion = $_appVersionDefault }
      Write-Host ''
      $_promptsShown = $true
    }

    # ── Monitoring Stack ──────────────────────────────────────────────────────
    if (Should-PromptDeployParameter -Name 'EnableMonitoring') {
      $_enableMonitoringDefault = if ($EnableMonitoring) { 'true' } else { 'false' }
      Write-Host ''
      Write-Host '  Monitoring Stack' -ForegroundColor Cyan
      Write-Host $_sep -ForegroundColor DarkGray
      Write-Host '  Deploy Log Analytics workspace, Application Insights, and alert resources.'
      Write-Host '  Strongly recommended for production — enables diagnostics and smart alerts.'
      Write-Host ''
      do {
        $_v = (Read-Host "  Enable monitoring [$_enableMonitoringDefault]").Trim().ToLowerInvariant()
        if ($_v -eq '') { $_v = $_enableMonitoringDefault }
        if ($_v -notin @('true', 'false')) {
          Write-Host "  $_wrn Enter true or false." -ForegroundColor Yellow
          $_v = ''
        }
      } while (-not $_v)
      $EnableMonitoring = $_v -eq 'true'
      Write-Host ''
      $_promptsShown = $true
    }

    # ── Failure Anomalies Alert ───────────────────────────────────────────────
    if ($EnableMonitoring -and (Should-PromptDeployParameter -Name 'EnableFailureAnomaliesAlert')) {
      $_failureAnomaliesDefault = if ($EnableFailureAnomaliesAlert) { 'true' } else { 'false' }
      Write-Host ''
      Write-Host '  Failure Anomalies Alert' -ForegroundColor Cyan
      Write-Host $_sep -ForegroundColor DarkGray
      Write-Host '  Enable the Application Insights Failure Anomalies smart detector alert.'
      Write-Host '  Sends an email notification when the failure rate spikes unexpectedly.'
      Write-Host ''
      do {
        $_v = (Read-Host "  Enable Failure Anomalies alert [$_failureAnomaliesDefault]").Trim().ToLowerInvariant()
        if ($_v -eq '') { $_v = $_failureAnomaliesDefault }
        if ($_v -notin @('true', 'false')) {
          Write-Host "  $_wrn Enter true or false." -ForegroundColor Yellow
          $_v = ''
        }
      } while (-not $_v)
      $EnableFailureAnomaliesAlert = $_v -eq 'true'
      Write-Host ''
      $_promptsShown = $true
    }
    elseif (-not $EnableMonitoring) {
      $EnableFailureAnomaliesAlert = $false
    }

    # ── Always-Ready Instances ────────────────────────────────────────────────
    if ($HostingPlan -eq 'FlexConsumption' -and (Should-PromptDeployParameter -Name 'AlwaysReadyInstances')) {
      $_alwaysReadyDefault = $AlwaysReadyInstances.ToString()
      Write-Host ''
      Write-Host '  Always-Ready Instances' -ForegroundColor Cyan
      Write-Host $_sep -ForegroundColor DarkGray
      Write-Host '  Number of pre-warmed instances kept ready for Flex Consumption.'
      Write-Host '  0 = fully on-demand (cold starts possible), 1 = warm default for most deployments.'
      Write-Host ''
      do {
        $_raw = (Read-Host "  Always-ready instances [$_alwaysReadyDefault]").Trim()
        if ($_raw -eq '') { $_raw = $_alwaysReadyDefault }
        if ($_raw -match '^[0-9]+$') {
          $AlwaysReadyInstances = [int]$_raw
        }
        else {
          Write-Host "  $_wrn Enter 0 or a positive integer." -ForegroundColor Yellow
          $_raw = ''
        }
      } while (-not $_raw)
      Write-Host ''
      $_promptsShown = $true
    }

    # ── Maximum Flex Instances ────────────────────────────────────────────────
    if ($HostingPlan -eq 'FlexConsumption' -and (Should-PromptDeployParameter -Name 'MaximumFlexInstances')) {
      $_maximumFlexInstancesDefault = $MaximumFlexInstances.ToString()
      Write-Host ''
      Write-Host '  Maximum Flex Instances' -ForegroundColor Cyan
      Write-Host $_sep -ForegroundColor DarkGray
      Write-Host '  Hard scale-out cap for Flex Consumption — controls the maximum number of'
      Write-Host '  concurrent function instances allowed for this app. Default is 10.'
      Write-Host ''
      do {
        $_raw = (Read-Host "  Maximum Flex instances [$_maximumFlexInstancesDefault]").Trim()
        if ($_raw -eq '') { $_raw = $_maximumFlexInstancesDefault }
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

    # ── Flex Instance Memory ──────────────────────────────────────────────────
    if ($HostingPlan -eq 'FlexConsumption' -and (Should-PromptDeployParameter -Name 'InstanceMemoryMB')) {
      $_instanceMemoryDefault = $InstanceMemoryMB.ToString()
      Write-Host ''
      Write-Host '  Flex Instance Memory' -ForegroundColor Cyan
      Write-Host $_sep -ForegroundColor DarkGray
      Write-Host '  Memory size per Flex Consumption instance.'
      Write-Host '  Supported in this template: 512 or 2048 MB. Recommended default: 2048 MB.'
      Write-Host ''
      do {
        $_raw = (Read-Host "  Instance memory in MB [$_instanceMemoryDefault]").Trim()
        if ($_raw -eq '') { $_raw = $_instanceMemoryDefault }
        if ($_raw -in @('512', '2048')) {
          $InstanceMemoryMB = [int]$_raw
        }
        else {
          Write-Host "  $_wrn Enter 512 or 2048." -ForegroundColor Yellow
          $_raw = ''
        }
      } while (-not $_raw)
      Write-Host ''
      $_promptsShown = $true
    }

    # ── Graph Permission Assignment ───────────────────────────────────────────
    if (Should-PromptDeployParameter -Name 'SkipGraphRoleAssignments') {
      $_graphPermissionsDefault = if ($SkipGraphRoleAssignments) { '2' } else { '1' }
      Write-Host ''
      Write-Host '  Graph Permission Assignment' -ForegroundColor Cyan
      Write-Host $_sep -ForegroundColor DarkGray
      Write-Host '  Bicep assigns Microsoft Graph app roles to the Managed Identity during'
      Write-Host '  deployment. This requires Privileged Role Administrator in Entra ID.'
      Write-Host ''
      Write-Host '    [1]  Assign now (default) — requires Privileged Role Administrator'
      Write-Host '    [2]  Defer — run setup-graph-permissions.ps1 after deployment'
      Write-Host '         (useful when a separate PAW or account holds that Entra role)'
      Write-Host ''
      do {
        $_choice = (Read-Host "  Graph permissions [$_graphPermissionsDefault]").Trim()
        if ($_choice -eq '') { $_choice = $_graphPermissionsDefault }
        if ($_choice -notin @('1', '2')) {
          Write-Host "  $_wrn Enter 1 or 2." -ForegroundColor Yellow
        }
      } while ($_choice -notin @('1', '2'))
      $SkipGraphRoleAssignments = $_choice -eq '2'
      if ($SkipGraphRoleAssignments) {
        Write-Host '  Graph role assignments deferred to setup-graph-permissions.ps1.' -ForegroundColor DarkGray
      }
      Write-Host ''
      $_promptsShown = $true
    }

    # ── Required role guidance ────────────────────────────────────────────────
    Write-Hint @(
      'Required Azure RBAC role:  Owner or Contributor  (on the target resource group)'
      '  Owner is also needed for Storage role assignments. Contributor covers the rest.'
      '  For resource provider registration: Contributor or higher at subscription level.'
      ''
      'Required Entra roles:'
      '  Cloud Application Administrator — to create/update the EasyAuth App Registration'
      $(if ($SkipGraphRoleAssignments) {
          '  Privileged Role Administrator   — needed for Graph role assignments (deferred)'
        }
        else {
          '  Privileged Role Administrator   — to assign Graph app roles to the Managed Identity'
        })
      ''
      'PIM eligible roles: activate before running this script, then re-run.'
    )
    Write-Link -Url 'https://portal.azure.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/azurerbac' `
      -Text 'PIM → My roles → Azure resources  (activate eligible role)'
    Write-Link -Url 'https://entra.microsoft.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/aadRoles' `
      -Text 'PIM → My roles → Entra roles  (activate eligible role)'

    $_preflightOk = Test-DeploymentPrerequisite `
      -ResourceGroupName $ResourceGroupName `
      -SkipGraphRoleAssignments:$SkipGraphRoleAssignments

    Save-DeploySessionCache `
      -SubscriptionName $script:SubscriptionName `
      -SubscriptionId $script:SubscriptionId `
      -TenantId $script:TenantId `
      -AzdEnvironmentName $AzdEnvironmentName `
      -ResourceGroupName $ResourceGroupName `
      -AzureLocation $AzureLocation `
      -TenantName $TenantName `
      -FunctionAppName $FunctionAppName `
      -HostingPlan $HostingPlan `
      -DeployAzureMaps:$DeployAzureMaps `
      -AppVersion $AppVersion `
      -Environment $Environment `
      -Criticality $Criticality `
      -EnableMonitoring:$EnableMonitoring `
      -EnableFailureAnomaliesAlert:$EnableFailureAnomaliesAlert `
      -AlwaysReadyInstances $AlwaysReadyInstances `
      -MaximumFlexInstances $MaximumFlexInstances `
      -InstanceMemoryMB $InstanceMemoryMB `
      -SkipGraphRoleAssignments:$SkipGraphRoleAssignments

    if ($PreflightOnly) {
      $_pf = [System.Collections.Generic.List[string]]::new()
      if ($_preflightOk) {
        $_pf.Add('Preflight completed successfully:')
        $_pf.Add('')
        $_pf.Add('  Required tools are available.')
        $_pf.Add('  Azure sign-in and subscription selection succeeded.')
        $_pf.Add('  Required Azure/Entra roles were visible for the signed-in account.')
        $_pf.Add('')
        $_pf.Add('Next run:')
        $_pf.Add('  Re-run the same command without -PreflightOnly to deploy.')
        $_pf.Add('  Use -WhatIf to collect settings and preview azd provision without creating resources.')
      }
      else {
        $_pf.Add('Preflight completed with warnings:')
        $_pf.Add('')
        if ($script:LastPreflightMissingPermissions) {
          $_pf.Add('  Review the missing-role request above.')
          $_pf.Add('  Activate eligible PIM roles or ask an administrator to grant them.')
        }
        if ($script:LastPreflightUnverifiedPermissions) {
          $_pf.Add('  Some role checks could not be completed for the signed-in account.')
          $_pf.Add('  Confirm the required roles in Azure/Entra before deploying.')
        }
        $_pf.Add('  Then re-run the same command.')
      }
      Write-NextStep @($_pf)
      if ($_preflightOk) { exit 0 }
      exit 1
    }

    # ── Confirmation summary ──────────────────────────────────────────────────
    # When all parameters were supplied on the command line or via the session cache (no interactive
    # prompts shown) we display a summary so the operator can verify before
    # the script commits any changes — unless -Confirm:$false or -WhatIf was passed.
    if (-not $_promptsShown -and
      $WhatIfPreference -ne [System.Management.Automation.SwitchParameter]$true -and
      $ConfirmPreference -ne 'None') {
      Write-Host ''
      Write-Host '  Planned operations' -ForegroundColor Cyan
      Write-Host $_sep -ForegroundColor DarkGray
      Write-Host "  azd environment     : $AzdEnvironmentName"
      Write-Host "  Subscription        : $($script:SubscriptionName) ($($script:SubscriptionId))"
      Write-Host "  Resource group      : $ResourceGroupName"
      Write-Host "  Azure location      : $AzureLocation"
      Write-Host "  Environment tag     : $(if ($Environment) { $Environment } else { '(not set)' })"
      Write-Host "  Criticality tag     : $(if ($Criticality) { $Criticality } else { '(not set)' })"
      Write-Host "  SharePoint tenant   : $TenantName"
      Write-Host "  Function App        : $(if ($FunctionAppName) { $FunctionAppName } else { '(auto-generated by Bicep)' })"
      Write-Host "  Hosting plan        : $HostingPlan"
      if ($HostingPlan -eq 'FlexConsumption') {
        Write-Host "  Always-ready        : $AlwaysReadyInstances"
        Write-Host "  Max flex instances  : $MaximumFlexInstances"
        Write-Host "  Instance memory MB  : $InstanceMemoryMB"
      }
      Write-Host "  Azure Maps          : $DeployAzureMaps"
      Write-Host "  Monitoring          : $EnableMonitoring"
      Write-Host "  App version         : $AppVersion"
      if ($SkipGraphRoleAssignments) {
        Write-Host '  Graph roles         : deferred to setup-graph-permissions.ps1'
      }
      else {
        Write-Host '  Graph roles         : assign during deployment'
      }
      Write-Host ''
      Write-Host '  Deployment: azd provision (creates or updates all Azure resources).' -ForegroundColor DarkGray
      Write-Host '  All deployment operations are idempotent — re-running is safe.' -ForegroundColor DarkGray
      Write-Host ''
      do {
        $reply = (Read-Host '  Proceed, re-configure, or abort? [Y/r/n]').Trim()
        if (-not $reply) { $reply = 'y' }
        if ($reply -notmatch '^(?i:y|yes|r|reconfigure|re-configure|n|no)$') {
          Write-Host "  $_wrn Enter Y, R, or N." -ForegroundColor Yellow
          $reply = ''
        }
      } while (-not $reply)
      if ($reply -match '^(?i:r|reconfigure|re-configure)$') {
        $script:ReconfigureMode = $true
        Write-Host ''
        continue
      }
      if ($reply -notmatch '^(?i:y|yes)$') {
        Write-Host 'Aborted.' -ForegroundColor Yellow
        exit 0
      }
      Write-Host ''
    }
    $script:ReconfigureMode = $false
    break
  }

  # ── Deploy ────────────────────────────────────────────────────────────────
  Invoke-AzdProvision `
    -EnvName $AzdEnvironmentName `
    -ResourceGroup $ResourceGroupName `
    -Location $AzureLocation `
    -SharePointTenant $TenantName `
    -Environment $Environment `
    -Criticality $Criticality `
    -AppName $FunctionAppName `
    -Plan $HostingPlan `
    -Maps:$DeployAzureMaps `
    -Version $AppVersion `
    -Monitoring:$EnableMonitoring `
    -FailureAlert:$EnableFailureAnomaliesAlert `
    -AlwaysReadyInstances $AlwaysReadyInstances `
    -FlexInstances $MaximumFlexInstances `
    -InstanceMemoryMB $InstanceMemoryMB `
    -SkipRoles:$SkipGraphRoleAssignments

  # ── Read outputs from azd env ─────────────────────────────────────────────
  # Read the Bicep outputs so the operator gets the exact values to paste into
  # the web part property pane without switching to the Azure portal.
  $_azdFunctionBaseUrl = $null
  $_azdWebPartClientId = $null
  $_azdMiOid = $null
  if (-not $_whatIf) {
    try {
      $_azdEnvVals = azd env get-values 2>$null
      foreach ($_azdLine in $_azdEnvVals) {
        if ($_azdLine -match '^functionAppUrl="?([^"]+)"?') { $_azdFunctionBaseUrl = $Matches[1] }
        elseif ($_azdLine -match '^sponsorApiEndpointUrl="?([^"]+)"?') { $_azdFunctionBaseUrl = $Matches[1] -replace '/api/getGuestSponsors$' }
        elseif ($_azdLine -match '^sponsorApiUrl="?([^"]+)"?') { $_azdFunctionBaseUrl = $Matches[1] -replace '/api/getGuestSponsors$' }
        elseif ($_azdLine -match '^webPartClientId="?([^"]+)"?') { $_azdWebPartClientId = $Matches[1] }
        elseif ($_azdLine -match '^managedIdentityObjectId="?([^"]+)"?') { $_azdMiOid = $Matches[1] }
      }
    }
    catch {
      # Non-fatal — values can be found in the Azure portal.
      Write-Verbose "Could not read azd env values after provision: $_"
    }
    # Fallback: resolve the Function App hostname via Azure CLI if azd env
    # did not contain the functionAppUrl output.
    if (-not $_azdFunctionBaseUrl -and $env:AZURE_RESOURCE_GROUP) {
      try {
        $_azdHostname = (Invoke-AzureCli -Arguments @(
            'functionapp', 'list',
            '--resource-group', $env:AZURE_RESOURCE_GROUP,
            '--query', '[0].defaultHostName',
            '-o', 'tsv'
          )).Trim()
        if ($_azdHostname) { $_azdFunctionBaseUrl = "https://$_azdHostname" }
      }
      catch {
        Write-Verbose "Could not resolve Function App URL after azd provision: $_"
      }
    }
    # Cache the Managed Identity Object ID so setup-graph-permissions.ps1
    # can skip its own prompt when run in the same PowerShell session.
    if ($_azdMiOid) {
      $Global:GsiSetup_ManagedIdentityObjectId = $_azdMiOid
      Write-Host "  $_chk Managed Identity Object ID cached: $_azdMiOid" -ForegroundColor Green
    }
    $Global:GsiSetup_TenantId = $script:TenantId
  }

  # ── NEXT STEPS ────────────────────────────────────────────────────────────
  $_ns = [System.Collections.Generic.List[string]]::new()
  if ($_whatIf) {
    $_ns.Add('WhatIf preview completed:')
    $_ns.Add('')
    $_ns.Add('  azd provision --preview completed.')
    $_ns.Add('  No Azure resources were created or changed by the preview.')
    $_ns.Add('  Local azd environment values were updated for this preview run.')
    $_ns.Add('  Re-run without -WhatIf to execute the deployment with the values above.')
    $_ns.Add('')
    $_ns.Add('Expected web part configuration after a real deployment:')
  }
  else {
    $_ns.Add('Deployment completed successfully:')
    $_ns.Add('')
    $_ns.Add('  App Registration  — created/updated by Bicep (Graph extension)')
    $_ns.Add('  Azure resources   — deployed by Bicep')
    if ($SkipGraphRoleAssignments) {
      $_ns.Add('  Graph permissions — DEFERRED: run setup-graph-permissions.ps1')
    }
    else {
      $_ns.Add('  Graph permissions — assigned by Bicep (Graph extension)')
    }
    $_ns.Add('  Function App      — restarted by post-provision hook')
    $_ns.Add('')
    $_ns.Add('Configure the web part (SharePoint property pane → Guest Sponsor API):')
  }
  if (-not $_whatIf -and $SkipGraphRoleAssignments) {
    $_ns.Add('')
    $_ns.Add('Graph permissions — run setup-graph-permissions.ps1:')
    $_miDisplay = if ($_azdMiOid) { $_azdMiOid } else { 'run: azd env get-values → managedIdentityObjectId' }
    $_ns.Add("  -ManagedIdentityObjectId : $_miDisplay")
    $_ns.Add("  -TenantId                : $($script:TenantId)")
    $_ns.Add('')
    $_graphPermScript = Join-Path $PSScriptRoot 'setup-graph-permissions.ps1'
    if (Test-Path $_graphPermScript) {
      $_ns.Add("  & '$_graphPermScript'")
    }
  }
  if (-not $_whatIf) {
    # Keep current wording for successful real deployments.
  }
  if ($_whatIf) {
    $_ns.Add('  Base URL               : available after deployment')
    $_ns.Add('  Application (client) ID: available after deployment')
  }
  elseif ($_azdFunctionBaseUrl) {
    $_ns.Add("  Base URL               : $_azdFunctionBaseUrl")
  }
  else {
    $_ns.Add('  Base URL               : see Function App hostname in the Azure portal')
    if ($env:AZURE_RESOURCE_GROUP) {
      $_ns.Add("  Resource group         : $($env:AZURE_RESOURCE_GROUP)")
    }
  }
  if (-not $_whatIf) {
    if ($_azdWebPartClientId) {
      $_ns.Add("  Application (client) ID: $_azdWebPartClientId")
    }
    else {
      $_ns.Add('  Application (client) ID: see post-provision hook output above')
    }
    $_ns.Add('')
    $_ns.Add('Finish in SharePoint:')
    $_ns.Add('  1. Open the guest landing page in edit mode.')
    $_ns.Add('  2. Open the Guest Sponsor Info web part property pane.')
    $_ns.Add('  3. Paste the Base URL and Application (client) ID under Guest Sponsor API.')
    $_ns.Add('  4. Save or publish the page.')
    $_ns.Add('')
    $_ns.Add('If you need to retry, re-run the same command. The azd environment and')
    $_ns.Add('Azure resources are reused or updated idempotently.')
  }
  Write-NextStep @($_ns)
}
catch {
  Write-Failure @(
    'The deployment did not finish.'
    ''
    'You can safely re-run the same install command after fixing the error above.'
    'The existing azd environment will be reused, and Bicep will update or skip'
    'resources that were already created.'
    ''
    'Do not delete partially created Azure resources unless the error message'
    'explicitly tells you to do so.'
  )
  throw
}
#endregion
