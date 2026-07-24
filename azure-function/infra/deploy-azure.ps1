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

.PARAMETER AzureLoginMode
  Azure CLI login mode. "auto" (default) chooses browser login for local
  consoles and falls back to device code in Azure Cloud Shell and other
  remote/headless terminals. Use "browser" or "device-code" to override that
  detection explicitly. Device code sign-in must be permitted for the account,
  and an admin account restricted to a Privileged Access Workstation has to
  confirm the code on that workstation.

.PARAMETER TenantName
    SharePoint tenant short name — the part before .sharepoint.com
    (e.g. "contoso" for contoso.sharepoint.com).

.PARAMETER FunctionAppName
    Globally unique Function App name (2-58 characters). Bicep auto-generates
    one (e.g. "gsi-a1b2c3d4") when left blank.

.PARAMETER DeployAzureMaps
    Deploy an Azure Maps account for address map rendering. Defaults to true.

.PARAMETER AppVersion
  Advanced override for the Function package version tag. Defaults to
  "latest" when deploy-azure.ps1 runs directly. When invoked via
  install.ps1, the installer usually aligns this with `-Version`.

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
  Number of always-ready (pre-warmed) instances for Flex Consumption. Defaults to 0.

.PARAMETER InstanceMemoryMB
  Memory size per Flex Consumption instance in MB. Defaults to 512.

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
  [ValidateSet('auto', 'browser', 'device-code')]
  [string]$AzureLoginMode = 'auto',
  [string]$TenantName,
  [string]$FunctionAppName,
  [bool]$DeployAzureMaps = $true,
  [string]$AppVersion = 'latest',
  [AllowEmptyString()]
  [string]$Environment = '',
  [AllowEmptyString()]
  [string]$Criticality = '',
  [bool]$EnableMonitoring = $true,
  [bool]$EnableFailureAnomaliesAlert = $false,
  [int]$AlwaysReadyInstances = 0,
  [int]$MaximumFlexInstances = 10,
  [ValidateSet(512, 2048)]
  [int]$InstanceMemoryMB = 512,
  [bool]$SkipGraphRoleAssignments = $false,
  [Parameter(DontShow)]
  [string]$InstallerVersion = '',
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
$script:InferredDeployParameters = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$script:UsedDeploySessionCache = $false
$script:UsedInferredDeployDefaults = $false
$script:ReconfigureMode = $false
$script:AzureSessionRoot = $null
$script:AzureCliConfigDir = $null
$script:AzdConfigDir = $null
$script:AzureAuthIsolationMode = ''
$script:AzureSessionRootPrefix = 'gsi-azure-session'
# Well-known first-party application ID of the Azure CLI. Microsoft Graph
# tokens must be issued to this application for the Entra bootstrap to have
# the delegated scopes it needs.
$script:AzureCliApplicationId = '04b07795-8ddb-461a-bbee-02f9e1bf7b46'
$script:AzdUsesAzureCliAuth = $false
$script:CanRepairGraphPermissionsInThisRun = $false
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
# Turn a raw error message into printable box lines. Azure CLI reports failed
# deployments as a single long JSON blob, so hard-wrap it and cap the output
# instead of flooding the console.
function Format-ErrorDetailLine {
  param(
    [AllowEmptyString()][string]$Message,
    [int]$MaxLineLength = 96,
    [int]$MaxLineCount = 16
  )

  if ([string]::IsNullOrWhiteSpace($Message)) { return @() }

  $_detail = [System.Collections.Generic.List[string]]::new()
  foreach ($_rawLine in ($Message -split "`r?`n")) {
    $_line = $_rawLine.Trim()
    if (-not $_line) { continue }
    while ($_line.Length -gt $MaxLineLength) {
      $_detail.Add("  $($_line.Substring(0, $MaxLineLength))")
      $_line = $_line.Substring($MaxLineLength)
    }
    $_detail.Add("  $_line")
  }

  if ($_detail.Count -gt $MaxLineCount) {
    $_omitted = $_detail.Count - $MaxLineCount
    return @($_detail[0..($MaxLineCount - 1)]) + @("  ... ($_omitted more line(s) omitted)")
  }

  return @($_detail)
}

# Script-level trap: on authorization errors, print role guidance instead of a
# raw exception — but always keep the original message visible, so a
# misclassified error can still be diagnosed. Other errors fall through to
# PowerShell's own error output.
trap {
  $_errMsg = if ($_.Exception -and $_.Exception.Message) { $_.Exception.Message } else { [string]$_ }

  # A denial signal must be present before any role guidance is printed.
  # Matching on a Graph resource type alone is not enough: Bicep compile
  # errors and ARM validation failures name 'Microsoft.Graph/...' too, and
  # were previously reported as missing permissions.
  $_deniedSignal = $_errMsg -match '(?i)\bAuthorization_RequestDenied\b|insufficient\s+privileges|\bForbidden\b|\b403\b'
  $_graphContext = $_errMsg -match '(?i)\bAuthorization_RequestDenied\b|graph\.microsoft\.com|Microsoft\.Graph/|appRoleAssignments?|appRoleAssignedTo|microsoft\.directory/'
  $_azureContext = $_errMsg -match '(?i)AuthorizationFailed|LinkedAuthorizationFailed|does not have authorization'

  if ($_deniedSignal -and $_graphContext) {
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
      ''
      'Original error:'
      (Format-ErrorDetailLine -Message $_errMsg)
    )
    Write-Link -Url 'https://entra.microsoft.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/aadRoles' `
      -Text 'PIM → My roles → Entra roles  (activate eligible roles)'
    Write-Link -Url 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/RolesManagementMenuBlade/~/AllRoles' `
      -Text 'Entra admin center → Roles and administrators'
    return
  }
  if ($_azureContext -or $_deniedSignal) {
    Write-Failure -Lines @(
      'The request was denied — your account lacks the required permissions.'
      ''
      'Required Azure RBAC role:  Contributor  (on the target subscription or resource group)'
      '  Owner also works. Provider registration, whether manual or implicit, needs Contributor or higher at subscription level.'
      ''
      'If your role is eligible (PIM): activate it, then re-run.'
      'If you do not have the role yet: request it from your Azure admin.'
      ''
      'Original error:'
      (Format-ErrorDetailLine -Message $_errMsg)
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
  $_stderrFile = Join-Path ([System.IO.Path]::GetTempPath()) "gsi-az-stderr-$([guid]::NewGuid().ToString('n')).log"
  try {
    $env:PYTHONWARNINGS = if ($_previousPythonWarnings) {
      "ignore::SyntaxWarning,$_previousPythonWarnings"
    }
    else {
      'ignore::SyntaxWarning'
    }

    $output = & $script:AzPath @Arguments 2>$_stderrFile
    if ($LASTEXITCODE -ne 0) {
      $_stderrText = ''
      if (Test-Path -Path $_stderrFile) {
        $_stderrText = (Get-Content -Path $_stderrFile -Raw -ErrorAction SilentlyContinue).Trim()
      }

      if ($_stderrText) {
        throw "Azure CLI failed (exit code $LASTEXITCODE): az $($Arguments -join ' ')`n$_stderrText"
      }

      throw "Azure CLI failed (exit code $LASTEXITCODE): az $($Arguments -join ' ')"
    }
  }
  finally {
    Remove-Item -Path $_stderrFile -Force -ErrorAction SilentlyContinue

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
  $_authSessionBehavior = if (Test-AzureCloudShell) {
    'Azure auth is isolated to this run, so the wizard signs in with the Azure CLI instead of reusing the Cloud Shell identity.'
  }
  else {
    'Azure auth is isolated to this PowerShell console so older az/azd logins do not leak into this run.'
  }

  Write-Hint @(
    'Before deployment this wizard checks the local tools and signs in to Azure.'
    $_authSessionBehavior
    ''
    'It can install missing tools when needed:'
    '  PowerShell bootstrapper (install.sh): PowerShell 7+'
    '  This deployment wizard: Azure CLI (az) and Azure Developer CLI (azd)'
    '  macOS fallback: Homebrew when Azure CLI installation needs it'
    ''
    'Interactive steps you may see: tool installation prompts, sudo/admin password prompts,'
    'az login, subscription selection, and deployment parameter prompts.'
  )
}

function Test-AzureCloudShell {
  return (
    $env:ACC_TERMID -or
    $env:ACC_LOCATION -or
    $env:CLOUD_SHELL -or
    (($env:AZUREPS_HOST_ENVIRONMENT -as [string]) -match 'cloud-shell')
  ) -as [bool]
}

function Get-AzureCliLoginPreference {
  if ($AzureLoginMode -eq 'browser') {
    return [pscustomobject]@{
      Mode       = 'browser'
      StatusText = 'browser (forced by -AzureLoginMode)'
      DetailText = 'The script will not fall back to device code for this run.'
      StartText  = 'Starting az login in your default browser...'
    }
  }

  if ($AzureLoginMode -eq 'device-code') {
    return [pscustomobject]@{
      Mode       = 'device-code'
      StatusText = 'device code (forced by -AzureLoginMode)'
      DetailText = 'Use this when browser launch is unavailable or intentionally disabled.'
      StartText  = 'Starting az login with device code...'
    }
  }

  if (Test-AzureCloudShell) {
    return [pscustomobject]@{
      Mode       = 'device-code'
      StatusText = 'device code (auto for Azure Cloud Shell)'
      DetailText = 'The Cloud Shell container has no browser of its own.'
      StartText  = 'Starting az login with device code...'
    }
  }

  $_isRemoteTerminal = (
    $env:SSH_CONNECTION -or
    $env:SSH_CLIENT -or
    $env:SSH_TTY -or
    $env:CODESPACES -or
    $env:GITHUB_CODESPACES_PORT_FORWARDING_DOMAIN -or
    $env:REMOTE_CONTAINERS -or
    ($env:TERM_PROGRAM -eq 'vscode' -and (Test-Path -Path '/.dockerenv'))
  ) -as [bool]
  $_isHeadlessLinux = (
    -not (Test-WindowsHost) -and
    -not (Test-MacOSHost) -and
    -not $env:DISPLAY -and
    -not $env:WAYLAND_DISPLAY
  )

  if ($_isRemoteTerminal -or $_isHeadlessLinux) {
    return [pscustomobject]@{
      Mode       = 'device-code'
      StatusText = 'device code (auto for remote/headless terminal)'
      DetailText = 'This avoids browser launch delays when the terminal session has no direct GUI.'
      StartText  = 'Starting az login with device code...'
    }
  }

  return [pscustomobject]@{
    Mode       = 'browser'
    StatusText = 'browser (default for local console)'
    DetailText = ''
    StartText  = 'Starting az login in your default browser...'
  }
}

function Get-CommandVersionText {
  param(
    [Parameter(Mandatory)][string]$Command,
    [Parameter(Mandatory)][string[]]$Arguments
  )

  try {
    $output = & $Command @Arguments 2>$null
    if ($LASTEXITCODE -ne 0) { return $null }
    $line = @($output | Where-Object { $_ } | Select-Object -First 1)[0]
    if ($line) { return $line.Trim() }
  }
  catch {
    return $null
  }

  return $null
}

function Get-AzureCliVersionText {
  $_version = Get-CommandVersionText -Command $script:AzPath -Arguments @(
    'version', '--query', '"azure-cli"', '-o', 'tsv'
  )
  if ($_version) {
    return $_version
  }

  $_fallbackVersion = Get-CommandVersionText -Command $script:AzPath -Arguments @('--version')
  if ($_fallbackVersion -match '^azure-cli\s+(.+)$') {
    return $Matches[1].Trim()
  }
  if ($_fallbackVersion) {
    return $_fallbackVersion
  }

  return 'available'
}

function Get-AzureCliBicepVersionText {
  $_versionLine = Get-CommandVersionText -Command $script:AzPath -Arguments @('bicep', 'version')
  if (-not $_versionLine) {
    return $null
  }

  $_cleanVersionLine = $_versionLine.Trim()
  if ($_cleanVersionLine -match '([0-9]+\.[0-9]+\.[0-9]+(?:[-+][0-9A-Za-z.-]+)?)') {
    return $Matches[1]
  }

  return $_cleanVersionLine
}

function Update-AzureCliBicep {
  [CmdletBinding(SupportsShouldProcess)]
  param()

  # 'az bicep install' pins whatever version was current at install time and is
  # never refreshed afterwards. entra-auth.bicep needs a release that
  # understands the 'extension' keyword and the 'extensions' block in
  # bicepconfig.json, so refresh a pre-existing backend rather than trusting a
  # binary cached in $AZURE_CONFIG_DIR by an earlier run.
  if (-not $PSCmdlet.ShouldProcess('Azure CLI Bicep backend', 'Upgrade to the latest release')) {
    return
  }

  try {
    Invoke-AzureCliQuiet -Arguments @('bicep', 'upgrade') | Out-Null
  }
  catch {
    # Non-fatal: an outdated backend still compiles most templates, and a
    # template it cannot compile now reports the real Bicep error.
    Write-Verbose "Could not upgrade the Azure CLI Bicep backend: $_"
    Write-Host "  $_wrn Could not check for a newer Azure CLI Bicep version; continuing with the installed one." -ForegroundColor Yellow
    Write-Host '       Run ''az bicep upgrade'' manually if a template fails to compile.' -ForegroundColor DarkGray
  }
}

function Install-AzureCliBicepIfNeeded {
  $_bicepVersion = Get-AzureCliBicepVersionText
  if ($_bicepVersion) {
    Update-AzureCliBicep
    return
  }

  Write-Host "  $_wrn Azure CLI Bicep backend is not available." -ForegroundColor Yellow
  $answer = (Read-Host '  Install it now via Azure CLI? [Y/n]').Trim()
  if ($answer -ne '' -and $answer -notmatch '^[Yy]') {
    throw @(
      'Azure CLI Bicep support is required because this wizard runs small pre- and post-provision Bicep deployments before azd provision.',
      'Install it manually with: az bicep install',
      'Then re-run this script.'
    ) -join ' '
  }

  Write-Host "  $_arr Installing Azure CLI Bicep backend..." -ForegroundColor Cyan
  Invoke-AzureCli -Arguments @('bicep', 'install')

  $_bicepVersion = Get-AzureCliBicepVersionText
  if (-not $_bicepVersion) {
    throw 'Azure CLI Bicep backend was installed, but the current session still cannot use it. Open a new terminal and re-run the script.'
  }

  Write-Host "  $_chk Azure CLI Bicep backend is available." -ForegroundColor Green
}

function Get-AzdVersionText {
  $_versionLine = Get-CommandVersionText -Command $script:AzdPath -Arguments @('version', '--no-prompt')
  if (-not $_versionLine) {
    $_versionLine = Get-CommandVersionText -Command $script:AzdPath -Arguments @('version')
  }

  if (-not $_versionLine) {
    return 'available'
  }

  $_cleanVersionLine = $_versionLine.Trim()
  if ($_cleanVersionLine -match '(?i)\bazd version\s+([0-9]+\.[0-9]+\.[0-9]+(?:[-+][0-9A-Za-z.-]+)?)') {
    $_version = $Matches[1]
    $_channel = $null
    if ($_cleanVersionLine -match '\((stable|beta|preview|alpha|nightly|canary|rc[^)]*)\)') {
      $_channel = $Matches[1]
    }

    if ($_channel) {
      return "$($_version) ($($_channel.ToLowerInvariant()))"
    }

    return $_version
  }

  $_cleanVersionLine = $_cleanVersionLine -replace '^\s*azd version\s+', ''
  $_cleanVersionLine = $_cleanVersionLine -replace '\s+\(commit [0-9a-f]{7,40}\)', ''
  $_cleanVersionLine = $_cleanVersionLine -replace '\s{2,}', ' '

  if ($_cleanVersionLine) {
    return $_cleanVersionLine.Trim()
  }

  return 'available'
}

function Write-PreflightStatus {
  param(
    [Parameter(Mandatory)][string]$Label,
    [Parameter(Mandatory)][string]$Value
  )

  Write-Host ('  {0} {1,-20}: {2}' -f $_chk, $Label, $Value) -ForegroundColor Green
}

function Show-CloudShellSessionBanner {
  if (-not (Test-AzureCloudShell)) {
    return
  }

  Write-Hint @(
    'Azure Cloud Shell detected.'
    ''
    'A device code sign-in follows. This run signs in with the Azure CLI instead'
    'of reusing the Cloud Shell identity, which is not permitted to create the'
    'Entra App Registration. Your existing Cloud Shell login stays intact: this'
    'run keeps its own config directories and does not touch ~/.azure.'
    ''
    'Two things have to be in place for that sign-in:'
    '  - Device code sign-in must be allowed for your account. Some tenants block'
    '    it via Conditional Access; run the installer from a local console then.'
    '  - If your admin account is restricted to a Privileged Access Workstation,'
    '    confirm the device code on that PAW. Other devices will be rejected.'
    ''
    'If azd is missing, the wizard can install it into your Cloud Shell home'
    'directory without sudo. Persisted Cloud Shell storage usually keeps that'
    'azd install for later sessions.'
  )
}

function Show-ToolVersion {
  $pwshVersion = $PSVersionTable.PSVersion.ToString()
  $azVersion = Get-AzureCliVersionText
  $azBicepVersion = Get-AzureCliBicepVersionText
  $azdVersion = Get-AzdVersionText
  $azLoginPreference = Get-AzureCliLoginPreference

  Write-Host ''
  Write-Host '  Tool preflight' -ForegroundColor Cyan
  Write-Host $_sep -ForegroundColor DarkGray
  Write-PreflightStatus -Label 'PowerShell' -Value $pwshVersion
  Write-PreflightStatus -Label 'Azure CLI (az)' -Value $azVersion
  Write-PreflightStatus -Label 'Azure CLI Bicep' -Value $(if ($azBicepVersion) { $azBicepVersion } else { 'not available' })
  Write-PreflightStatus -Label 'Azure Developer CLI' -Value $azdVersion
  Write-PreflightStatus -Label 'az login mode' -Value $azLoginPreference.StatusText
  if ($azLoginPreference.DetailText) {
    Write-Host "       $($azLoginPreference.DetailText)" -ForegroundColor DarkGray
  }
  if ($script:AzureAuthIsolationMode -eq 'isolated') {
    Write-PreflightStatus -Label 'Azure auth session' -Value 'isolated to this PowerShell console'
    Write-Host '       Existing logins in ~/.azure and ~/.azd are ignored for this console.' -ForegroundColor DarkGray
    Write-Host '       Close the console to drop this isolated login context.' -ForegroundColor DarkGray
  }
  elseif ($script:AzureAuthIsolationMode -eq 'caller-supplied') {
    Write-PreflightStatus -Label 'Azure auth session' -Value 'using caller-provided config dirs'
  }
  if ($script:AzdUsesAzureCliAuth) {
    Write-PreflightStatus -Label 'azd auth mode' -Value 'reuses the active Azure CLI login'
    Write-Host '       azd does not keep an independent tenant login for this script run.' -ForegroundColor DarkGray
  }
  Write-Host '       azd uses its own scoped Bicep CLI during azd provision.' -ForegroundColor DarkGray

  Show-CloudShellSessionBanner
}

function Initialize-AzureSessionRoot {
  if ($Global:GsiDeploy_AzureSessionRoot) {
    $_existingRoot = [string]$Global:GsiDeploy_AzureSessionRoot
    if ($_existingRoot -and (Test-Path $_existingRoot)) {
      return $_existingRoot
    }
  }

  $_tempRoot = [System.IO.Path]::GetTempPath()
  $_cutoffUtc = [System.DateTime]::UtcNow.AddDays(-7)
  foreach ($_staleDir in Get-ChildItem -Path $_tempRoot -Directory -Filter "$($script:AzureSessionRootPrefix)-*" -ErrorAction SilentlyContinue) {
    try {
      if ($_staleDir.LastWriteTimeUtc -lt $_cutoffUtc) {
        Remove-Item -Path $_staleDir.FullName -Recurse -Force -ErrorAction Stop
      }
    }
    catch {
      Write-Verbose "Could not remove stale Azure session directory '$($_staleDir.FullName)': $_"
    }
  }

  $_sessionRoot = Join-Path $_tempRoot "$($script:AzureSessionRootPrefix)-$PID-$([System.Guid]::NewGuid().ToString('n'))"
  $null = New-Item -Path $_sessionRoot -ItemType Directory -Force
  $Global:GsiDeploy_AzureSessionRoot = $_sessionRoot

  return $_sessionRoot
}

function Initialize-AzureAuthIsolation {
  if ($script:AzureCliConfigDir -and $script:AzdConfigDir) {
    return
  }

  # Azure Cloud Shell is deliberately not special-cased here. Reusing its
  # shared ~/.azure meant reusing a token issued to the Cloud Shell
  # application, whose Microsoft Graph scopes do not cover creating the EasyAuth
  # App Registration — see Confirm-AzureCliGraphClient. An isolated config
  # directory makes the wizard sign in with the Azure CLI itself and leaves the
  # ambient Cloud Shell login untouched.
  $_managedRoot = if ($Global:GsiDeploy_AzureSessionRoot) { [string]$Global:GsiDeploy_AzureSessionRoot } else { '' }
  $_managedAzureConfigDir = if ($_managedRoot) { Join-Path $_managedRoot '.azure' } else { '' }
  $_managedAzdConfigDir = if ($_managedRoot) { Join-Path $_managedRoot '.azd' } else { '' }
  $_hasAzureConfigDir = -not [string]::IsNullOrWhiteSpace($env:AZURE_CONFIG_DIR)
  $_hasAzdConfigDir = -not [string]::IsNullOrWhiteSpace($env:AZD_CONFIG_DIR)
  $_reusingManagedDirs = (
    $_managedRoot -and
    $_hasAzureConfigDir -and
    $_hasAzdConfigDir -and
    $env:AZURE_CONFIG_DIR -eq $_managedAzureConfigDir -and
    $env:AZD_CONFIG_DIR -eq $_managedAzdConfigDir
  )

  if ($_reusingManagedDirs) {
    $script:AzureSessionRoot = $_managedRoot
    $script:AzureAuthIsolationMode = 'isolated'
  }
  elseif ($_hasAzureConfigDir -and $_hasAzdConfigDir) {
    $script:AzureAuthIsolationMode = 'caller-supplied'
  }
  else {
    $script:AzureSessionRoot = Initialize-AzureSessionRoot
    $env:AZURE_CONFIG_DIR = Join-Path $script:AzureSessionRoot '.azure'
    $env:AZD_CONFIG_DIR = Join-Path $script:AzureSessionRoot '.azd'
    $script:AzureAuthIsolationMode = 'isolated'
  }

  foreach ($_configDir in @($env:AZURE_CONFIG_DIR, $env:AZD_CONFIG_DIR)) {
    $null = New-Item -Path $_configDir -ItemType Directory -Force
  }

  $script:AzureCliConfigDir = $env:AZURE_CONFIG_DIR
  $script:AzdConfigDir = $env:AZD_CONFIG_DIR
}

function Enable-AzdAzureCliAuth {
  foreach ($_configKey in @('auth.useAzCliAuth', 'auth.useAzureCliCredentials')) {
    & $script:AzdPath 'config' 'set' $_configKey 'true' | Out-Null
    if ($LASTEXITCODE -eq 0) {
      $script:AzdUsesAzureCliAuth = $true
      return
    }
  }

  throw 'Azure Developer CLI could not be configured to reuse the Azure CLI login for this session.'
}

function Enable-AzdDeploymentStackSupport {
  $_configKey = 'alpha.deployment.stacks'
  $_currentValue = ''

  try {
    $_configOutput = & $script:AzdPath 'config' 'get' $_configKey 2>$null
    if ($LASTEXITCODE -eq 0) {
      $_currentValue = [string](@($_configOutput | Where-Object { $_ } | Select-Object -First 1)[0]).Trim()
    }
  }
  catch {
    $_currentValue = ''
  }

  if ($_currentValue -notmatch '^(?i:on|true)$') {
    & $script:AzdPath 'config' 'set' $_configKey 'on' | Out-Null
    if ($LASTEXITCODE -ne 0) {
      throw @(
        'Azure Developer CLI deployment stacks could not be enabled for this session.',
        'Run ''azd config set alpha.deployment.stacks on'' and retry the deployment.'
      ) -join ' '
    }
  }

  Write-Host "  $_chk azd deployment stacks: enabled for the active azd config." -ForegroundColor Green
}

function Get-DeployedFunctionAppInfo {
  param(
    [Parameter(Mandatory)][string]$ResourceGroup,
    [string]$FunctionAppName
  )

  try {
    $_query = '{name:name,defaultHostName:defaultHostName,principalId:identity.principalId}'
    $_appJson = if ($FunctionAppName) {
      Invoke-AzureCliQuiet -Arguments @(
        'functionapp', 'show',
        '--resource-group', $ResourceGroup,
        '--name', $FunctionAppName,
        '--query', $_query,
        '-o', 'json'
      )
    }
    else {
      Invoke-AzureCliQuiet -Arguments @(
        'functionapp', 'list',
        '--resource-group', $ResourceGroup,
        '--query', "[0].$_query",
        '-o', 'json'
      )
    }

    $_appJsonText = ($_appJson -join "`n").Trim()
    if (-not $_appJsonText -or $_appJsonText -eq 'null') {
      return $null
    }

    return $_appJsonText | ConvertFrom-Json
  }
  catch {
    Write-Verbose "Could not resolve Function App metadata from Azure CLI: $_"
    return $null
  }
}

function Test-ResourceGroupPresence {
  param([Parameter(Mandatory)][string]$ResourceGroupName)

  try {
    return ((Invoke-AzureCliQuiet -Arguments @('group', 'exists', '--name', $ResourceGroupName)).Trim() -eq 'true')
  }
  catch {
    return $false
  }
}

function Get-AzdEnvironmentStoredValue {
  param(
    [Parameter(Mandatory)][string]$EnvName,
    [Parameter(Mandatory)][string]$Name
  )

  $_envFile = Join-Path (Get-RepoRoot) ".azure/$EnvName/.env"
  if (-not (Test-Path $_envFile)) {
    return ''
  }

  $_pattern = '^' + [regex]::Escape($Name) + '="?([^\"]*)"?$'
  $_matchedLine = Get-Content -Path $_envFile -ErrorAction SilentlyContinue |
  Where-Object { $_ -match $_pattern } |
  Select-Object -First 1
  if (-not $_matchedLine) {
    return ''
  }

  return $Matches[1]
}

function Get-AzdEnvironmentStoredBooleanValue {
  param(
    [Parameter(Mandatory)][string]$EnvName,
    [Parameter(Mandatory)][string]$Name
  )

  $_value = Get-AzdEnvironmentStoredValue -EnvName $EnvName -Name $Name
  if ($_value -match '^(?i:true|false)$') {
    return $_value -match '^(?i:true)$'
  }

  return $null
}

function Get-AzdEnvironmentStoredIntegerValue {
  param(
    [Parameter(Mandatory)][string]$EnvName,
    [Parameter(Mandatory)][string]$Name
  )

  $_value = Get-AzdEnvironmentStoredValue -EnvName $EnvName -Name $Name
  $_parsedValue = 0
  if ([int]::TryParse($_value, [ref]$_parsedValue)) {
    return $_parsedValue
  }

  return $null
}

function Use-InferredDeployValue {
  param(
    [Parameter(Mandatory)][string]$Key,
    [Parameter(Mandatory)][ref]$Target,
    $Value,
    [switch]$AllowEmptyString,
    [switch]$ReplaceExistingInferred
  )

  if ($script:ExplicitDeployParameters.Contains($Key) -or $script:CachedDeployParameters.Contains($Key)) {
    return
  }
  if ($script:InferredDeployParameters.Contains($Key) -and -not $ReplaceExistingInferred) {
    return
  }
  if ($null -eq $Value) {
    return
  }
  if (-not $AllowEmptyString -and $Value -is [string] -and [string]::IsNullOrWhiteSpace($Value)) {
    return
  }

  $Target.Value = $Value
  $null = $script:InferredDeployParameters.Add($Key)
  $script:UsedInferredDeployDefaults = $true
}

function Get-TagValue {
  param(
    [AllowNull()][object]$Tags,
    [Parameter(Mandatory)][string]$Name
  )

  if ($null -eq $Tags) {
    return ''
  }

  if ($Tags -is [System.Collections.IDictionary]) {
    foreach ($_key in $Tags.Keys) {
      if ([string]$_key -ieq $Name) {
        return [string]$Tags[$_key]
      }
    }

    return ''
  }

  $_property = $Tags.PSObject.Properties | Where-Object { $_.Name -ieq $Name } | Select-Object -First 1
  if ($_property) {
    return [string]$_property.Value
  }

  return ''
}

function Use-AzdEnvironmentStoredDeploymentProfile {
  param(
    [Parameter(Mandatory)][string]$EnvName,
    [Parameter(Mandatory)][ref]$ResourceGroupName,
    [Parameter(Mandatory)][ref]$AzureLocation,
    [Parameter(Mandatory)][ref]$TenantName,
    [Parameter(Mandatory)][ref]$FunctionAppName,
    [Parameter(Mandatory)][ref]$DeployAzureMaps,
    [Parameter(Mandatory)][ref]$Environment,
    [Parameter(Mandatory)][ref]$Criticality,
    [Parameter(Mandatory)][ref]$EnableMonitoring,
    [Parameter(Mandatory)][ref]$EnableFailureAnomaliesAlert,
    [Parameter(Mandatory)][ref]$AlwaysReadyInstances,
    [Parameter(Mandatory)][ref]$MaximumFlexInstances,
    [Parameter(Mandatory)][ref]$InstanceMemoryMB
  )

  if ([string]::IsNullOrWhiteSpace($EnvName)) {
    return
  }

  Use-InferredDeployValue -Key 'ResourceGroupName' -Target $ResourceGroupName `
    -Value (Get-AzdEnvironmentStoredValue -EnvName $EnvName -Name 'AZURE_RESOURCE_GROUP') -ReplaceExistingInferred
  Use-InferredDeployValue -Key 'AzureLocation' -Target $AzureLocation `
    -Value (Get-AzdEnvironmentStoredValue -EnvName $EnvName -Name 'AZURE_LOCATION') -ReplaceExistingInferred
  Use-InferredDeployValue -Key 'TenantName' -Target $TenantName `
    -Value (Get-AzdEnvironmentStoredValue -EnvName $EnvName -Name 'AZURE_SHAREPOINT_TENANT_NAME') -ReplaceExistingInferred

  $_storedFunctionAppName = Get-AzdEnvironmentStoredValue -EnvName $EnvName -Name 'functionAppName'
  if (-not $_storedFunctionAppName) {
    $_storedFunctionAppName = Get-AzdEnvironmentStoredValue -EnvName $EnvName -Name 'AZURE_FUNCTION_APP_NAME'
  }
  Use-InferredDeployValue -Key 'FunctionAppName' -Target $FunctionAppName `
    -Value $_storedFunctionAppName -AllowEmptyString -ReplaceExistingInferred

  Use-InferredDeployValue -Key 'Environment' -Target $Environment `
    -Value (Get-AzdEnvironmentStoredValue -EnvName $EnvName -Name 'AZURE_TAG_ENVIRONMENT') `
    -AllowEmptyString -ReplaceExistingInferred
  Use-InferredDeployValue -Key 'Criticality' -Target $Criticality `
    -Value (Get-AzdEnvironmentStoredValue -EnvName $EnvName -Name 'AZURE_TAG_CRITICALITY') `
    -AllowEmptyString -ReplaceExistingInferred
  Use-InferredDeployValue -Key 'DeployAzureMaps' -Target $DeployAzureMaps `
    -Value (Get-AzdEnvironmentStoredBooleanValue -EnvName $EnvName -Name 'AZURE_DEPLOY_AZURE_MAPS') -ReplaceExistingInferred
  Use-InferredDeployValue -Key 'EnableMonitoring' -Target $EnableMonitoring `
    -Value (Get-AzdEnvironmentStoredBooleanValue -EnvName $EnvName -Name 'AZURE_ENABLE_MONITORING') -ReplaceExistingInferred
  Use-InferredDeployValue -Key 'EnableFailureAnomaliesAlert' -Target $EnableFailureAnomaliesAlert `
    -Value (Get-AzdEnvironmentStoredBooleanValue -EnvName $EnvName -Name 'AZURE_ENABLE_FAILURE_ANOMALIES_ALERT') -ReplaceExistingInferred
  Use-InferredDeployValue -Key 'AlwaysReadyInstances' -Target $AlwaysReadyInstances `
    -Value (Get-AzdEnvironmentStoredIntegerValue -EnvName $EnvName -Name 'AZURE_ALWAYS_READY_INSTANCES') -ReplaceExistingInferred
  Use-InferredDeployValue -Key 'MaximumFlexInstances' -Target $MaximumFlexInstances `
    -Value (Get-AzdEnvironmentStoredIntegerValue -EnvName $EnvName -Name 'AZURE_MAXIMUM_FLEX_INSTANCES') -ReplaceExistingInferred
  Use-InferredDeployValue -Key 'InstanceMemoryMB' -Target $InstanceMemoryMB `
    -Value (Get-AzdEnvironmentStoredIntegerValue -EnvName $EnvName -Name 'AZURE_INSTANCE_MEMORY_MB') -ReplaceExistingInferred
}

function Use-LiveDeploymentProfile {
  param(
    [Parameter(Mandatory)][string]$ResourceGroupName,
    [AllowEmptyString()][string]$FunctionAppName,
    [Parameter(Mandatory)][ref]$AzureLocation,
    [Parameter(Mandatory)][ref]$Environment,
    [Parameter(Mandatory)][ref]$Criticality,
    [Parameter(Mandatory)][ref]$FunctionAppNameTarget
  )

  if (-not (Test-ResourceGroupPresence -ResourceGroupName $ResourceGroupName)) {
    return
  }

  try {
    $_groupJson = Invoke-AzureCliQuiet -Arguments @(
      'group', 'show',
      '--name', $ResourceGroupName,
      '--query', '{location:location,tags:tags}',
      '-o', 'json'
    )
    $_groupText = ($_groupJson -join "`n").Trim()
    if ($_groupText -and $_groupText -ne 'null') {
      $_groupInfo = $_groupText | ConvertFrom-Json
      Use-InferredDeployValue -Key 'AzureLocation' -Target $AzureLocation `
        -Value ([string]$_groupInfo.location) -ReplaceExistingInferred
      Use-InferredDeployValue -Key 'Environment' -Target $Environment `
        -Value (Get-TagValue -Tags $_groupInfo.tags -Name 'environment') `
        -AllowEmptyString -ReplaceExistingInferred
      Use-InferredDeployValue -Key 'Criticality' -Target $Criticality `
        -Value (Get-TagValue -Tags $_groupInfo.tags -Name 'criticality') `
        -AllowEmptyString -ReplaceExistingInferred
    }
  }
  catch {
    Write-Verbose "Could not derive location or tags from resource group '$ResourceGroupName': $_"
  }

  $_functionAppInfo = Get-DeployedFunctionAppInfo -ResourceGroup $ResourceGroupName -FunctionAppName $FunctionAppName
  if ($_functionAppInfo -and $_functionAppInfo.name) {
    Use-InferredDeployValue -Key 'FunctionAppName' -Target $FunctionAppNameTarget `
      -Value ([string]$_functionAppInfo.name) -AllowEmptyString -ReplaceExistingInferred
  }
}

function Get-GraphPermissionAssignmentRecommendation {
  param(
    [Parameter(Mandatory)][string]$EnvName,
    [Parameter(Mandatory)][string]$ResourceGroupName,
    [AllowEmptyString()][string]$FunctionAppName
  )

  $_signals = [System.Collections.Generic.List[string]]::new()
  $_storedManagedIdentityObjectId = Get-AzdEnvironmentStoredValue -EnvName $EnvName -Name 'managedIdentityObjectId'
  $_storedGraphAssignmentMarker = Get-AzdEnvironmentStoredValue -EnvName $EnvName -Name 'graphPermissionsAssignedManagedIdentityObjectId'
  $_storedFunctionAppName = Get-AzdEnvironmentStoredValue -EnvName $EnvName -Name 'functionAppName'
  $_effectiveFunctionAppName = if ($FunctionAppName) { $FunctionAppName } elseif ($_storedFunctionAppName) { $_storedFunctionAppName } else { '' }

  if ($_storedManagedIdentityObjectId) {
    $_signals.Add('azd environment already contains a Managed Identity object ID')
  }
  if ($_storedGraphAssignmentMarker) {
    $_signals.Add('azd environment already contains a Graph permission assignment marker')
  }

  if (Test-ResourceGroupPresence -ResourceGroupName $ResourceGroupName) {
    $_functionAppInfo = Get-DeployedFunctionAppInfo -ResourceGroup $ResourceGroupName -FunctionAppName $_effectiveFunctionAppName
    if ($_functionAppInfo) {
      $_signals.Add("Azure Function App '$($_functionAppInfo.name)' already exists in the target resource group")
    }
  }

  return [pscustomobject]@{
    ProbablyUpdate = $_signals.Count -gt 0
    Signals        = @($_signals)
  }
}

function Get-EasyAuthAppRegistrationUniqueName {
  param([Parameter(Mandatory)][string]$FunctionAppName)

  return "guest-sponsor-info-proxy-$FunctionAppName"
}

function Get-EasyAuthAppRegistrationIdentifierUri {
  param([Parameter(Mandatory)][string]$FunctionAppName)

  return "api://guest-sponsor-info-$FunctionAppName"
}

function Test-EasyAuthUniqueNameConflict {
  param([AllowEmptyString()][string]$Message)

  return -not [string]::IsNullOrWhiteSpace($Message) -and
  $Message -match '(?i)uniqueName[^\r\n]*already exists'
}

function Get-DeletedEasyAuthApplication {
  param([Parameter(Mandatory)][string]$FunctionAppName)

  $_identifierUri = Get-EasyAuthAppRegistrationIdentifierUri -FunctionAppName $FunctionAppName
  $_select = [System.Uri]::EscapeDataString('id,appId,displayName,deletedDateTime,identifierUris')
  $_filter = [System.Uri]::EscapeDataString("displayName eq '$($script:AppRegistrationDisplayName)'")
  $_url = "https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.application?`$select=$_select&`$filter=$_filter"

  try {
    $_rawJson = Invoke-AzureCliQuiet -Arguments @(
      'rest',
      '--method', 'GET',
      '--url', $_url,
      '-o', 'json'
    )

    $_jsonText = ($_rawJson -join "`n").Trim()
    if (-not $_jsonText -or $_jsonText -eq 'null') {
      return @()
    }

    $_payload = $_jsonText | ConvertFrom-Json
    return @($_payload.value | Where-Object {
        $_.identifierUris -and ($_.identifierUris -contains $_identifierUri)
      })
  }
  catch {
    Write-Verbose "Could not inspect deleted Entra applications for '$FunctionAppName': $_"
    return @()
  }
}

function Remove-DeletedEasyAuthApplication {
  [CmdletBinding(SupportsShouldProcess)]
  param([Parameter(Mandatory)][object[]]$DeletedApplications)

  foreach ($_deletedApp in $DeletedApplications) {
    if ($PSCmdlet.ShouldProcess("deleted Entra application '$($_deletedApp.id)'", 'Permanently delete from deleted applications')) {
      Invoke-AzureCli -Arguments @(
        'rest',
        '--method', 'DELETE',
        '--url', "https://graph.microsoft.com/v1.0/directory/deletedItems/$($_deletedApp.id)"
      )
    }
  }
}

function Resolve-DeletedEasyAuthApplicationConflict {
  param([Parameter(Mandatory)][string]$FunctionAppName)

  $_deletedApps = @(Get-DeletedEasyAuthApplication -FunctionAppName $FunctionAppName)
  if ($_deletedApps.Count -eq 0) {
    return $false
  }

  Write-Important @(
    'A soft-deleted Entra App Registration is blocking the EasyAuth bootstrap.'
    'The deleted application still reserves the deterministic uniqueName for this Function App.'
    ''
    'To continue, the deleted application can be permanently removed from Entra deleted applications.'
    'This is irreversible, but it only affects the deleted app object that is already in the recycle bin.'
  )

  Write-Host '  Matching deleted Entra application(s):' -ForegroundColor Cyan
  foreach ($_deletedApp in $_deletedApps) {
    $_deletedWhen = if ($_.deletedDateTime) { [string]$_.deletedDateTime } else { 'unknown' }
    $_deletedAppId = if ($_.appId) { [string]$_.appId } else { 'unknown' }
    Write-Host "    • appId: $_deletedAppId   deleted: $_deletedWhen" -ForegroundColor DarkGray
  }

  do {
    $_answer = (Read-Host '  Permanently delete the matching deleted application(s) now? [Y/n]').Trim()
    if (-not $_answer) { $_answer = 'y' }
    if ($_answer -notmatch '^(?i:y|yes|n|no)$') {
      Write-Host "  $_wrn Enter Y or N." -ForegroundColor Yellow
      $_answer = ''
    }
  } while (-not $_answer)

  if ($_answer -notmatch '^(?i:y|yes)$') {
    return $false
  }

  Write-Host ''
  Write-Host '  Removing matching deleted Entra application(s)...' -ForegroundColor Cyan
  Remove-DeletedEasyAuthApplication -DeletedApplications $_deletedApps
  return $true
}

function Get-WebPartClientId {
  param([Parameter(Mandatory)][string]$FunctionAppName)

  try {
    $_identifierUri = Get-EasyAuthAppRegistrationIdentifierUri -FunctionAppName $FunctionAppName
    $_clientId = (Invoke-AzureCliQuiet -Arguments @(
        'ad', 'app', 'show',
        '--id', $_identifierUri,
        '--query', 'appId',
        '-o', 'tsv'
      )).Trim()

    if ($_clientId -and $_clientId -ne 'null') {
      return $_clientId
    }
  }
  catch {
    Write-Verbose "Could not resolve EasyAuth App Registration client ID by identifier URI from Entra: $_"
  }

  try {
    $_appRegUniqueName = Get-EasyAuthAppRegistrationUniqueName -FunctionAppName $FunctionAppName
    $_clientId = (Invoke-AzureCliQuiet -Arguments @(
        'ad', 'app', 'list',
        '--filter', "uniqueName eq '$_appRegUniqueName'",
        '--query', '[0].appId',
        '-o', 'tsv'
      )).Trim()

    if ($_clientId -and $_clientId -ne 'null') {
      return $_clientId
    }
  }
  catch {
    Write-Verbose "Could not resolve EasyAuth App Registration client ID by uniqueName from Entra: $_"
  }

  return $null
}

function Get-SetupGraphPermissionsScriptReference {
  $_localScript = Join-Path $PSScriptRoot 'setup-graph-permissions.ps1'
  if (Test-Path $_localScript) {
    return $_localScript
  }

  $_releaseBaseUrl = 'https://github.com/workoho/spfx-guest-sponsor-info/releases'
  if ($InstallerVersion -and $InstallerVersion -ne 'latest') {
    if ($InstallerVersion -match '^v[0-9]+\.[0-9]+\.[0-9]+([.-][A-Za-z0-9.]+)?$') {
      return "$_releaseBaseUrl/download/$InstallerVersion/setup-graph-permissions.ps1"
    }

    $_installerRef = [System.Uri]::EscapeDataString($InstallerVersion)
    return "https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/$_installerRef/azure-function/infra/setup-graph-permissions.ps1"
  }

  return "$_releaseBaseUrl/latest/download/setup-graph-permissions.ps1"
}

function Initialize-ResourceGroup {
  [CmdletBinding(SupportsShouldProcess)]
  param(
    [Parameter(Mandatory)][string]$ResourceGroup,
    [Parameter(Mandatory)][string]$Location
  )

  try {
    $_exists = (Invoke-AzureCliQuiet -Arguments @(
        'group', 'exists',
        '--name', $ResourceGroup
      ) | Out-String).Trim()

    if ($_exists -eq 'true') {
      return
    }
  }
  catch {
    Write-Verbose "Could not determine whether resource group '$ResourceGroup' already exists: $_"
  }

  if (-not $PSCmdlet.ShouldProcess("resource group '$ResourceGroup'", "Create in location '$Location'")) {
    return
  }

  Write-Host ''
  Write-Host "  $_arr Creating resource group '$ResourceGroup'..." -ForegroundColor Cyan
  Invoke-AzureCli -Arguments @(
    'group', 'create',
    '--name', $ResourceGroup,
    '--location', $Location,
    '--output', 'none'
  )
}

function Resolve-EffectiveFunctionAppName {
  param(
    [Parameter(Mandatory)][string]$ResourceGroup,
    [AllowEmptyString()][string]$FunctionAppName
  )

  if (Test-FunctionAppNameLength -Value $FunctionAppName) {
    return $FunctionAppName
  }

  $_existingFunctionApp = Get-DeployedFunctionAppInfo -ResourceGroup $ResourceGroup
  if ($_existingFunctionApp -and $_existingFunctionApp.name) {
    return [string]$_existingFunctionApp.name
  }

  if ($_whatIf) {
    return 'gsi-preview-placeholder'
  }

  $_templatePath = Join-Path $PSScriptRoot 'resolve-function-app-name.bicep'
  $_deploymentName = "gsi-resolve-function-name-$([guid]::NewGuid().ToString('N').Substring(0, 8))"
  $_resolvedName = (Invoke-AzureCliQuiet -Arguments @(
      'deployment', 'group', 'create',
      '--name', $_deploymentName,
      '--resource-group', $ResourceGroup,
      '--template-file', $_templatePath,
      '--query', 'properties.outputs.effectiveFunctionAppName.value',
      '-o', 'tsv'
    ) | Out-String).Trim()

  if (-not (Test-FunctionAppNameLength -Value $_resolvedName)) {
    throw 'Could not resolve a deterministic Function App name for the deployment.'
  }

  return $_resolvedName
}

function Invoke-EntraAuthBootstrapProvision {
  param(
    [Parameter(Mandatory)][string]$ResourceGroup,
    [Parameter(Mandatory)][string]$FunctionAppName
  )

  $_templatePath = Join-Path $PSScriptRoot 'entra-auth.bicep'
  $_deploymentName = 'gsi-entra-auth-pre'

  try {
    $_outputsJson = Invoke-AzureCliQuiet -Arguments @(
      'deployment', 'group', 'create',
      '--name', $_deploymentName,
      '--resource-group', $ResourceGroup,
      '--template-file', $_templatePath,
      '--parameters', "functionAppName=$FunctionAppName",
      '--query', 'properties.outputs',
      '-o', 'json'
    )
  }
  catch {
    $_existingClientId = $null
    if (Test-EasyAuthUniqueNameConflict -Message $_.Exception.Message) {
      $_existingClientId = Get-WebPartClientId -FunctionAppName $FunctionAppName
      if ($_existingClientId) {
        return [pscustomobject]@{
          webPartClientId = [pscustomobject]@{ value = $_existingClientId }
        }
      }

      if (Resolve-DeletedEasyAuthApplicationConflict -FunctionAppName $FunctionAppName) {
        Write-Host ''
        Write-Host '  Retrying Entra App Registration bootstrap...' -ForegroundColor Cyan
        $_outputsJson = Invoke-AzureCliQuiet -Arguments @(
          'deployment', 'group', 'create',
          '--name', $_deploymentName,
          '--resource-group', $ResourceGroup,
          '--template-file', $_templatePath,
          '--parameters', "functionAppName=$FunctionAppName",
          '--query', 'properties.outputs',
          '-o', 'json'
        )
      }
      else {
        throw
      }
    }
    else {
      throw
    }
  }

  $_outputsText = ($_outputsJson -join "`n").Trim()
  if (-not $_outputsText -or $_outputsText -eq 'null') {
    throw 'The Entra auth deployment did not return any outputs.'
  }

  return $_outputsText | ConvertFrom-Json
}

function Invoke-GraphPermissionProvision {
  param(
    [Parameter(Mandatory)][string]$ResourceGroup,
    [Parameter(Mandatory)][string]$ManagedIdentityObjectId
  )

  $_templatePath = Join-Path $PSScriptRoot 'assign-graph-permissions.bicep'
  Invoke-AzureCliQuiet -Arguments @(
    'deployment', 'group', 'create',
    '--name', 'gsi-graph-permissions',
    '--resource-group', $ResourceGroup,
    '--template-file', $_templatePath,
    '--parameters', "managedIdentityObjectId=$ManagedIdentityObjectId",
    '--output', 'none'
  ) | Out-Null
}

function Test-GraphPermissionAssignmentCurrent {
  param(
    [AllowEmptyString()][string]$AssignedManagedIdentityObjectId,
    [AllowEmptyString()][string]$CurrentManagedIdentityObjectId
  )

  return -not [string]::IsNullOrWhiteSpace($CurrentManagedIdentityObjectId) -and
  $AssignedManagedIdentityObjectId -eq $CurrentManagedIdentityObjectId
}

function Set-GraphPermissionAssignmentMarker {
  [CmdletBinding(SupportsShouldProcess)]
  param(
    [Parameter(Mandatory)][string]$RepoRoot,
    [Parameter(Mandatory)][string]$ManagedIdentityObjectId
  )

  $markerName = 'graphPermissionsAssignedManagedIdentityObjectId'
  if (-not $PSCmdlet.ShouldProcess("azd environment marker '$markerName'", "Store Managed Identity Object ID '$ManagedIdentityObjectId'")) {
    return
  }

  Push-Location -Path $RepoRoot
  try {
    Invoke-Azd -Arguments @('env', 'set', $markerName, $ManagedIdentityObjectId)
  }
  finally {
    Pop-Location
  }

  Set-Item -Path "Env:$markerName" -Value $ManagedIdentityObjectId
}

function Restart-DeployedFunctionApp {
  [CmdletBinding(SupportsShouldProcess)]
  param(
    [Parameter(Mandatory)][string]$ResourceGroup,
    [Parameter(Mandatory)][string]$FunctionAppName
  )

  if (-not $PSCmdlet.ShouldProcess("Function App '$FunctionAppName' in resource group '$ResourceGroup'", 'Restart to activate Graph permissions')) {
    return
  }

  Write-Host ''
  Write-Host "  $_arr Restarting Function App '$FunctionAppName' to activate Graph permissions..." -ForegroundColor Cyan
  Invoke-AzureCli -Arguments @(
    'functionapp', 'restart',
    '--resource-group', $ResourceGroup,
    '--name', $FunctionAppName,
    '--output', 'none'
  )
}

function Get-InstallerSourceDisplayText {
  if ($InstallerVersion) {
    if ($InstallerVersion -eq 'latest' -or $InstallerVersion -match '^v[0-9]+\.[0-9]+\.[0-9]+([.-][A-Za-z0-9.]+)?$') {
      return "downloaded release package ($InstallerVersion)"
    }

    return "downloaded repository snapshot ($InstallerVersion)"
  }

  if (Test-Path (Join-Path $PSScriptRoot 'azure.yaml')) {
    return 'local extracted infra package'
  }

  return 'local repository checkout'
}

function Get-FunctionPackageDisplayText {
  if ($AppVersion -eq 'latest') {
    return 'latest release'
  }

  return $AppVersion
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

function Resolve-AzureAccountContext {
  $script:SubscriptionName = (Invoke-AzureCli -Arguments @('account', 'show', '--query', 'name', '-o', 'tsv')).Trim()
  $script:SubscriptionId = (Invoke-AzureCli -Arguments @('account', 'show', '--query', 'id', '-o', 'tsv')).Trim()
  $script:TenantId = (Invoke-AzureCli -Arguments @('account', 'show', '--query', 'tenantId', '-o', 'tsv')).Trim()
}

function Get-GraphTokenClientId {
  # Read the client application ID (appid claim) from the Microsoft Graph
  # access token the current Azure CLI session hands out. Returns an empty
  # string when no Graph token can be obtained or decoded — callers treat that
  # as "unknown" and continue rather than blocking the deployment.
  try {
    $_token = (Invoke-AzureCliQuiet -Arguments @(
        'account', 'get-access-token',
        '--resource', 'https://graph.microsoft.com',
        '--query', 'accessToken',
        '-o', 'tsv')).Trim()
  }
  catch {
    Write-Verbose "Could not obtain a Microsoft Graph access token: $_"
    return ''
  }

  $_segments = $_token -split '\.'
  if ($_segments.Count -lt 2) { return '' }

  # JWT payloads use base64url; restore standard base64 padding before decoding.
  $_payload = $_segments[1].Replace('-', '+').Replace('_', '/')
  while ($_payload.Length % 4) { $_payload += '=' }

  try {
    $_claims = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($_payload)) | ConvertFrom-Json
    return [string]$_claims.appid
  }
  catch {
    Write-Verbose "Could not decode the Microsoft Graph access token payload: $_"
    return ''
  }
}

function Confirm-AzureCliGraphClient {
  # Microsoft Graph authorizes delegated calls as the intersection of the
  # signed-in user's directory roles and the scopes granted to the client
  # application that issued the token. Azure Cloud Shell authenticates through
  # its own first-party application, whose Graph scopes do not cover creating
  # the EasyAuth App Registration: entra-auth.bicep then fails with
  # Authorization_RequestDenied even for a Global Administrator, and the
  # directory-role preflight above still reports every role as active.
  # Signing in with the Azure CLI itself mints a token that carries the
  # required scopes.
  #
  # Initialize-AzureAuthIsolation normally prevents this by giving every run
  # its own AZURE_CONFIG_DIR, so Cloud Shell reaches the regular az login path.
  # This stays as a safety net for caller-supplied config directories, and
  # re-authenticates without asking: a session that cannot create the App
  # Registration has no usable alternative to offer.
  $_clientId = Get-GraphTokenClientId
  if (-not $_clientId -or $_clientId -eq $script:AzureCliApplicationId) {
    return
  }

  Write-Important @(
    'The active Azure session was not issued to the Azure CLI.'
    ''
    "  Token client application : $_clientId"
    "  Azure CLI expects        : $($script:AzureCliApplicationId)"
    ''
    'Microsoft Graph grants a delegated call only what both your directory roles'
    'and that application allow. Creating the Entra App Registration therefore'
    'fails with Authorization_RequestDenied, no matter which roles are active.'
    ''
    'Signing in with the Azure CLI now. Your other Azure sessions stay intact.'
  )

  # Device code, not browser: this situation occurs in hosted shells that have
  # no browser of their own.
  $_loginArgs = @('login', '--use-device-code')
  if ($script:TenantId) {
    $_loginArgs += @('--tenant', $script:TenantId)
  }

  Write-Host "  $_arr Starting az login with device code..." -ForegroundColor Cyan
  & $script:AzPath @_loginArgs | Out-Null
  if ($LASTEXITCODE -ne 0) {
    throw "Azure CLI sign-in failed (exit code $LASTEXITCODE). Run 'az logout' followed by 'az login --use-device-code', then re-run this script."
  }

  $_clientId = Get-GraphTokenClientId
  if ($_clientId -and $_clientId -ne $script:AzureCliApplicationId) {
    Write-Host "  $_wrn The Graph token is still issued to application $_clientId." -ForegroundColor Yellow
    Write-Host '       Run ''az logout'', then ''az login --use-device-code'', and re-run this script.' -ForegroundColor DarkGray
    return
  }

  Write-Host "  $_chk Azure CLI sign-in completed; Graph calls now use the Azure CLI application." -ForegroundColor Green
}

function Connect-AzureCliIfNeeded {
  if (-not (Test-AzureCliAccountAvailable)) {
    $_loginPreference = Get-AzureCliLoginPreference
    Write-Host "  $_arr No active Azure CLI session found for this PowerShell console. $($_loginPreference.StartText)" -ForegroundColor Cyan

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
    if ($_loginPreference.Mode -eq 'device-code') {
      $_loginArgs += '--use-device-code'
    }
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

  Resolve-AzureAccountContext

  if ($AzureTenantId) {
    $_expectedTenantId = $AzureTenantId.Trim()
    if ($_expectedTenantId -and $script:TenantId -and $_expectedTenantId -ne $script:TenantId) {
      throw "Azure CLI signed in to tenant '$($script:TenantId)', but -AzureTenantId requested '$($_expectedTenantId)'. Re-run and pick the correct tenant before continuing."
    }
  }

  # A reused ambient session — Azure Cloud Shell in particular — can hold a
  # token issued to a different client application, which Microsoft Graph
  # authorizes differently from an Azure CLI token.
  Confirm-AzureCliGraphClient
  # A re-login can leave a different subscription selected.
  Resolve-AzureAccountContext

  if ($script:AzdUsesAzureCliAuth) {
    Write-Host "  $_chk Azure tenant        : $($script:TenantId)" -ForegroundColor Green
    Write-Host '       azd will reuse this Azure CLI tenant for all deployment commands.' -ForegroundColor DarkGray
  }

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

  return $script:ExplicitDeployParameters.Contains($Name) -or
  $script:CachedDeployParameters.Contains($Name) -or
  $script:InferredDeployParameters.Contains($Name)
}

function Test-DeployParameterPromptRequired {
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
  $script:CanRepairGraphPermissionsInThisRun = $false
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
    $deploymentRoleNames = @('Owner', 'Contributor')
    $roleAssignmentRoleNames = @('Owner', 'User Access Administrator', 'Role Based Access Control Administrator')
    $roleAssignmentRoleLabel = 'Owner/User Access Administrator/Role Based Access Control Administrator'
    $hasSubscriptionContributor = ($subscriptionRoles | Where-Object { $_ -in $deploymentRoleNames }).Count -gt 0
    $hasContributor = ($allAzureRoles | Where-Object { $_ -in $deploymentRoleNames }).Count -gt 0
    $hasRoleAssignment = ($allAzureRoles | Where-Object { $_ -in $roleAssignmentRoleNames }).Count -gt 0

    if ($hasContributor) {
      if ($hasSubscriptionContributor) {
        Write-Host "  $_chk Azure deployment role: Contributor/Owner visible." -ForegroundColor Green
      }
      else {
        Write-Host "  $_chk Azure deployment role: Contributor/Owner visible on the target resource group." -ForegroundColor Green
        Write-Host '       Later deployments usually work with this scope.' -ForegroundColor DarkGray
        Write-Host '       A bootstrap run can still need Contributor or Owner on the subscription' -ForegroundColor DarkGray
        Write-Host '       for provider registration.' -ForegroundColor DarkGray
      }
    }
    else {
      if (-not $resourceGroupExists) {
        Write-Host "  $_wrn Azure deployment role missing: Contributor or Owner at subscription scope." -ForegroundColor Yellow
        $missing.Add("Azure Contributor or Owner on $subscriptionScope when this run still needs provider registration or initial resource group creation")
      }
      else {
        Write-Host "  $_wrn Azure deployment role missing: Contributor or Owner." -ForegroundColor Yellow
        $missing.Add("Azure Contributor or Owner on $resourceGroupScope (or inherited from the subscription)")
      }
    }

    if ($hasRoleAssignment) {
      Write-Host "  $_chk Azure role assignment permission: $roleAssignmentRoleLabel visible." -ForegroundColor Green
    }
    else {
      if (-not $resourceGroupExists) {
        Write-Host "  $_wrn Azure role assignment permission missing: Owner, User Access Administrator, or Role Based Access Control Administrator at subscription scope." -ForegroundColor Yellow
        $missing.Add("Azure Owner, User Access Administrator, or Role Based Access Control Administrator on $subscriptionScope when this run still needs initial resource group creation (or inherited from a higher scope)")
      }
      else {
        Write-Host "  $_wrn Azure role assignment permission missing: Owner, User Access Administrator, or Role Based Access Control Administrator." -ForegroundColor Yellow
        $missing.Add("Azure Owner, User Access Administrator, or Role Based Access Control Administrator on $resourceGroupScope (or inherited from the subscription)")
      }
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
    $script:CanRepairGraphPermissionsInThisRun = -not $SkipGraphRoleAssignments -and $hasPrivilegedRoleAdmin

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
    [Parameter(Mandatory)][string]$TenantId,
    [Parameter(Mandatory)][string]$SharePointTenant,
    [Parameter(Mandatory)][AllowEmptyString()][string]$FunctionAppName,
    [Parameter(Mandatory)][string]$AppClientId,
    [Parameter(Mandatory)][AllowEmptyString()][string]$Environment,
    [Parameter(Mandatory)][AllowEmptyString()][string]$Criticality,
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

  function Sync-AzdOptionalEnvValue {
    param(
      [Parameter(Mandatory)][string]$Name,
      [AllowEmptyString()][string]$Value
    )

    if (-not [string]::IsNullOrWhiteSpace($Value)) {
      Invoke-Azd -Arguments @('env', 'set', $Name, $Value)
      Set-Item -Path "Env:$Name" -Value $Value
      return
    }

    $_envFile = Join-Path $repoRoot ".azure/$EnvName/.env"
    if (Test-Path $_envFile) {
      $_keepPattern = '^' + [regex]::Escape($Name) + '='
      $_filteredLines = @(
        Get-Content -Path $_envFile -ErrorAction Stop |
        Where-Object { $_ -notmatch $_keepPattern }
      )
      Set-Content -Path $_envFile -Value $_filteredLines
    }

    Remove-Item -Path "Env:$Name" -ErrorAction SilentlyContinue
  }

  Write-Host ''
  Write-Host '  azd provision' -ForegroundColor Cyan
  Write-Host $_sep -ForegroundColor DarkGray

  # Tell azd to reuse the Azure CLI token so the user is not prompted to
  # log in a second time via a separate azd browser window.
  Enable-AzdAzureCliAuth
  # Enable deployment stacks in the active azd config so the wizard does not
  # silently fall back to non-stack provisioning when it uses an isolated .azd.
  Enable-AzdDeploymentStackSupport

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
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_TENANT_ID', $TenantId)
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_RESOURCE_GROUP', $ResourceGroup)
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_LOCATION', $Location)
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_SHAREPOINT_TENANT_NAME', $SharePointTenant)
    Sync-AzdOptionalEnvValue -Name 'AZURE_FUNCTION_APP_NAME' -Value $FunctionAppName
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_WEB_PART_CLIENT_ID', $AppClientId)
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_TAG_ENVIRONMENT', $Environment)
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_TAG_CRITICALITY', $Criticality)
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_APP_VERSION', $Version)
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_ENABLE_MONITORING', $Monitoring.ToString().ToLower())
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_ENABLE_FAILURE_ANOMALIES_ALERT', $FailureAlert.ToString().ToLower())
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_ALWAYS_READY_INSTANCES', $AlwaysReadyInstances.ToString())
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_MAXIMUM_FLEX_INSTANCES', $FlexInstances.ToString())
    Invoke-Azd -Arguments @('env', 'set', 'AZURE_INSTANCE_MEMORY_MB', $InstanceMemoryMB.ToString())
    # Store deployment params in azd env so pre-provision can derive the same
    # conditional provider checks and post-provision can reuse the choices.
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
  $env:AZURE_TENANT_ID = $TenantId
  $env:AZURE_LOCATION = $Location
  $env:AZURE_RESOURCE_GROUP = $ResourceGroup
  $env:AZURE_SHAREPOINT_TENANT_NAME = $SharePointTenant
  if ($FunctionAppName) {
    $env:AZURE_FUNCTION_APP_NAME = $FunctionAppName
  }
  else {
    Remove-Item -Path 'Env:AZURE_FUNCTION_APP_NAME' -ErrorAction SilentlyContinue
  }
  $env:AZURE_WEB_PART_CLIENT_ID = $AppClientId
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
    Write-Host '       Bicep deploys the Azure-only infrastructure, configures EasyAuth with the' -ForegroundColor DarkGray
    Write-Host '       pre-created Entra App Registration, and publishes the function package via' -ForegroundColor DarkGray
    Write-Host '       the native Flex OneDeploy path. Graph permissions are handled after azd.' -ForegroundColor DarkGray
  }
  Write-Host ''

  Push-Location -Path $repoRoot
  $_previousExternalPostProvisionEntra = $env:GSI_EXTERNAL_POST_PROVISION_ENTRA
  $env:GSI_EXTERNAL_POST_PROVISION_ENTRA = 'true'
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
    if ([string]::IsNullOrWhiteSpace($_previousExternalPostProvisionEntra)) {
      Remove-Item -Path 'Env:GSI_EXTERNAL_POST_PROVISION_ENTRA' -ErrorAction SilentlyContinue
    }
    else {
      $env:GSI_EXTERNAL_POST_PROVISION_ENTRA = $_previousExternalPostProvisionEntra
    }
    Pop-Location
  }
}

#region Main
try {
  Write-Host ''
  Write-Host "  Guest Sponsor Info  $(if ($_u) { [string][char]0x00B7 } else { '|' })  Azure Deployment" -ForegroundColor DarkCyan
  Write-Host $_sep -ForegroundColor DarkGray
  Write-Host "  Installer source   : $(Get-InstallerSourceDisplayText)" -ForegroundColor DarkGray
  Write-Host "  Function package   : $(Get-FunctionPackageDisplayText)" -ForegroundColor DarkGray
  Write-Host ''
  Show-PreflightOverview

  # ── Install tools and connect to Azure ────────────────────────────────────
  Initialize-AzureAuthIsolation
  Install-AzureCliIfNeeded
  Install-AzureCliBicepIfNeeded
  Install-AzdIfNeeded
  Enable-AzdAzureCliAuth
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
    $_usedAzdStoredDefaults = $false
    $_usedLiveAzureDefaults = $false

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

    $_inferredDefaultsBefore = $script:InferredDeployParameters.Count
    Use-AzdEnvironmentStoredDeploymentProfile `
      -EnvName $AzdEnvironmentName `
      -ResourceGroupName ([ref]$ResourceGroupName) `
      -AzureLocation ([ref]$AzureLocation) `
      -TenantName ([ref]$TenantName) `
      -FunctionAppName ([ref]$FunctionAppName) `
      -DeployAzureMaps ([ref]$DeployAzureMaps) `
      -Environment ([ref]$Environment) `
      -Criticality ([ref]$Criticality) `
      -EnableMonitoring ([ref]$EnableMonitoring) `
      -EnableFailureAnomaliesAlert ([ref]$EnableFailureAnomaliesAlert) `
      -AlwaysReadyInstances ([ref]$AlwaysReadyInstances) `
      -MaximumFlexInstances ([ref]$MaximumFlexInstances) `
      -InstanceMemoryMB ([ref]$InstanceMemoryMB)
    $_usedAzdStoredDefaults = $script:InferredDeployParameters.Count -gt $_inferredDefaultsBefore

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

    $_inferredDefaultsBefore = $script:InferredDeployParameters.Count
    Use-LiveDeploymentProfile `
      -ResourceGroupName $ResourceGroupName `
      -FunctionAppName $FunctionAppName `
      -AzureLocation ([ref]$AzureLocation) `
      -Environment ([ref]$Environment) `
      -Criticality ([ref]$Criticality) `
      -FunctionAppNameTarget ([ref]$FunctionAppName)
    $_usedLiveAzureDefaults = $script:InferredDeployParameters.Count -gt $_inferredDefaultsBefore
    if ($_usedAzdStoredDefaults -or $_usedLiveAzureDefaults) {
      Write-Host ''
      Write-Host '  Existing deployment settings were reused as safe defaults to avoid accidental drift.' -ForegroundColor DarkGray
      if ($_usedAzdStoredDefaults) {
        Write-Host "       Source: .azure/$AzdEnvironmentName/.env" -ForegroundColor DarkGray
      }
      if ($_usedLiveAzureDefaults) {
        Write-Host '       Source: current Azure resource group / Function App' -ForegroundColor DarkGray
      }
      Write-Host '       Use re-configure later if you intentionally want to change them.' -ForegroundColor DarkGray
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
    if (Test-DeployParameterPromptRequired -Name 'Environment') {
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
    if (Test-DeployParameterPromptRequired -Name 'Criticality') {
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

    if (Test-DeployParameterPromptRequired -Name 'FunctionAppName') {
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

    # ── Azure Maps ────────────────────────────────────────────────────────────
    if (Test-DeployParameterPromptRequired -Name 'DeployAzureMaps') {
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
    if (Test-DeployParameterPromptRequired -Name 'AppVersion') {
      $_appVersionDefault = Get-PromptDefaultValue -CurrentValue $AppVersion -FallbackValue 'latest'
      Write-Host ''
      Write-Host '  Function Package Override (Advanced)' -ForegroundColor Cyan
      Write-Host $_sep -ForegroundColor DarkGray
      Write-Host '  Most deployments should keep the default value.'
      Write-Host '  Set this only when you intentionally want a different published'
      Write-Host '  Function package version than the installer default.'
      Write-Host '  Use "latest" to always pull the most recent published release.'
      Write-Link -Url 'https://github.com/workoho/spfx-guest-sponsor-info/releases' `
        -Text "GitHub releases $_arr workoho/spfx-guest-sponsor-info"
      Write-Host ''
      $AppVersion = (Read-Host "  Function package override [$_appVersionDefault]").Trim()
      if ($AppVersion -eq '') { $AppVersion = $_appVersionDefault }
      Write-Host ''
      $_promptsShown = $true
    }

    # ── Monitoring Stack ──────────────────────────────────────────────────────
    if (Test-DeployParameterPromptRequired -Name 'EnableMonitoring') {
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
    if ($EnableMonitoring -and (Test-DeployParameterPromptRequired -Name 'EnableFailureAnomaliesAlert')) {
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
    if (Test-DeployParameterPromptRequired -Name 'AlwaysReadyInstances') {
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
    if (Test-DeployParameterPromptRequired -Name 'MaximumFlexInstances') {
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
    if (Test-DeployParameterPromptRequired -Name 'InstanceMemoryMB') {
      $_instanceMemoryDefault = $InstanceMemoryMB.ToString()
      Write-Host ''
      Write-Host '  Flex Instance Memory' -ForegroundColor Cyan
      Write-Host $_sep -ForegroundColor DarkGray
      Write-Host '  Memory size per Flex Consumption instance.'
      Write-Host '  Supported in this template: 512 or 2048 MB. Recommended default: 512 MB.'
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
    $_graphPermissionRecommendation = Get-GraphPermissionAssignmentRecommendation `
      -EnvName $AzdEnvironmentName `
      -ResourceGroupName $ResourceGroupName `
      -FunctionAppName $FunctionAppName
    $_graphPermissionsDefaultToSkip = $_graphPermissionRecommendation.ProbablyUpdate -and
    -not $script:ExplicitDeployParameters.Contains('SkipGraphRoleAssignments')
    if ($_graphPermissionsDefaultToSkip) {
      $SkipGraphRoleAssignments = $true
    }

    if (Test-DeployParameterPromptRequired -Name 'SkipGraphRoleAssignments') {
      $_graphPermissionsDefault = if ($SkipGraphRoleAssignments) { '2' } else { '1' }
      Write-Host ''
      Write-Host '  Graph Permission Assignment' -ForegroundColor Cyan
      Write-Host $_sep -ForegroundColor DarkGray
      Write-Host '  Bicep assigns Microsoft Graph app roles to the Managed Identity during'
      Write-Host '  deployment. This requires Privileged Role Administrator in Entra ID.'
      Write-Host ''
      if ($_graphPermissionRecommendation.ProbablyUpdate) {
        Write-Host '  Existing deployment signals were detected. This looks like an update run.' -ForegroundColor DarkGray
        foreach ($_signal in $_graphPermissionRecommendation.Signals) {
          Write-Host "       - $_signal" -ForegroundColor DarkGray
        }
        Write-Host '  Default is to leave Microsoft Graph permissions unchanged for this run.' -ForegroundColor DarkGray
        Write-Host '  Choose [1] only when you intentionally want to assign or repair them now.' -ForegroundColor DarkGray
        Write-Host ''
        Write-Host '    [1]  Assign / repair now — requires Privileged Role Administrator'
        Write-Host '    [2]  Skip for this run (default for updates) — do not change Graph permissions'
      }
      else {
        Write-Host '    [1]  Assign now (default) — requires Privileged Role Administrator'
        Write-Host '    [2]  Defer — run setup-graph-permissions.ps1 after deployment'
        Write-Host '         (useful when a separate PAW or account holds that Entra role)'
      }
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
        if ($_graphPermissionRecommendation.ProbablyUpdate) {
          Write-Host '  Automatic Graph permission changes are skipped for this update run.' -ForegroundColor DarkGray
        }
        else {
          Write-Host '  Graph role assignments deferred to setup-graph-permissions.ps1.' -ForegroundColor DarkGray
        }
      }
      Write-Host ''
      $_promptsShown = $true
    }
    elseif ($_graphPermissionsDefaultToSkip) {
      Write-Host ''
      Write-Host '  Existing deployment detected — automatic Graph permission changes are skipped by default for this run.' -ForegroundColor DarkGray
    }

    # ── Required role guidance ────────────────────────────────────────────────
    Write-Hint @(
      'Recommended Azure RBAC scope for a bootstrap run: subscription level.'
      '  Contributor covers provider registration and, if desired, initial resource group creation.'
      '  Owner, User Access Administrator, or Role Based Access Control Administrator is also needed'
      '  for role assignments.'
      'Later deployments usually work with target-resource-group scope once the providers are'
      '  already registered and the resource group exists.'
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
    # Show the summary and ask "Proceed / re-configure / abort?" when:
    #   a) all parameters came from the command line / session cache (no interactive
    #      prompts were shown), OR
    #   b) the preflight detected at least one missing required permission — in that
    #      case the operator must explicitly confirm they want to proceed despite the
    #      known gap, rather than having the deployment fail silently mid-run.
    # Skipped when -Confirm:$false or -WhatIf was passed.
    if ((-not $_promptsShown -or $script:LastPreflightMissingPermissions) -and
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
      Write-Host "  Always-ready        : $AlwaysReadyInstances"
      Write-Host "  Max flex instances  : $MaximumFlexInstances"
      Write-Host "  Instance memory MB  : $InstanceMemoryMB"
      Write-Host "  Azure Maps          : $DeployAzureMaps"
      Write-Host "  Monitoring          : $EnableMonitoring"
      Write-Host "  Function package    : $AppVersion"
      if ($SkipGraphRoleAssignments) {
        if ($_graphPermissionRecommendation.ProbablyUpdate) {
          Write-Host '  Graph roles         : unchanged by default for detected update'
        }
        else {
          Write-Host '  Graph roles         : deferred to setup-graph-permissions.ps1'
        }
      }
      else {
        Write-Host '  Graph roles         : app registration before Azure, role assignment after Azure'
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
  if (-not $_whatIf) {
    Initialize-ResourceGroup -ResourceGroup $ResourceGroupName -Location $AzureLocation
  }

  $_effectiveFunctionAppName = Resolve-EffectiveFunctionAppName `
    -ResourceGroup $ResourceGroupName `
    -FunctionAppName $FunctionAppName
  $FunctionAppName = $_effectiveFunctionAppName

  $_provisionAppClientId = $null
  $_graphPermissionsStatus = if ($SkipGraphRoleAssignments) { 'managed separately in this mode' } else { 'ready to use' }
  $_functionAppRestartStatus = if ($SkipGraphRoleAssignments) { 'not run automatically in this mode' } else { 'completed automatically' }

  if ($_whatIf) {
    $_provisionAppClientId = Get-WebPartClientId -FunctionAppName $FunctionAppName
    if (-not $_provisionAppClientId) {
      $_provisionAppClientId = '00000000-0000-0000-0000-000000000000'
    }
  }
  else {
    $_provisionAppClientId = Get-WebPartClientId -FunctionAppName $FunctionAppName
    if ($_provisionAppClientId) {
      Write-Host ''
      Write-Host '  Reusing existing Entra App Registration...' -ForegroundColor DarkGray
    }
    else {
      Write-Host ''
      Write-Host '  Preparing Entra App Registration...' -ForegroundColor Cyan
      $_entraAppOutputs = Invoke-EntraAuthBootstrapProvision `
        -ResourceGroup $ResourceGroupName `
        -FunctionAppName $FunctionAppName
      $_provisionAppClientId = [string]$_entraAppOutputs.webPartClientId.value
      if (-not $_provisionAppClientId) {
        throw 'The Entra App Registration deployment did not return a client ID.'
      }
    }
  }

  Invoke-AzdProvision `
    -EnvName $AzdEnvironmentName `
    -ResourceGroup $ResourceGroupName `
    -Location $AzureLocation `
    -TenantId $script:TenantId `
    -SharePointTenant $TenantName `
    -FunctionAppName $FunctionAppName `
    -AppClientId $_provisionAppClientId `
    -Environment $Environment `
    -Criticality $Criticality `
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
  $_graphPermissionsAssignedMiOid = $null
  $_azdFunctionAppName = if ($FunctionAppName) { $FunctionAppName } else { $null }
  $_outputResourceGroup = if ($ResourceGroupName) { $ResourceGroupName } elseif ($env:AZURE_RESOURCE_GROUP) { $env:AZURE_RESOURCE_GROUP } else { $null }
  if (-not $_whatIf) {
    try {
      $_azdEnvVals = azd env get-values 2>$null
      foreach ($_azdLine in $_azdEnvVals) {
        if ($_azdLine -match '^functionAppName="?([^\"]+)"?') { $_azdFunctionAppName = $Matches[1] }
        elseif ($_azdLine -match '^AZURE_FUNCTION_APP_NAME="?([^\"]+)"?') { $_azdFunctionAppName = $Matches[1] }
        if ($_azdLine -match '^functionAppUrl="?([^"]+)"?') { $_azdFunctionBaseUrl = $Matches[1] }
        elseif ($_azdLine -match '^sponsorApiEndpointUrl="?([^"]+)"?') { $_azdFunctionBaseUrl = $Matches[1] -replace '/api/getGuestSponsors$' }
        elseif ($_azdLine -match '^sponsorApiUrl="?([^"]+)"?') { $_azdFunctionBaseUrl = $Matches[1] -replace '/api/getGuestSponsors$' }
        elseif ($_azdLine -match '^webPartClientId="?([^"]+)"?') { $_azdWebPartClientId = $Matches[1] }
        elseif ($_azdLine -match '^AZURE_WEB_PART_CLIENT_ID="?([^\"]+)"?') { $_azdWebPartClientId = $Matches[1] }
        elseif ($_azdLine -match '^managedIdentityObjectId="?([^"]+)"?') { $_azdMiOid = $Matches[1] }
        elseif ($_azdLine -match '^graphPermissionsAssignedManagedIdentityObjectId="?([^\"]+)"?') { $_graphPermissionsAssignedMiOid = $Matches[1] }
      }
    }
    catch {
      # Non-fatal — values can be found in the Azure portal.
      Write-Verbose "Could not read azd env values after provision: $_"
    }

    $_functionAppMetadata = $null
    if ($_outputResourceGroup) {
      $_functionAppMetadata = Get-DeployedFunctionAppInfo -ResourceGroup $_outputResourceGroup -FunctionAppName $_azdFunctionAppName
      if ($_functionAppMetadata) {
        if (-not $_azdFunctionAppName -and $_functionAppMetadata.name) {
          $_azdFunctionAppName = [string]$_functionAppMetadata.name
        }
        if (-not $_azdFunctionBaseUrl -and $_functionAppMetadata.defaultHostName) {
          $_azdFunctionBaseUrl = "https://$($_functionAppMetadata.defaultHostName)"
        }
        if (-not $_azdMiOid -and $_functionAppMetadata.principalId) {
          $_azdMiOid = [string]$_functionAppMetadata.principalId
        }
      }
    }

    if (-not $_azdFunctionBaseUrl -and $_azdFunctionAppName -and $_outputResourceGroup) {
      try {
        $_defaultHostName = (Invoke-AzureCliQuiet -Arguments @(
            'functionapp', 'show',
            '--resource-group', $_outputResourceGroup,
            '--name', $_azdFunctionAppName,
            '--query', 'defaultHostName',
            '-o', 'tsv'
          ) | Out-String).Trim()

        if ($_defaultHostName -and $_defaultHostName -ne 'null') {
          $_azdFunctionBaseUrl = "https://$_defaultHostName"
        }
      }
      catch {
        Write-Verbose "Could not resolve Function App base URL from Azure CLI after provision: $_"
      }
    }

    if (-not $_azdFunctionBaseUrl -and $_azdFunctionAppName) {
      $_azdFunctionBaseUrl = "https://$($_azdFunctionAppName).azurewebsites.net"
    }

    if (-not $_azdWebPartClientId -and $_azdFunctionAppName) {
      $_azdWebPartClientId = Get-WebPartClientId -FunctionAppName $_azdFunctionAppName
    }

    # Cache the Managed Identity Object ID so setup-graph-permissions.ps1
    # can skip its own prompt when run in the same PowerShell session.
    if ($_azdMiOid) {
      $Global:GsiSetup_ManagedIdentityObjectId = $_azdMiOid
      Write-Host "  $_chk Managed Identity Object ID cached: $_azdMiOid" -ForegroundColor Green
    }
    $Global:GsiSetup_TenantId = $script:TenantId

    if (-not $SkipGraphRoleAssignments) {
      if (-not $_azdMiOid) {
        throw 'Managed Identity Object ID was not available after azd provision; Graph role assignments cannot continue.'
      }

      $_graphPermissionsAlreadyMarked = Test-GraphPermissionAssignmentCurrent `
        -AssignedManagedIdentityObjectId $_graphPermissionsAssignedMiOid `
        -CurrentManagedIdentityObjectId $_azdMiOid

      if ($_graphPermissionsAlreadyMarked -and -not $script:CanRepairGraphPermissionsInThisRun) {
        $_graphPermissionsStatus = 'already assigned for the current managed identity'
        $_functionAppRestartStatus = 'not needed (managed identity unchanged)'
        Write-Host ''
        Write-Host "  $_chk Microsoft Graph permissions already match the current Managed Identity." -ForegroundColor Green
      }
      else {
        Write-Host ''
        if ($_graphPermissionsAlreadyMarked -and $script:CanRepairGraphPermissionsInThisRun) {
          Write-Host '  Re-applying Microsoft Graph permissions to repair possible manual drift...' -ForegroundColor Cyan
        }
        else {
          Write-Host '  Assigning Microsoft Graph permissions...' -ForegroundColor Cyan
        }
        Invoke-GraphPermissionProvision `
          -ResourceGroup $ResourceGroupName `
          -ManagedIdentityObjectId $_azdMiOid
        Set-GraphPermissionAssignmentMarker `
          -RepoRoot (Get-RepoRoot) `
          -ManagedIdentityObjectId $_azdMiOid
        $_graphPermissionsStatus = if ($_graphPermissionsAlreadyMarked) {
          're-applied during this run to repair possible drift'
        }
        else {
          'assigned during this run'
        }
        if (-not $_azdWebPartClientId) {
          $_azdWebPartClientId = $_provisionAppClientId
        }
        if (-not $_azdFunctionAppName) {
          throw 'Function App name was not available after azd provision; automatic restart to activate Graph permissions cannot continue.'
        }
        Restart-DeployedFunctionApp -ResourceGroup $ResourceGroupName -FunctionAppName $_azdFunctionAppName
        $_functionAppRestartStatus = 'completed automatically'
      }
    }
  }

  # ── NEXT STEPS ────────────────────────────────────────────────────────────
  $_ns = [System.Collections.Generic.List[string]]::new()
  if ($_whatIf) {
    $_ns.Add('Preview summary:')
    $_ns.Add('')
    $_ns.Add(('  {0,-28}: {1}' -f 'Preview run', 'completed'))
    $_ns.Add(('  {0,-28}: {1}' -f 'Azure resources', 'unchanged'))
    $_ns.Add(('  {0,-28}: {1}' -f 'Deployment settings', 'updated for preview'))
    $_ns.Add('')
    $_ns.Add('Expected web part configuration after a real deployment:')
  }
  else {
    $_ns.Add('Deployment summary:')
    $_ns.Add('')
    $_ns.Add(('  {0,-28}: {1}' -f 'Application registration', 'created or updated'))
    $_ns.Add(('  {0,-28}: {1}' -f 'Azure resources', 'created or updated'))
    $_ns.Add(('  {0,-28}: {1}' -f 'Microsoft Graph permissions', $_graphPermissionsStatus))
    $_ns.Add(('  {0,-28}: {1}' -f 'Function app restart', $_functionAppRestartStatus))
    $_ns.Add('')
    $_ns.Add('Web part configuration (SharePoint property pane → Guest Sponsor API):')
  }
  if ($_whatIf) {
    $_ns.Add(('  {0,-28}: {1}' -f 'Guest Sponsor API Base URL', 'available after deployment'))
    $_ns.Add(('  {0,-28}: {1}' -f 'Guest Sponsor API Client ID', 'available after deployment'))
  }
  elseif ($_azdFunctionBaseUrl) {
    $_ns.Add(('  {0,-28}: {1}' -f 'Guest Sponsor API Base URL', $_azdFunctionBaseUrl))
  }
  else {
    $_ns.Add(('  {0,-28}: {1}' -f 'Guest Sponsor API Base URL', 'could not be resolved automatically'))
  }
  if (-not $_whatIf) {
    if ($_azdWebPartClientId) {
      $_ns.Add(('  {0,-28}: {1}' -f 'Guest Sponsor API Client ID', $_azdWebPartClientId))
    }
    else {
      $_ns.Add(('  {0,-28}: {1}' -f 'Guest Sponsor API Client ID', 'check Microsoft Entra app registrations if needed'))
    }
    if ($SkipGraphRoleAssignments) {
      $_ns.Add('')
      $_ns.Add('Microsoft Graph permissions: separate management path')
      $_miDisplay = if ($_azdMiOid) { $_azdMiOid } else { 'see Azure portal -> Function App -> Identity' }
      $_ns.Add(('  {0,-28}: {1}' -f 'Managed identity object ID', $_miDisplay))
      $_ns.Add(('  {0,-28}: {1}' -f 'TenantId', $script:TenantId))
      $_ns.Add(('  {0,-28}: {1}' -f 'If needed, run this script', (Get-SetupGraphPermissionsScriptReference)))
      $_ns.Add(('  {0,-28}: {1}' -f 'When to run it', 'if the web part shows permission errors'))
    }
    $_ns.Add('')
    $_ns.Add('Finish in SharePoint:')
    $_ns.Add('  1. Open the guest landing page in edit mode.')
    $_ns.Add('  2. Open the Guest Sponsor Info web part property pane.')
    $_ns.Add('  3. Paste the Base URL and Application (client) ID under Guest Sponsor API.')
    $_ns.Add('  4. Save or publish the page.')
    $_ns.Add('')
    $_ns.Add('Retry:')
    $_ns.Add('  Re-run the same command. The saved deployment settings and Azure resources')
    $_ns.Add('  are reused safely.')
  }
  Write-NextStep @($_ns)
}
catch {
  Write-Failure @(
    'Failure summary:'
    ''
    ('  {0,-28}: {1}' -f 'Deployment status', 'not completed')
    ('  {0,-28}: {1}' -f 'Retry', 'fix the error above, then run the same command again')
    ('  {0,-28}: {1}' -f 'Deployment settings', 'reused on the next run')
    ('  {0,-28}: {1}' -f 'Azure resources', 'existing resources are reused, updated, or skipped')
    ''
    'Do not delete partially created Azure resources unless the error message'
    'explicitly tells you to do so.'
  )
  throw
}
#endregion
