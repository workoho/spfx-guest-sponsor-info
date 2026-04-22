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
  and the repository's KQL alert resources. Defaults to false.

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

#Requires -Version 5.1
[CmdletBinding()]
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
  [bool]$EnableMonitoring = $false,
  [bool]$EnableFailureAnomaliesAlert = $false,
  [int]$MaximumFlexInstances = 10
)

$ErrorActionPreference = 'Stop'

$script:RepoRawBaseUrl = 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main'
$script:AppRegistrationDisplayName = 'Guest Sponsor Info - SharePoint Web Part Auth'
$script:TempPaths = [System.Collections.Generic.List[string]]::new()
$script:StagedRepoRoot = $null
$script:AzPath = $null
$script:AzdPath = $null
$script:SubscriptionName = ''
$script:SubscriptionId = ''
$script:TenantId = ''

if ([Console]::OutputEncoding.CodePage -ne 65001) {
  try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8
  }
  catch { $null = $_ }
}

function Write-Section {
  param([Parameter(Mandatory)][string]$Title)

  Write-Host ''
  Write-Host ('=' * 70) -ForegroundColor Cyan
  Write-Host "  $Title" -ForegroundColor Cyan
  Write-Host ('=' * 70) -ForegroundColor Cyan
}

function Write-InfoLine {
  param([Parameter(Mandatory)][string]$Message)

  Write-Host "  -> $Message" -ForegroundColor Cyan
}

function Write-SuccessLine {
  param([Parameter(Mandatory)][string]$Message)

  Write-Host "  + $Message" -ForegroundColor Green
}

function Write-WarningLine {
  param([Parameter(Mandatory)][string]$Message)

  Write-Host "  ! $Message" -ForegroundColor Yellow
}

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

  Write-InfoLine -Message 'Installing Azure CLI via winget...'
  winget install --exact --id Microsoft.AzureCLI --accept-source-agreements --accept-package-agreements | Out-Host
  $script:AzPath = Get-AzureCliPath

  if (-not $script:AzPath) {
    throw 'Azure CLI was installed, but the current session still cannot find az. Open a new terminal and re-run the script.'
  }

  Write-SuccessLine -Message 'Azure CLI is available.'
}

function Install-AzdIfNeeded {
  $script:AzdPath = Get-AzdPath
  if ($script:AzdPath) {
    return
  }

  if (-not (Test-WindowsHost)) {
    throw @(
      'Azure Developer CLI (azd) is not installed.',
      'Install it first and re-run this script.',
      'Docs: https://learn.microsoft.com/azure/developer/azure-developer-cli/install-azd'
    ) -join ' '
  }

  $winget = Get-Command -Name winget -ErrorAction SilentlyContinue
  if (-not $winget) {
    throw @(
      'Azure Developer CLI (azd) is not installed and winget is not available.',
      'Install it manually and re-run this script.',
      'Recommended Windows path: winget install microsoft.azd'
    ) -join ' '
  }

  $answer = (Read-Host -Prompt 'azd is required for the azd path. Install it now via winget? [Y/n]').Trim()
  if ($answer -ne '' -and $answer -notmatch '^[Yy]') {
    throw 'The azd deployment path requires Azure Developer CLI.'
  }

  Write-InfoLine -Message 'Installing Azure Developer CLI via winget...'
  winget install microsoft.azd --accept-source-agreements --accept-package-agreements | Out-Host
  $script:AzdPath = Get-AzdPath

  if (-not $script:AzdPath) {
    throw 'Azure Developer CLI was installed, but the current session still cannot find azd. Open a new terminal and re-run the script.'
  }

  Write-SuccessLine -Message 'Azure Developer CLI is available.'
}

function Install-BicepCliIfNeeded {
  if (Test-BicepReady) {
    return
  }

  $answer = (Read-Host -Prompt 'Bicep is not available yet. Install it now via az bicep install? [Y/n]').Trim()
  if ($answer -ne '' -and $answer -notmatch '^[Yy]') {
    throw 'The Bicep deployment path requires Bicep CLI.'
  }

  Write-InfoLine -Message 'Installing Bicep via Azure CLI...'
  Invoke-AzureCli -Arguments @('bicep', 'install') | Out-Null
  Write-SuccessLine -Message 'Bicep CLI is available.'
}

function Connect-AzureCliIfNeeded {
  try {
    Invoke-AzureCli -Arguments @('account', 'show', '--output', 'none') | Out-Null
  }
  catch {
    Write-InfoLine -Message 'No active Azure CLI session found. Starting az login...'
    Invoke-AzureCli -Arguments @('login') | Out-Null
  }

  $script:SubscriptionName = (Invoke-AzureCli -Arguments @('account', 'show', '--query', 'name', '-o', 'tsv')).Trim()
  $script:SubscriptionId = (Invoke-AzureCli -Arguments @('account', 'show', '--query', 'id', '-o', 'tsv')).Trim()
  $script:TenantId = (Invoke-AzureCli -Arguments @('account', 'show', '--query', 'tenantId', '-o', 'tsv')).Trim()
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

  if ($ToolState.AzdReady) {
    return 'Azd'
  }

  if ($ToolState.BicepReady) {
    return 'Bicep'
  }

  if ($ToolState.AzureCliReady) {
    return 'ArmJson'
  }

  return 'Bicep'
}

function Get-DefaultModeReason {
  param(
    [Parameter(Mandatory)][string]$SelectedMode,
    [Parameter(Mandatory)][pscustomobject]$ToolState
  )

  switch ($SelectedMode) {
    'Azd' {
      return 'azd and Azure CLI are already available, and this repository already contains an azd workflow with pre- and post-provision hooks.'
    }
    'Bicep' {
      if ($ToolState.BicepReady) {
        return 'Azure CLI and Bicep are already available, so the preferred direct CLI path is ready immediately.'
      }

      return 'Azure CLI plus az bicep install is the smallest modern install path when nothing is ready yet.'
    }
    'ArmJson' {
      return 'Azure CLI is already available, so ARM JSON is the fastest no-install fallback on this machine.'
    }
    default {
      return 'Automatic selection could not determine a better default.'
    }
  }
}

function Read-ModeChoice {
  param([Parameter(Mandatory)][string]$DefaultMode)

  Write-Host ''
  Write-Host 'Choose the console deployment method:'
  Write-Host '  [1] azd provision   (repo-native workflow; runs azd hooks)'
  Write-Host '  [2] Bicep           (preferred direct Azure CLI path)'
  Write-Host '  [3] ARM JSON        (direct compatibility fallback)'
  Write-Host ''

  $defaultOption = switch ($DefaultMode) {
    'Azd' { '1' }
    'Bicep' { '2' }
    'ArmJson' { '3' }
    default { '2' }
  }

  do {
    $choice = (Read-Host -Prompt "Option [1/2/3, default: $defaultOption]").Trim()
    if ($choice -eq '') {
      $choice = $defaultOption
    }
  } while ($choice -notin @('1', '2', '3'))

  switch ($choice) {
    '1' { return 'Azd' }
    '2' { return 'Bicep' }
    '3' { return 'ArmJson' }
  }
}

function Read-RequiredValue {
  param([Parameter(Mandatory)][string]$Prompt)

  do {
    $value = (Read-Host -Prompt $Prompt).Trim()
  } while (-not $value)

  return $value
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

function Read-BoolValue {
  param(
    [Parameter(Mandatory)][string]$Prompt,
    [Parameter(Mandatory)][bool]$DefaultValue
  )

  $defaultText = if ($DefaultValue) { 'true' } else { 'false' }
  do {
    $value = (Read-Host -Prompt "$Prompt [$defaultText]").Trim().ToLowerInvariant()
    if ($value -eq '') {
      return $DefaultValue
    }
  } while ($value -notin @('true', 'false'))

  return ($value -eq 'true')
}

function Get-DetectedTenantName {
  try {
    $derivedName = Invoke-AzureCli -Arguments @(
      'rest',
      '--method', 'GET',
      '--url', 'https://graph.microsoft.com/v1.0/organization?$select=verifiedDomains',
      '--query', 'value[0].verifiedDomains[?isDefault].name | [0]',
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
  Write-InfoLine -Message "Downloading $RelativePath from GitHub..."
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

  Write-SuccessLine -Message "Temporary repo assets staged at $($script:StagedRepoRoot)."
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

function Initialize-DeploymentMode {
  param([Parameter(Mandatory)][string]$SelectedMode)

  switch ($SelectedMode) {
    'Azd' {
      Install-AzureCliIfNeeded
      Install-AzdIfNeeded
      Connect-AzureCliIfNeeded
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
    Write-SuccessLine -Message "Resource group $Name already exists."
    return
  }

  $location = Read-DefaultValue -Prompt 'Azure location for the new resource group' -DefaultValue 'westeurope'
  Write-InfoLine -Message "Creating resource group $Name in $location..."
  Invoke-AzureCli -Arguments @('group', 'create', '--name', $Name, '--location', $location) | Out-Null
  Write-SuccessLine -Message "Resource group $Name created."
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

  Write-Section -Title 'Provider preflight'
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
        Write-SuccessLine -Message "$provider is registered."
      }
      'Registering' {
        Write-WarningLine -Message "$provider is still registering. Deployment can usually continue."
      }
      'NotRegistered' {
        Write-WarningLine -Message "$provider is not registered."
        $missingProviders += $provider
      }
      'Unregistered' {
        Write-WarningLine -Message "$provider is not registered."
        $missingProviders += $provider
      }
      default {
        Write-WarningLine -Message "$provider returned state: $state"
        $missingProviders += $provider
      }
    }
  }

  if ($missingProviders.Count -eq 0) {
    Write-SuccessLine -Message 'All required resource providers are ready.'
    return
  }

  foreach ($provider in $missingProviders) {
    Write-InfoLine -Message "Registering $provider..."
    try {
      Invoke-AzureCli -Arguments @('provider', 'register', '--namespace', $provider, '--wait') | Out-Null
      Write-SuccessLine -Message "$provider registered."
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

function Resolve-WebPartClientId {
  param([string]$CurrentClientId)

  if ($CurrentClientId) {
    return $CurrentClientId
  }

  $answer = (Read-Host -Prompt 'No webPartClientId supplied. Create or reuse the EasyAuth App Registration now? [Y/n]').Trim()
  if ($answer -eq '' -or $answer -match '^[Yy]') {
    $setupScriptPath = Get-RepoFilePath -RelativePath 'azure-function/infra/setup-app-registration.ps1'
    & $setupScriptPath -TenantId $script:TenantId -Confirm:$false

    $resolvedClientId = (Invoke-AzureCli -Arguments @(
        'ad', 'app', 'list',
        '--display-name', $script:AppRegistrationDisplayName,
        '--query', '[0].appId',
        '-o', 'tsv'
      )).Trim()

    if (-not $resolvedClientId) {
      throw 'The App Registration script completed, but no client ID could be resolved afterwards.'
    }

    return $resolvedClientId
  }

  return (Read-RequiredValue -Prompt 'EasyAuth App Registration Client ID (webPartClientId)')
}

function Invoke-AzdProvision {
  $repoRoot = Get-RepoRoot

  Write-Section -Title 'azd provision'
  Write-InfoLine -Message 'Running azd provision because this repository already contains azure.yaml and azd hooks.'
  Write-InfoLine -Message 'Use azd up later if you also want azd to handle future code redeploy cycles.'

  Push-Location -Path $repoRoot
  try {
    Invoke-Azd -Arguments @('provision')
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

  Write-Section -Title 'Bicep deployment'
  Invoke-AzureCli -Arguments @(
    'deployment', 'group', 'create',
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

  Write-Section -Title 'ARM JSON deployment'
  Invoke-AzureCli -Arguments @(
    'deployment', 'group', 'create',
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

try {
  Write-Section -Title 'Guest Sponsor Info deployment helper'

  $toolState = Get-ToolState
  $defaultMode = Select-DefaultMode -ToolState $toolState
  $defaultReason = Get-DefaultModeReason -SelectedMode $defaultMode -ToolState $toolState

  Write-Host 'Detected local tooling:'
  Write-Host "  az       : $(if ($toolState.AzureCliReady) { 'ready' } else { 'missing' })"
  Write-Host "  az bicep : $(if ($toolState.BicepReady) { 'ready' } elseif ($toolState.AzureCliReady) { 'available after az bicep install' } else { 'Azure CLI missing' })"
  if ($toolState.AzdReady) {
    Write-Host '  azd      : ready'
  }
  elseif ($toolState.AzdInstalled) {
    Write-Host '  azd      : installed, but Azure CLI is still missing for this repository workflow'
  }
  else {
    Write-Host '  azd      : missing'
  }

  Write-Host ''
  Write-Host "Default suggestion: $defaultMode"
  Write-Host "Reason: $defaultReason"

  $selectedMode = $Mode
  if ($selectedMode -eq 'Auto') {
    $selectedMode = Read-ModeChoice -DefaultMode $defaultMode
  }

  Initialize-DeploymentMode -SelectedMode $selectedMode

  Write-Host ''
  Write-SuccessLine -Message "Active subscription: $($script:SubscriptionName) ($($script:SubscriptionId))"
  Write-SuccessLine -Message "Tenant ID: $($script:TenantId)"

  if ($selectedMode -eq 'Azd') {
    Invoke-AzdProvision
    return
  }

  if (-not $ResourceGroupName) {
    $ResourceGroupName = Read-RequiredValue -Prompt 'Azure resource group name'
  }

  Initialize-ResourceGroup -Name $ResourceGroupName

  if (-not $TenantName) {
    $detectedTenantName = Get-DetectedTenantName
    if ($detectedTenantName) {
      $TenantName = Read-DefaultValue -Prompt 'SharePoint tenant name (without .sharepoint.com)' -DefaultValue $detectedTenantName
    }
    else {
      $TenantName = Read-RequiredValue -Prompt 'SharePoint tenant name (without .sharepoint.com)'
    }
  }

  if (-not $FunctionAppName) {
    $FunctionAppName = Read-RequiredValue -Prompt 'Globally unique Function App name'
  }

  $HostingPlan = Read-DefaultValue -Prompt 'Hosting plan (Consumption or FlexConsumption)' -DefaultValue $HostingPlan
  $DeployAzureMaps = Read-BoolValue -Prompt 'Deploy Azure Maps (true or false)' -DefaultValue $DeployAzureMaps
  $AppVersion = Read-DefaultValue -Prompt 'Function package version' -DefaultValue $AppVersion
  $EnableMonitoring = Read-BoolValue -Prompt 'Enable monitoring stack (Log Analytics, Application Insights, alerts) (true or false)' -DefaultValue $EnableMonitoring
  if ($EnableMonitoring) {
    $EnableFailureAnomaliesAlert = Read-BoolValue -Prompt 'Enable Failure Anomalies smart detector alert (true or false)' -DefaultValue $EnableFailureAnomaliesAlert
  }
  else {
    $EnableFailureAnomaliesAlert = $false
  }

  if ($HostingPlan -eq 'FlexConsumption') {
    $MaximumFlexInstances = [int](Read-DefaultValue -Prompt 'Maximum Flex instances' -DefaultValue ([string]$MaximumFlexInstances))
  }

  $WebPartClientId = Resolve-WebPartClientId -CurrentClientId $WebPartClientId
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
      -MonitoringEnabled:$EnableMonitoring `
      -EnableFailureAlert:$EnableFailureAnomaliesAlert `
      -FlexScaleLimit $MaximumFlexInstances
  }
}
finally {
  foreach ($path in $script:TempPaths) {
    if (Test-Path -Path $path) {
      Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
    }
  }
}
