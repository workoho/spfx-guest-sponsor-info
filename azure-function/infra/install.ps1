#!/usr/bin/env -S pwsh -NoLogo -NoProfile

<#
.SYNOPSIS
    Downloads the Guest Sponsor Info infra package and runs the deployment wizard.

.DESCRIPTION
  Downloads the Guest Sponsor Info infra package or a repository snapshot
  from GitHub, extracts it to a temporary directory, and runs
  deploy-azure.ps1. All parameters are forwarded to the wizard. The
  temporary directory is removed when the wizard exits.

    This script is the recommended iwr entry point:

      & ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/v1.2.3/azure-function/infra/install.ps1').Content))

.PARAMETER Version
  Installer payload source. Supports release tags (e.g. "v1.2.0"),
  "latest" for the newest published release, and "main" for the current
  main branch snapshot. When omitted, the installer first reuses its stamped
  installer ref when that ref points to a release tag. As a compatibility
  fallback, it can still infer a release tag from the install.ps1 URL and
  otherwise falls back to the newest published release ("latest"). Release
  packaging stamps installer metadata so tagged install.ps1 files resolve their
  own release without mutating the Version parameter default.

.PARAMETER AzdEnvironmentName
    Forwarded to deploy-azure.ps1.

.PARAMETER ResourceGroupName
    Forwarded to deploy-azure.ps1.

.PARAMETER AzureLocation
    Forwarded to deploy-azure.ps1.

.PARAMETER AzureTenantId
    Forwarded to deploy-azure.ps1.

.PARAMETER AzureLoginMode
  Forwarded to deploy-azure.ps1. Overrides automatic Azure CLI login-mode
  detection. Supported values: "auto", "browser", "device-code". In Azure
  Cloud Shell and on remote or headless terminals, "auto" selects device code.
  That flow must be permitted for your account, and PAW-restricted admin
  accounts have to confirm the code on their Privileged Access Workstation.

.PARAMETER TenantName
    Forwarded to deploy-azure.ps1.

.PARAMETER FunctionAppName
    Forwarded to deploy-azure.ps1.

.PARAMETER DeployAzureMaps
    Forwarded to deploy-azure.ps1.

.PARAMETER AppVersion
  Advanced override for the Function package version. When omitted,
  install.ps1 reuses `-Version` for `latest` and release tags, while
  `-Version main` keeps the function package on `latest`. Most installations
  should leave AppVersion unset and use it only to force a different published
  Function release.

.PARAMETER EnableMonitoring
    Forwarded to deploy-azure.ps1.

.PARAMETER EnableFailureAnomaliesAlert
    Forwarded to deploy-azure.ps1.

.PARAMETER MaximumFlexInstances
    Forwarded to deploy-azure.ps1.

.PARAMETER AlwaysReadyInstances
    Forwarded to deploy-azure.ps1.

.PARAMETER InstanceMemoryMB
    Forwarded to deploy-azure.ps1.

.PARAMETER SkipGraphRoleAssignments
    Forwarded to deploy-azure.ps1.

.PARAMETER PreflightOnly
    Forwarded to deploy-azure.ps1.

.EXAMPLE
    & ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/v1.2.3/azure-function/infra/install.ps1').Content))

.EXAMPLE
    & ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/v1.2.3/azure-function/infra/install.ps1').Content)) -Version v1.2.0 -ResourceGroupName rg-gsi -TenantName contoso

.NOTES
    Copyright 2026 Workoho GmbH <https://workoho.com>
    Author: Julian Pawlowski <https://github.com/jpawlowski>
    Licensed under PolyForm Shield License 1.0.0
    <https://polyformproject.org/licenses/shield/1.0.0>
#>

#Requires -Version 5.1
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
  [string]$Version = '',
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
  [bool]$EnableMonitoring = $true,
  [bool]$EnableFailureAnomaliesAlert = $false,
  [int]$AlwaysReadyInstances = 0,
  [int]$MaximumFlexInstances = 10,
  [ValidateSet(512, 2048)]
  [int]$InstanceMemoryMB = 512,
  [bool]$SkipGraphRoleAssignments = $false,
  [switch]$PreflightOnly
)

$ErrorActionPreference = 'Stop'
$script:InstallerRef = 'v1.2.3'
$script:InstallerInvocationLine = [string]$MyInvocation.Line

function Get-HttpStatusCodeFromException {
  param([System.Exception]$Exception)

  $response = $Exception.Response
  if ($null -eq $response) {
    return $null
  }

  try {
    return [int]$response.StatusCode
  }
  catch {
    return $null
  }
}

function Get-ReleaseAssetSha256 {
  param(
    [string]$Version,
    [string]$AssetName
  )

  $_apiUrl = if ($Version -eq 'latest') {
    'https://api.github.com/repos/workoho/spfx-guest-sponsor-info/releases/latest'
  }
  else {
    "https://api.github.com/repos/workoho/spfx-guest-sponsor-info/releases/tags/$Version"
  }

  try {
    $_release = Invoke-RestMethod `
      -Uri $_apiUrl `
      -Headers @{
      'Accept'               = 'application/vnd.github+json'
      'X-GitHub-Api-Version' = '2022-11-28'
      'User-Agent'           = 'gsi-installer'
    } `
      -UseBasicParsing
  }
  catch {
    $_statusCode = Get-HttpStatusCodeFromException -Exception $_.Exception
    if ($_statusCode -eq 404) {
      if ($Version -eq 'latest') {
        throw "GitHub API lookup for the latest release returned 404. Cannot resolve checksum fallback."
      }

      throw "GitHub API lookup for release '$Version' returned 404. Cannot resolve checksum fallback."
    }

    throw "GitHub API lookup for release asset checksum failed: $($_.Exception.Message)"
  }

  $_asset = @($_release.assets) | Where-Object { $_.name -eq $AssetName } | Select-Object -First 1
  if ($null -eq $_asset) {
    throw "Asset '$AssetName' not found in GitHub release metadata. Cannot resolve checksum fallback."
  }

  $_digest = [string]$_asset.digest
  if ([string]::IsNullOrWhiteSpace($_digest)) {
    throw "Asset '$AssetName' has no digest in GitHub release metadata. Cannot resolve checksum fallback."
  }

  if ($_digest -notmatch '^sha256:') {
    throw "Asset '$AssetName' digest is not SHA256 ('$_digest'). Cannot verify download."
  }

  return $_digest.Substring(7).ToUpperInvariant()
}

function Test-ReleasePackageVersion {
  param([Parameter(Mandatory)][string]$Version)

  return $Version -eq 'latest' -or $Version -match '^v[0-9]+\.[0-9]+\.[0-9]+([.-][A-Za-z0-9.]+)?$'
}

function Test-SupportedInstallerVersion {
  param([Parameter(Mandatory)][string]$Version)

  return $Version -eq 'main' -or (Test-ReleasePackageVersion -Version $Version)
}

function Assert-SupportedInstallerVersion {
  param(
    [Parameter(Mandatory)][string]$Version,
    [string]$ParameterName = 'Version'
  )

  if (-not (Test-SupportedInstallerVersion -Version $Version)) {
    throw "Unsupported $ParameterName '$Version'. Supported values: latest, main, or a release tag like v1.2.3."
  }
}

function Get-InstallerInvocationRefFallback {
  $_invocationLine = $script:InstallerInvocationLine
  if ([string]::IsNullOrWhiteSpace($_invocationLine)) {
    return ''
  }

  if ($_invocationLine -match 'raw\.githubusercontent\.com/workoho/spfx-guest-sponsor-info/(?<ref>.+?)/azure-function/infra/install\.ps1') {
    return $Matches.ref
  }

  return ''
}

function Get-InstallerSourceRef {
  $_declaredRef = [string]$script:InstallerRef
  if (-not [string]::IsNullOrWhiteSpace($_declaredRef) -and $_declaredRef -ne 'main') {
    return $_declaredRef
  }

  $_invocationRef = Get-InstallerInvocationRefFallback
  if (-not [string]::IsNullOrWhiteSpace($_invocationRef)) {
    return $_invocationRef
  }

  return $_declaredRef
}

function Get-InstallerImplicitVersion {
  $_sourceRef = Get-InstallerSourceRef
  if (Test-ReleasePackageVersion -Version $_sourceRef) {
    return $_sourceRef
  }

  return ''
}

function Resolve-InstallerPayloadVersion {
  param([string]$RequestedVersion)

  if (-not [string]::IsNullOrWhiteSpace($RequestedVersion)) {
    Assert-SupportedInstallerVersion -Version $RequestedVersion
    return $RequestedVersion
  }

  $_implicitVersion = Get-InstallerImplicitVersion
  if (-not [string]::IsNullOrWhiteSpace($_implicitVersion)) {
    return $_implicitVersion
  }

  return 'latest'
}

function Get-InstallerDownloadPlan {
  param([Parameter(Mandatory)][string]$Version)

  $_releaseBaseUrl = 'https://github.com/workoho/spfx-guest-sponsor-info/releases'
  Assert-SupportedInstallerVersion -Version $Version

  if (Test-ReleasePackageVersion -Version $Version) {
    return [pscustomobject]@{
      SourceKind   = 'release-package'
      DisplayLabel = "release package ($Version)"
      ZipUrl       = if ($Version -eq 'latest') {
        "$_releaseBaseUrl/latest/download/guest-sponsor-info-infra.zip"
      }
      else {
        "$_releaseBaseUrl/download/$Version/guest-sponsor-info-infra.zip"
      }
      ChecksumsUrl = if ($Version -eq 'latest') {
        "$_releaseBaseUrl/latest/download/checksums.txt"
      }
      else {
        "$_releaseBaseUrl/download/$Version/checksums.txt"
      }
    }
  }

  return [pscustomobject]@{
    SourceKind   = 'repo-snapshot'
    DisplayLabel = 'repository snapshot (main)'
    ZipUrl       = 'https://github.com/workoho/spfx-guest-sponsor-info/archive/refs/heads/main.zip'
    ChecksumsUrl = $null
  }
}

function Get-FunctionPackageDisplayText {
  param([Parameter(Mandatory)][string]$Version)

  if ($Version -eq 'latest') {
    return 'latest release'
  }

  return $Version
}

function Resolve-FunctionPackageVersion {
  param(
    [Parameter(Mandatory)][string]$RequestedAppVersion,
    [Parameter(Mandatory)][string]$ResolvedInstallerVersion,
    [Parameter(Mandatory)][bool]$AppVersionWasExplicitlySet
  )

  if ($AppVersionWasExplicitlySet) {
    return $RequestedAppVersion
  }

  if (Test-ReleasePackageVersion -Version $ResolvedInstallerVersion) {
    return $ResolvedInstallerVersion
  }

  return 'latest'
}

# ── Resolve download URLs ─────────────────────────────────────────────────────
$_resolvedVersion = Resolve-InstallerPayloadVersion -RequestedVersion $Version
$_resolvedAppVersion = Resolve-FunctionPackageVersion `
  -RequestedAppVersion $AppVersion `
  -ResolvedInstallerVersion $_resolvedVersion `
  -AppVersionWasExplicitlySet $PSBoundParameters.ContainsKey('AppVersion')
$_downloadPlan = Get-InstallerDownloadPlan -Version $_resolvedVersion
$_zipUrl = $_downloadPlan.ZipUrl
$_checksumsUrl = $_downloadPlan.ChecksumsUrl

# ── Temporary paths ───────────────────────────────────────────────────────────
$_tempBase = [System.IO.Path]::GetTempPath()
$_tempSuffix = [System.Guid]::NewGuid().ToString('n')
$_zipFile = Join-Path $_tempBase "gsi-infra-$_tempSuffix.zip"
$_checksumsFile = Join-Path $_tempBase "gsi-infra-$_tempSuffix-checksums.txt"
$_extractDir = Join-Path $_tempBase "gsi-infra-$_tempSuffix"

Write-Host ''
Write-Host '  Guest Sponsor Info  ·  Installer' -ForegroundColor DarkCyan
Write-Host ('  ' + ('─' * 58)) -ForegroundColor DarkGray
Write-Host "  Installer payload : $($_downloadPlan.DisplayLabel)" -ForegroundColor DarkGray
Write-Host "  Function package  : $(Get-FunctionPackageDisplayText -Version $_resolvedAppVersion)" -ForegroundColor DarkGray
Write-Host "  Downloading $($_downloadPlan.DisplayLabel)..." -ForegroundColor DarkGray
Write-Host "  Source: $_zipUrl" -ForegroundColor DarkGray
Write-Host ''

try {
  # Download the infra ZIP to a temp file.
  try {
    Invoke-WebRequest -Uri $_zipUrl -OutFile $_zipFile -UseBasicParsing
  }
  catch {
    $_statusCode = Get-HttpStatusCodeFromException -Exception $_.Exception
    if ($_statusCode -eq 404) {
      if ($_downloadPlan.SourceKind -eq 'release-package' -and $_resolvedVersion -eq 'latest') {
        throw (
          "Infra package for the latest release was not found (404).`n" +
          "The release may not be fully published yet.`n" +
          "Retry in a few minutes, or run again with -Version vX.Y.Z."
        )
      }

      if ($_downloadPlan.SourceKind -eq 'release-package') {
        throw (
          "Infra package for release '$_resolvedVersion' was not found (404).`n" +
          "Verify the release tag and check that release assets are published.`n" +
          "URL: $_zipUrl"
        )
      }

      throw (
        "Repository snapshot for ref '$_resolvedVersion' was not found (404).`n" +
        "Verify the branch name and that the ref exists on GitHub.`n" +
        "URL: $_zipUrl"
      )
    }

    throw
  }

  if ($_downloadPlan.SourceKind -eq 'release-package') {
    # ── SHA256 integrity check ────────────────────────────────────────────────
    # Primary source: checksums.txt from the same release.
    # Fallback source: GitHub release asset digest (sha256) from the API.
    # This catches truncated downloads or CDN-level tampering.
    Write-Host '  Verifying SHA256 checksum...' -ForegroundColor DarkGray
    $_checksumsDownloaded = $false
    try {
      Invoke-WebRequest -Uri $_checksumsUrl -OutFile $_checksumsFile -UseBasicParsing
      $_checksumsDownloaded = $true
    }
    catch {
      $_statusCode = Get-HttpStatusCodeFromException -Exception $_.Exception
      if ($_statusCode -eq 404) {
        Write-Warning (
          "checksums.txt was not found (404). Falling back to GitHub release asset digest for integrity verification.`n" +
          "URL: $_checksumsUrl"
        )
      }
      else {
        throw
      }
    }

    $_expectedHash = $null

    if ($_checksumsDownloaded) {
      # Parse "hash  filename" lines; find the entry for guest-sponsor-info-infra.zip.
      foreach ($_line in (Get-Content $_checksumsFile)) {
        $_parts = $_line -split '\s+', 2
        if ($_parts.Length -eq 2 -and $_parts[1].Trim() -eq 'guest-sponsor-info-infra.zip') {
          $_expectedHash = $_parts[0].Trim().ToUpperInvariant()
          break
        }
      }
    }

    if (-not $_expectedHash) {
      if ($_checksumsDownloaded) {
        Write-Warning 'guest-sponsor-info-infra.zip entry was not found in checksums.txt. Falling back to GitHub release asset digest.'
      }

      $_expectedHash = Get-ReleaseAssetSha256 -Version $_resolvedVersion -AssetName 'guest-sponsor-info-infra.zip'
      Write-Host '  Using checksum source: GitHub release asset digest' -ForegroundColor DarkGray
    }

    $_actualHash = (Get-FileHash -Path $_zipFile -Algorithm SHA256).Hash.ToUpperInvariant()
    if ($_actualHash -ne $_expectedHash) {
      throw (
        "SHA256 mismatch for guest-sponsor-info-infra.zip.`n" +
        "  Expected: $_expectedHash`n" +
        "  Actual:   $_actualHash`n" +
        "Download may be corrupt or tampered with. Aborting."
      )
    }
    Write-Host '  checksum verified.' -ForegroundColor DarkGreen
    Write-Host ''
  }
  else {
    Write-Host '  Repository snapshots do not publish release checksums; skipping SHA256 verification.' -ForegroundColor DarkGray
    Write-Host ''
  }

  # Extract the ZIP (hash verified above).
  Expand-Archive -Path $_zipFile -DestinationPath $_extractDir -Force

  # Locate deploy-azure.ps1 inside the extracted tree.
  if ($_downloadPlan.SourceKind -eq 'release-package') {
    # The infra ZIP is flat — all files land at the ZIP root, so deploy-azure.ps1
    # is directly inside the extract directory (no azure-function/infra/ prefix).
    $_deployScript = Join-Path $_extractDir 'deploy-azure.ps1'
  }
  else {
    $_repoRoot = @(
      Get-ChildItem -Path $_extractDir -Directory -ErrorAction Stop | Select-Object -First 1
    )[0].FullName
    $_deployScript = Join-Path $_repoRoot 'azure-function/infra/deploy-azure.ps1'
  }

  if (-not (Test-Path $_deployScript)) {
    throw "deploy-azure.ps1 not found in the downloaded package. Expected: $_deployScript"
  }

  # Forward all parameters except Version to the wizard.
  # Build the forwarded params hash from PSBoundParameters, skipping Version.
  $_forwardParams = @{}
  foreach ($_key in $PSBoundParameters.Keys) {
    if ($_key -ne 'Version') {
      $_forwardParams[$_key] = $PSBoundParameters[$_key]
    }
  }
  if (-not $PSBoundParameters.ContainsKey('AppVersion')) {
    $_forwardParams['AppVersion'] = $_resolvedAppVersion
  }
  $_forwardParams['InstallerVersion'] = $_resolvedVersion

  & $_deployScript @_forwardParams
}
finally {
  # Always remove temp files, even when the wizard throws or the user aborts.
  Remove-Item -Path $_zipFile -Force -ErrorAction SilentlyContinue
  Remove-Item -Path $_checksumsFile -Force -ErrorAction SilentlyContinue
  Remove-Item -Path $_extractDir -Recurse -Force -ErrorAction SilentlyContinue
}
