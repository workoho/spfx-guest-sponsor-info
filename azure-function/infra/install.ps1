#!/usr/bin/env -S pwsh -NoLogo -NoProfile

<#
.SYNOPSIS
    Downloads the Guest Sponsor Info infra package and runs the deployment wizard.

.DESCRIPTION
    Downloads the infra release package from GitHub, extracts it to a temporary
    directory, and runs deploy-azure.ps1. All parameters are forwarded to the
    wizard. The temporary directory is removed when the wizard exits.

    This script is the recommended iwr entry point:

      & ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.ps1').Content))

.PARAMETER Version
    Release tag to download (e.g. "v1.2.0"). Defaults to "latest".

.PARAMETER AzdEnvironmentName
    Forwarded to deploy-azure.ps1.

.PARAMETER ResourceGroupName
    Forwarded to deploy-azure.ps1.

.PARAMETER AzureLocation
    Forwarded to deploy-azure.ps1.

.PARAMETER AzureTenantId
    Forwarded to deploy-azure.ps1.

.PARAMETER TenantName
    Forwarded to deploy-azure.ps1.

.PARAMETER FunctionAppName
    Forwarded to deploy-azure.ps1.

.PARAMETER HostingPlan
    Forwarded to deploy-azure.ps1.

.PARAMETER DeployAzureMaps
    Forwarded to deploy-azure.ps1.

.PARAMETER AppVersion
    Forwarded to deploy-azure.ps1.

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
    & ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.ps1').Content))

.EXAMPLE
    & ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.ps1').Content)) -Version v1.2.0 -ResourceGroupName rg-gsi -TenantName contoso

.NOTES
    Copyright 2026 Workoho GmbH <https://workoho.com>
    Author: Julian Pawlowski <https://github.com/jpawlowski>
    Licensed under PolyForm Shield License 1.0.0
    <https://polyformproject.org/licenses/shield/1.0.0>
#>

#Requires -Version 5.1
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
  [string]$Version = 'latest',
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

# ── Resolve download URL ──────────────────────────────────────────────────────
# "latest" resolves via the GitHub "latest" redirect; a specific tag uses
# the direct download path.
$_baseUrl = 'https://github.com/workoho/spfx-guest-sponsor-info/releases'
$_zipUrl = if ($Version -eq 'latest') {
  "$_baseUrl/latest/download/guest-sponsor-info-infra.zip"
}
else {
  "$_baseUrl/download/$Version/guest-sponsor-info-infra.zip"
}

# ── Temporary paths ───────────────────────────────────────────────────────────
$_tempBase = [System.IO.Path]::GetTempPath()
$_tempSuffix = [System.Guid]::NewGuid().ToString('n')
# Distinct paths for the ZIP and extracted directory so cleanup is explicit.
$_zipFile = Join-Path $_tempBase "gsi-infra-$_tempSuffix.zip"
$_extractDir = Join-Path $_tempBase "gsi-infra-$_tempSuffix"

Write-Host ''
Write-Host '  Guest Sponsor Info  ·  Installer' -ForegroundColor DarkCyan
Write-Host ('  ' + ('─' * 58)) -ForegroundColor DarkGray
Write-Host "  Downloading infra package ($Version)..." -ForegroundColor DarkGray
Write-Host "  Source: $_zipUrl" -ForegroundColor DarkGray
Write-Host ''

try {
  # Download the infra ZIP to a temp file.
  Invoke-WebRequest -Uri $_zipUrl -OutFile $_zipFile -UseBasicParsing

  # Extract the ZIP.
  Expand-Archive -Path $_zipFile -DestinationPath $_extractDir -Force

  # Locate deploy-azure.ps1 inside the extracted tree.
  # The infra ZIP is flat — all files land at the ZIP root, so deploy-azure.ps1
  # is directly inside the extract directory (no azure-function/infra/ prefix).
  $_deployScript = Join-Path $_extractDir 'deploy-azure.ps1'
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

  & $_deployScript @_forwardParams
}
finally {
  # Always remove temp files, even when the wizard throws or the user aborts.
  Remove-Item -Path $_zipFile -Force -ErrorAction SilentlyContinue
  Remove-Item -Path $_extractDir -Recurse -Force -ErrorAction SilentlyContinue
}
