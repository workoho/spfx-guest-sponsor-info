#!/usr/bin/env pwsh
# SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
# SPDX-FileCopyrightText: 2026 Julian Pawlowski <https://github.com/jpawlowski>
# SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0
#
# Pre-provision hook for Azure Developer CLI (azd).
# Runs before Bicep deployment to:
#   1. Derive a default Function App name from the azd environment name.
#   2. Detect or prompt for the SharePoint tenant name.
#
# The Entra App Registration and Microsoft Graph permission assignments are
# now managed declaratively by the Bicep template (Microsoft Graph Bicep
# extension v1.0).  The deploying principal needs:
#   - Application.ReadWrite.All  (Cloud Application Administrator,
#                                  Application Administrator, or Global Administrator)
#   - AppRoleAssignment.ReadWrite.All  (Privileged Role Administrator
#                                        or Global Administrator)
#
# All operations are idempotent — safe to re-run on 'azd provision' or 'azd up'.

$ErrorActionPreference = 'Stop'

$_invocationPath = (Get-Location).Path

function Get-AzdProjectRoot {
  if ($env:AZD_PROJECT_PATH -and (Test-Path -Path (Join-Path -Path $env:AZD_PROJECT_PATH -ChildPath 'azure.yaml'))) {
    return $env:AZD_PROJECT_PATH
  }

  if (Test-Path -Path (Join-Path -Path $_invocationPath -ChildPath 'azure.yaml')) {
    return $_invocationPath
  }

  $_candidate = $PSScriptRoot
  while ($_candidate) {
    if (Test-Path -Path (Join-Path -Path $_candidate -ChildPath 'azure.yaml')) {
      return $_candidate
    }

    $_parent = Split-Path -Path $_candidate -Parent
    if (-not $_parent -or $_parent -eq $_candidate) {
      break
    }
    $_candidate = $_parent
  }

  throw 'azure.yaml not found for azd pre-provision hook.'
}

Set-Location -Path (Get-AzdProjectRoot)

$envValues = azd env get-values

# ── 0a. Check Azure RBAC permission ─────────────────────────────────────────
# Contributor (or Owner) on the subscription is needed to register resource
# providers and to deploy Bicep resources.  The check is informational — a
# missing role does not abort the script, but it surfaces the gap early so
# the operator can activate a PIM role or request access before the actual
# deployment runs.
Write-Host ''
Write-Host 'Checking Azure role assignment...'
$_subIdMatch = ($envValues | Select-String '^AZURE_SUBSCRIPTION_ID="?([^"]+)"?').Matches
$_subId = if ($_subIdMatch -and $_subIdMatch.Count -gt 0) { $_subIdMatch[0].Groups[1].Value } else { $null }
if ($_subId) {
  try {
    $_userId = az ad signed-in-user show --query id -o tsv 2>$null
    if ($LASTEXITCODE -eq 0 -and $_userId) {
      $_rbacRaw = az role assignment list `
        --scope "/subscriptions/$_subId" `
        --assignee "$_userId" `
        --include-inherited `
        --query "[?contains(['Owner','Contributor'], roleDefinitionName)].roleDefinitionName" `
        -o tsv 2>$null
      if ($LASTEXITCODE -eq 0) {
        $_rbacList = @($_rbacRaw -split "`n" | Where-Object { $_ } | Select-Object -Unique)
        if ($_rbacList.Count -gt 0) {
          Write-Host "  + Azure RBAC: $($_rbacList -join ', ') on subscription."
        }
        else {
          Write-Host '  ! Azure RBAC: no Contributor or Owner role found on this subscription.'
          Write-Host '    Both are required for resource provider registration and Bicep deployment.'
          Write-Host '    Contact your subscription owner to request Contributor access or activate'
          Write-Host '    an eligible role via Azure PIM before re-running azd provision.'
          Write-Host '    Azure PIM: https://portal.azure.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/azurerbac'
        }
      }
      else {
        Write-Host '  ! Azure RBAC: role listing failed — continuing anyway.'
        Write-Host '    Required: Contributor or Owner on the subscription.'
      }
    }
    else {
      Write-Host '  ! Azure RBAC: could not identify the signed-in user — skipping check.'
      Write-Host '    Required: Contributor or Owner on the subscription.'
    }
  }
  catch {
    Write-Host '  ! Azure RBAC: check encountered an error — continuing anyway.'
    Write-Host '    Required: Contributor or Owner on the subscription.'
  }
}
else {
  Write-Host '  ! Azure RBAC: AZURE_SUBSCRIPTION_ID not yet set — skipping role check.'
  Write-Host '    Required: Contributor or Owner on the subscription.'
}
Write-Host ''

# ── 0. Validate required Azure resource providers ───────────────────────────
# Read from azd env — set by deploy-azure.ps1 via 'azd env set' before running provision.
# Fall back to main.parameters.json defaults when running azd directly without the wizard.
$_planMatch = ($envValues | Select-String '^AZURE_HOSTING_PLAN="?([^"]+)"?').Matches
$hostingPlan = if ($_planMatch -and $_planMatch.Count -gt 0) { $_planMatch[0].Groups[1].Value } else { 'Consumption' }
$_mapsMatch = ($envValues | Select-String '^AZURE_DEPLOY_AZURE_MAPS="?([^"]+)"?').Matches
$deployAzureMaps = -not ($_mapsMatch -and $_mapsMatch.Count -gt 0 -and $_mapsMatch[0].Groups[1].Value -eq 'false')
$requiredProviders = @(
  'Microsoft.AlertsManagement',
  'Microsoft.Authorization',
  'Microsoft.Insights',
  'Microsoft.ManagedIdentity',
  'Microsoft.OperationalInsights',
  'Microsoft.Resources',
  'Microsoft.Storage',
  'Microsoft.Web'
)

if ($hostingPlan -eq 'FlexConsumption') {
  $requiredProviders += 'Microsoft.ContainerInstance'
}

if ($deployAzureMaps) {
  $requiredProviders += 'Microsoft.Maps'
}

$requiredProviders = $requiredProviders | Sort-Object -Unique
$missingProviders = @()

Write-Host 'Checking required Azure resource providers...'
foreach ($provider in $requiredProviders) {
  $state = az provider show --namespace $provider --query registrationState -o tsv 2>$null

  switch ($state) {
    'Registered' {
      Write-Host "  + $provider is registered."
    }
    'Registering' {
      Write-Host "  ! $provider is still registering. Deployment can usually continue."
    }
    'NotRegistered' {
      Write-Host "  ! $provider is not registered."
      $missingProviders += $provider
    }
    'Unregistered' {
      Write-Host "  ! $provider is not registered."
      $missingProviders += $provider
    }
    '' {
      Write-Host "  ! $provider returned no state."
      $missingProviders += $provider
    }
    default {
      Write-Host "  ! $provider returned state: $state"
      $missingProviders += $provider
    }
  }
}

if ($missingProviders.Count -gt 0) {
  Write-Host 'Registering missing Azure resource providers...'
  foreach ($provider in $missingProviders) {
    Write-Host "  -> az provider register --namespace $provider --wait"
    try {
      az provider register --namespace $provider --wait | Out-Null
      Write-Host "  + $provider registered."
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
else {
  Write-Host '  + All required resource providers are ready.'
}

# ── 1. Detect or prompt for SharePoint tenant name ──────────────────────────
if ($envValues -notmatch 'AZURE_SHAREPOINT_TENANT_NAME=') {
  $derived = $null
  try {
    $raw = az rest `
      --method GET `
      --url 'https://graph.microsoft.com/v1.0/organization?$select=verifiedDomains' `
      --query 'value[0].verifiedDomains[?isInitial].name | [0]' `
      -o tsv 2>$null
    $derived = $raw -replace '\.onmicrosoft\.com$', ''
  }
  catch {
    # az rest for tenant detection is best-effort; failure is handled by the
    # Read-Host prompt below.
    Write-Verbose "Tenant name detection via az rest failed: $_"
  }

  if ($derived) {
    Write-Host "Detected SharePoint tenant name: $derived"
    azd env set AZURE_SHAREPOINT_TENANT_NAME $derived
  }
  else {
    $tenantName = Read-Host "Enter your SharePoint tenant name (e.g. 'contoso' for contoso.sharepoint.com)"
    azd env set AZURE_SHAREPOINT_TENANT_NAME $tenantName
  }
}

if ($envValues -notmatch 'AZURE_HOSTING_PLAN=') {
  azd env set AZURE_HOSTING_PLAN 'Consumption'
}

if ($envValues -notmatch 'AZURE_DEPLOY_AZURE_MAPS=') {
  azd env set AZURE_DEPLOY_AZURE_MAPS 'true'
}

if ($envValues -notmatch 'AZURE_TAG_ENVIRONMENT=') {
  azd env set AZURE_TAG_ENVIRONMENT ''
}

if ($envValues -notmatch 'AZURE_TAG_CRITICALITY=') {
  azd env set AZURE_TAG_CRITICALITY ''
}

if ($envValues -notmatch 'AZURE_APP_VERSION=') {
  azd env set AZURE_APP_VERSION 'latest'
}

if ($envValues -notmatch 'AZURE_ENABLE_MONITORING=') {
  azd env set AZURE_ENABLE_MONITORING 'true'
}

if ($envValues -notmatch 'AZURE_ENABLE_FAILURE_ANOMALIES_ALERT=') {
  azd env set AZURE_ENABLE_FAILURE_ANOMALIES_ALERT 'false'
}

if ($envValues -notmatch 'AZURE_ALWAYS_READY_INSTANCES=') {
  azd env set AZURE_ALWAYS_READY_INSTANCES '1'
}

if ($envValues -notmatch 'AZURE_MAXIMUM_FLEX_INSTANCES=') {
  azd env set AZURE_MAXIMUM_FLEX_INSTANCES '10'
}

if ($envValues -notmatch 'AZURE_INSTANCE_MEMORY_MB=') {
  azd env set AZURE_INSTANCE_MEMORY_MB '2048'
}

# ── 3. Entra role check ──────────────────────────────────────────────────────
# Cloud Application Administrator (or Application Administrator / Global Admin)
# is always required — Bicep creates and manages the App Registration.
#
# Privileged Role Administrator (or Global Admin) is required for Graph app role
# assignments to the Managed Identity. If that role is not available, defer by
# setting AZURE_SKIP_GRAPH_ROLE_ASSIGNMENTS=true before running azd provision:
#
#   azd env set AZURE_SKIP_GRAPH_ROLE_ASSIGNMENTS true
#
# Then run setup-graph-permissions.ps1 after deployment with the
# managedIdentityObjectId Bicep output (azd env get-values).
Write-Host ''
Write-Host 'Checking Entra roles...'
$_appRegRoles = @('Cloud Application Administrator', 'Application Administrator', 'Global Administrator')
$_assignmentRoles = @('Privileged Role Administrator', 'Global Administrator')
$_skipRoleAssignments = $env:AZURE_SKIP_GRAPH_ROLE_ASSIGNMENTS -eq 'true'
try {
  $_entraRaw = az rest `
    --method GET `
    --url 'https://graph.microsoft.com/v1.0/me/transitiveMemberOf/microsoft.graph.directoryRole?$select=displayName' `
    --query 'value[*].displayName' `
    -o tsv 2>$null
  if ($LASTEXITCODE -eq 0) {
    $_activeRoles = @(
      $_entraRaw -split "`n" |
      Where-Object { $_ } |
      ForEach-Object { $_.Trim() }
    )
    $_hasAppRegRole = @($_activeRoles | Where-Object { $_appRegRoles -contains $_ })
    $_hasAssignmentRole = @($_activeRoles | Where-Object { $_assignmentRoles -contains $_ })
    if (-not $_hasAppRegRole) {
      Write-Host '  ! Missing: Cloud Application Administrator, Application Administrator,'
      Write-Host '    or Global Administrator — required to create/update the App Registration.'
      Write-Host '    Bicep will fail without this role. Activate via PIM before re-running:'
      Write-Host '    https://entra.microsoft.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/aadRoles'
    }
    if (-not $_skipRoleAssignments -and -not $_hasAssignmentRole) {
      Write-Host '  ! Missing: Privileged Role Administrator (or Global Administrator) —'
      Write-Host '    needed to assign Graph app roles to the Managed Identity.'
      Write-Host '    Either activate the role via PIM, or defer the assignments:'
      Write-Host '      azd env set AZURE_SKIP_GRAPH_ROLE_ASSIGNMENTS true'
      Write-Host '    Then run setup-graph-permissions.ps1 after deployment.'
      Write-Host '    PIM → My roles → Entra roles:'
      Write-Host '    https://entra.microsoft.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/aadRoles'
    }
    if ($_hasAppRegRole) {
      if ($_skipRoleAssignments) {
        Write-Host "  + Entra role: $($_hasAppRegRole[0]) — App Registration management covered."
        Write-Host '    Graph role assignments: deferred to setup-graph-permissions.ps1.'
      }
      elseif ($_hasAssignmentRole) {
        if ($_hasAppRegRole[0] -eq $_hasAssignmentRole[0]) {
          Write-Host "  + Entra role: $($_hasAppRegRole[0]) — covers both required permissions."
        }
        else {
          Write-Host "  + Entra roles: $($_hasAppRegRole[0]) + $($_hasAssignmentRole[0]) — both required roles active."
        }
      }
    }
  }
  else {
    Write-Host '  ! Entra role check returned an error — continuing anyway.'
    Write-Host '    Required: Cloud Application Administrator (or similar).'
    if (-not $_skipRoleAssignments) {
      Write-Host '    Also required: Privileged Role Administrator (or Global Administrator).'
      Write-Host '    To defer Graph role assignments: azd env set AZURE_SKIP_GRAPH_ROLE_ASSIGNMENTS true'
    }
  }
}
catch {
  Write-Host '  ! Entra role check encountered an error — continuing anyway.'
  Write-Host '    Required: Cloud Application Administrator (or similar).'
  if (-not $_skipRoleAssignments) {
    Write-Host '    Also required: Privileged Role Administrator (or Global Administrator).'
    Write-Host '    To defer Graph role assignments: azd env set AZURE_SKIP_GRAPH_ROLE_ASSIGNMENTS true'
  }
}
Write-Host ''
