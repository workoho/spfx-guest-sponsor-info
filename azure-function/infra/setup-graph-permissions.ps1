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
param(
  [Parameter(Mandatory)][string]$ManagedIdentityObjectId,
  [Parameter(Mandatory)][string]$TenantId,
  [Parameter(Mandatory)][string]$FunctionAppClientId
)

$ErrorActionPreference = 'Stop'

# Dot-source callout box helpers when running from a local clone.
# When executed via iwr (remote one-liner), $PSScriptRoot is empty — fall
# back to plain Write-Host stubs.
$calloutFile = Join-Path $PSScriptRoot 'Write-Callout.ps1'
if ($PSScriptRoot -and (Test-Path $calloutFile)) {
  . $calloutFile
}
else {
  function Write-Hint { param([Parameter(ValueFromRemainingArguments)][string[]]$L) Write-Host ''; foreach ($l in $L) { if ($l) { Write-Host "  $l" } }; Write-Host '' }
  function Write-NextSteps { param([Parameter(ValueFromRemainingArguments)][string[]]$L) Write-Host ''; foreach ($l in $L) { if ($l) { Write-Host "  $l" } }; Write-Host '' }
  function Write-Important { param([Parameter(ValueFromRemainingArguments)][string[]]$L) Write-Host ''; foreach ($l in $L) { if ($l) { Write-Host "  $l" -ForegroundColor Yellow } }; Write-Host '' }
}

# Ensure Microsoft.Graph module is available.
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
  Write-Host "Installing Microsoft.Graph.Authentication module..." -ForegroundColor Cyan
  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
}
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Applications)) {
  Write-Host "Installing Microsoft.Graph.Applications module..." -ForegroundColor Cyan
  Install-Module Microsoft.Graph.Applications -Scope CurrentUser -Force
}

Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Applications

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
      Write-Host "  ⚠ $($role.Name) is not available as an Application permission in this tenant (Microsoft Teams may not be licensed). Skipping — sponsors will be shown without presence status." -ForegroundColor Yellow
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
    Write-Host "  ✓ $($role.Name) assigned." -ForegroundColor Green
    $assignedRoles += $role.Name
  }
  catch {
    if ($_.Exception.Message -like "*Permission being assigned already exists*") {
      Write-Host "  ✓ $($role.Name) already assigned — skipping." -ForegroundColor Yellow
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
  Write-Host "  ⚠ Could not resolve SP by display name — falling back to known app IDs: $($spWebClientAppIds -join ', ')" -ForegroundColor Yellow
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
  Write-Host "  ✓ Identifier URI set." -ForegroundColor Green
}
else {
  Write-Host "  ✓ Identifier URI already set." -ForegroundColor Yellow
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
  Write-Host "  ✓ 'user_impersonation' scope added (id: $($existingScope.Id))." -ForegroundColor Green
}
else {
  Write-Host "  ✓ 'user_impersonation' scope already exists (id: $($existingScope.Id))." -ForegroundColor Yellow
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
      Write-Host "  ✓ $spAppId pre-authorized." -ForegroundColor Green
    }
    catch {
      if ($_.Exception.Message -like "*cannot be found*") {
        Write-Host "  ⚠ $spAppId not found in Microsoft's app registry — skipping." -ForegroundColor Yellow
      }
      else {
        throw
      }
    }
  }
  else {
    Write-Host "  ✓ $spAppId already pre-authorized." -ForegroundColor Yellow
  }
}

# Ensure appRoleAssignmentRequired is false on the Service Principal (Enterprise App).
# Normally created on first user sign-in, but since we run this script before any user
# has consented, we create it explicitly here.
$sp = Get-MgServicePrincipal -Filter "appId eq '$FunctionAppClientId'" -ErrorAction SilentlyContinue
if (-not $sp) {
  Write-Host "  Service Principal not found — creating it now (no user has signed in yet)..." -ForegroundColor Cyan
  $sp = New-MgServicePrincipal -AppId $FunctionAppClientId -ErrorAction Stop
  Write-Host "  ✓ Service Principal created (Object ID: $($sp.Id))." -ForegroundColor Green
}
else {
  Write-Host "  ✓ Service Principal already exists (Object ID: $($sp.Id))." -ForegroundColor Yellow
}

# appRoleAssignmentRequired=false: all users (including guests) can acquire tokens without
# individual assignment — even with pre-authorization in place.
if ($sp.AppRoleAssignmentRequired) {
  Write-Host "  Disabling appRoleAssignmentRequired on the Enterprise App (was: true) ..." -ForegroundColor Cyan
  Update-MgServicePrincipal -ServicePrincipalId $sp.Id -AppRoleAssignmentRequired:$false -ErrorAction Stop
  Write-Host "  ✓ appRoleAssignmentRequired set to false." -ForegroundColor Green
}
else {
  Write-Host "  ✓ appRoleAssignmentRequired is already false — no user assignment needed." -ForegroundColor Yellow
}

# Hide from My Apps portal (tags: HideApp). This is a backend auth proxy — it should not
# appear as a launchable app in users' My Apps page.
$hasHideApp = $sp.Tags -contains 'HideApp'
if (-not $hasHideApp) {
  Write-Host "  Hiding Enterprise App from My Apps portal (visible to users: No) ..." -ForegroundColor Cyan
  $updatedTags = @($sp.Tags) + @('HideApp')
  Update-MgServicePrincipal -ServicePrincipalId $sp.Id -Tags $updatedTags -ErrorAction Stop
  Write-Host "  ✓ Enterprise App hidden from My Apps portal." -ForegroundColor Green
}
else {
  Write-Host "  ✓ Enterprise App is already hidden from My Apps portal." -ForegroundColor Yellow
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
  Write-Host "  ✓ Description set." -ForegroundColor Green
}
else {
  Write-Host "  ✓ Enterprise App description already set." -ForegroundColor Yellow
}

if ($sp.Notes -ne $spNotes) {
  Write-Host "  Setting Enterprise App notes ..." -ForegroundColor Cyan
  Update-MgServicePrincipal -ServicePrincipalId $sp.Id `
    -Notes $spNotes -ErrorAction Stop
  Write-Host "  ✓ Notes set." -ForegroundColor Green
}
else {
  Write-Host "  ✓ Enterprise App notes already set." -ForegroundColor Yellow
}
# Service Management Reference — shown under Enterprise App → Properties.
# Points to the GitHub Issues tracker so Ops teams know where to file tickets.
$desiredSmRef = 'https://github.com/workoho/spfx-guest-sponsor-info/issues'
if ($sp.ServiceManagementReference -ne $desiredSmRef) {
  Write-Host "  Setting Service Management Reference ..." -ForegroundColor Cyan
  Update-MgServicePrincipal -ServicePrincipalId $sp.Id `
    -ServiceManagementReference $desiredSmRef -ErrorAction Stop
  Write-Host "  ✓ Service Management Reference set." -ForegroundColor Green
}
else {
  Write-Host "  ✓ Service Management Reference already set." -ForegroundColor Yellow
}

# Homepage URL — visible under Enterprise App → Properties.
$desiredHomepage = 'https://github.com/workoho/spfx-guest-sponsor-info'
if ($sp.Homepage -ne $desiredHomepage) {
  Write-Host "  Setting Enterprise App homepage URL ..." -ForegroundColor Cyan
  Update-MgServicePrincipal -ServicePrincipalId $sp.Id `
    -Homepage $desiredHomepage -ErrorAction Stop
  Write-Host "  ✓ Homepage URL set." -ForegroundColor Green
}
else {
  Write-Host "  ✓ Enterprise App homepage URL already set." -ForegroundColor Yellow
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
Write-NextSteps @summaryLines
