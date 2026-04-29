// SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
// SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0
//
// Deploys the Azure Functions hosting stack:
//   - Storage Account (identity-based access, no keys)
//   - Blob container + deployment script for Flex Consumption package upload
//   - App Service Plan (Consumption Y1 or Flex Consumption FC1)
//   - Function App (including EasyAuth and all storage role assignments)
//
// Outputs the Managed Identity principalId and the Function App URL so the
// calling template can wire up Graph permissions and expose web part config.

@description('Azure region for all resources.')
param location string

@description('Name of the Function App. Must be globally unique across Azure.')
@minLength(2)
@maxLength(58)
param functionAppName string

@description('Hosting plan: "Consumption" (Y1/Dynamic, free tier) or "FlexConsumption" (FC1/Linux, reduced cold starts).')
@allowed(['Consumption', 'FlexConsumption'])
param hostingPlan string

@description('Number of always-ready instances (Flex Consumption only). 0 = on-demand; 1 = no cold starts.')
@minValue(0)
param alwaysReadyInstances int

@description('Maximum scale-out instance count (Flex Consumption only).')
@minValue(1)
@maxValue(1000)
param maximumFlexInstances int

@description('Memory per instance in MB (Flex Consumption only). 512 or 2048.')
@allowed([512, 2048])
param instanceMemoryMB int

@description('Daily memory-time budget in GB-seconds (Consumption only). 0 = unlimited.')
@minValue(0)
param dailyMemoryTimeQuotaGBs int

@description('Resolved function package ZIP URL (already normalised by the caller).')
param resolvedPackageUrl string

@description('Function package version label — written to the APP_VERSION app setting for telemetry.')
param appVersion string

@description('Entra tenant ID (GUID) — used in the EasyAuth issuer URL.')
param tenantId string

@description('SharePoint tenant name without domain suffix, e.g. "contoso" — used for CORS.')
param tenantName string

@description('Client ID (appId) of the EasyAuth App Registration.')
param appClientId string

@description('Application Insights connection string. Pass an empty string when monitoring is disabled.')
param appInsightsConnectionString string

@description('Resource tags to apply to all resources in this module.')
param tags object

// ── Derived values ────────────────────────────────────────────────────────────

var isFlexConsumption = hostingPlan == 'FlexConsumption'
var appServicePlanName = '${functionAppName}-plan'
var deploymentContainerName = 'app-package'
// Storage account names: lowercase, no hyphens, max 24 chars.
var rawStorageAccountName = toLower(replace(functionAppName, '-', ''))
var storageAccountName = length(rawStorageAccountName) > 24
  ? substring(rawStorageAccountName, 0, 24)
  : rawStorageAccountName

// ── Storage Account ───────────────────────────────────────────────────────────
// Identity-based access only — no shared keys, no connection strings in app settings.
resource storageAccount 'Microsoft.Storage/storageAccounts@2023-01-01' = {
  name: storageAccountName
  location: location
  tags: tags
  sku: { name: 'Standard_LRS' }
  kind: 'StorageV2'
  properties: {
    supportsHttpsTrafficOnly: true
    minimumTlsVersion: 'TLS1_2'
    allowBlobPublicAccess: false
    allowSharedKeyAccess: false
  }
}

// ── Flex Consumption: blob container for deployment package ───────────────────
// Flex Consumption cannot pull a ZIP from a remote URL; the package must live in
// a blob container. The container is created here; the deployment script below
// uploads the actual ZIP during ARM provisioning.
resource blobService 'Microsoft.Storage/storageAccounts/blobServices@2023-01-01' = if (isFlexConsumption) {
  parent: storageAccount
  name: 'default'
}

resource deploymentContainer 'Microsoft.Storage/storageAccounts/blobServices/containers@2023-01-01' = if (isFlexConsumption) {
  parent: blobService
  name: deploymentContainerName
  properties: { publicAccess: 'None' }
}

// ── Flex Consumption: deployment script (ZIP upload) ─────────────────────────
// An Azure CLI container script runs once per unique packageUrl during ARM
// provisioning: downloads the function ZIP and uploads it to the blob container.
// forceUpdateTag = hash of the URL so the script re-runs only when the URL changes.
//
// Deployment scripts require a User-Assigned Managed Identity — System-Assigned
// is not supported by the deploymentScripts resource type.
resource deployScriptIdentity 'Microsoft.ManagedIdentity/userAssignedIdentities@2023-01-31' = if (isFlexConsumption) {
  name: '${functionAppName}-deploy-id'
  location: location
  tags: tags
}

// Grant the deployment script identity Storage Blob Data Contributor on the
// storage account so it can upload the ZIP. Full owner access is not required.
resource deployScriptBlobRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (isFlexConsumption) {
  scope: storageAccount
  // Deterministic GUID: account + identity name + Storage Blob Data Contributor role ID.
  name: guid(storageAccount.id, '${functionAppName}-deploy-id', 'ba92f5b4-2d11-453d-a403-e96b0029c9fe')
  properties: {
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      'ba92f5b4-2d11-453d-a403-e96b0029c9fe' // Storage Blob Data Contributor
    )
    // BCP318: safe — deployScriptIdentity is always deployed when isFlexConsumption.
    #disable-next-line BCP318
    principalId: deployScriptIdentity.properties.principalId
    principalType: 'ServicePrincipal'
  }
}

// Extracted to avoid BCP318 inside the resource body.
// Safe: shares the same isFlexConsumption condition as deployZipScript.
#disable-next-line BCP318
var deployScriptIdentityResourceId = deployScriptIdentity.id

resource deployZipScript 'Microsoft.Resources/deploymentScripts@2023-08-01' = if (isFlexConsumption) {
  name: '${functionAppName}-deploy-zip'
  location: location
  kind: 'AzureCLI'
  tags: tags
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: {
      '${deployScriptIdentityResourceId}': {}
    }
  }
  dependsOn: [deployScriptBlobRole, deploymentContainer]
  properties: {
    azCliVersion: '2.60.0'
    retentionInterval: 'PT2H'
    forceUpdateTag: uniqueString(resolvedPackageUrl)
    environmentVariables: [
      { name: 'STORAGE_ACCOUNT', value: storageAccount.name }
      { name: 'CONTAINER', value: deploymentContainerName }
      { name: 'PACKAGE_URL', value: resolvedPackageUrl }
    ]
    // Retry loop handles RBAC propagation delay (~30-60 s after role assignment).
    scriptContent: '''
      set -euo pipefail
      curl -sSfL -o /tmp/function.zip "$PACKAGE_URL"
      for attempt in 1 2 3; do
        az storage blob upload \
          --account-name "$STORAGE_ACCOUNT" \
          --container-name "$CONTAINER" \
          --name function.zip \
          --file /tmp/function.zip \
          --auth-mode login \
          --overwrite && break || {
            echo "Upload attempt $attempt failed — waiting for RBAC propagation"
            sleep 30
          }
      done
    '''
  }
}

// ── App Service Plan ──────────────────────────────────────────────────────────
// Exactly one of the two plan resources is deployed depending on hostingPlan.

// Consumption (Y1 / Dynamic) — includes the Azure free grant.
resource appServicePlanConsumption 'Microsoft.Web/serverfarms@2023-01-01' = if (!isFlexConsumption) {
  name: appServicePlanName
  location: location
  tags: tags
  sku: { name: 'Y1', tier: 'Dynamic' }
  properties: {}
}

// Flex Consumption (FC1 / Linux) — requires API version 2023-12-01 or later.
// Not available in all Azure regions: https://aka.ms/flex-region
resource appServicePlanFlex 'Microsoft.Web/serverfarms@2023-12-01' = if (isFlexConsumption) {
  name: appServicePlanName
  location: location
  tags: tags
  kind: 'functionapp'
  sku: { name: 'FC1', tier: 'FlexConsumption' }
  properties: { reserved: true } // Linux
}

// ── App settings ──────────────────────────────────────────────────────────────

var monitoringAppSettings = !empty(appInsightsConnectionString)
  ? [{ name: 'APPLICATIONINSIGHTS_CONNECTION_STRING', value: appInsightsConnectionString }]
  : []

// Shared across both hosting plans.
var sharedAppSettings = [
  // Identity-based storage connection — no account key stored anywhere.
  { name: 'AzureWebJobsStorage__accountName', value: storageAccount.name }
  { name: 'AzureWebJobsStorage__credential', value: 'managedidentity' }
  { name: 'TENANT_ID', value: tenantId }
  { name: 'ALLOWED_AUDIENCE', value: appClientId }
  { name: 'CORS_ALLOWED_ORIGIN', value: 'https://${tenantName}.sharepoint.com' }
  { name: 'SPONSOR_LOOKUP_TIMEOUT_MS', value: '5000' }
  { name: 'BATCH_TIMEOUT_MS', value: '4000' }
  { name: 'PRESENCE_TIMEOUT_MS', value: '2500' }
  { name: 'NODE_ENV', value: 'production' }
  { name: 'APP_VERSION', value: appVersion }
]

var effectiveSharedAppSettings = concat(sharedAppSettings, monitoringAppSettings)

// Consumption-only settings: Flex Consumption uses functionAppConfig instead.
var consumptionOnlyAppSettings = [
  { name: 'FUNCTIONS_EXTENSION_VERSION', value: '~4' }
  { name: 'FUNCTIONS_WORKER_RUNTIME', value: 'node' }
  { name: 'WEBSITE_NODE_DEFAULT_VERSION', value: '~22' }
  { name: 'WEBSITE_RUN_FROM_PACKAGE', value: resolvedPackageUrl }
]

// EasyAuth configuration — identical for both hosting plans.
var easyAuthProperties = {
  globalValidation: {
    requireAuthentication: true
    unauthenticatedClientAction: 'Return401'
  }
  identityProviders: {
    azureActiveDirectory: {
      enabled: true
      registration: {
        clientId: appClientId
        openIdIssuer: 'https://sts.windows.net/${tenantId}/'
      }
      validation: {
        allowedAudiences: [appClientId]
      }
    }
  }
  login: {
    tokenStore: { enabled: false }
  }
}

// ── Function App — Consumption ────────────────────────────────────────────────
resource functionApp 'Microsoft.Web/sites@2023-01-01' = if (!isFlexConsumption) {
  name: functionAppName
  location: location
  tags: tags
  kind: 'functionapp'
  identity: { type: 'SystemAssigned' }
  properties: {
    serverFarmId: appServicePlanConsumption.id
    httpsOnly: true
    // Suspend the app when the daily GB-second budget is exhausted (cost guard).
    // Set dailyMemoryTimeQuotaGBs=0 to disable (not recommended).
    dailyMemoryTimeQuota: dailyMemoryTimeQuotaGBs
    siteConfig: {
      appSettings: concat(effectiveSharedAppSettings, consumptionOnlyAppSettings)
      cors: {
        allowedOrigins: ['https://${tenantName}.sharepoint.com']
        supportCredentials: false
      }
    }
  }
}

resource authSettings 'Microsoft.Web/sites/config@2023-01-01' = if (!isFlexConsumption) {
  name: 'authsettingsV2'
  parent: functionApp
  properties: easyAuthProperties
}

// ── Function App — Flex Consumption ──────────────────────────────────────────
resource functionAppFlex 'Microsoft.Web/sites@2023-12-01' = if (isFlexConsumption) {
  name: functionAppName
  location: location
  tags: tags
  kind: 'functionapp,linux'
  identity: { type: 'SystemAssigned' }
  dependsOn: [deploymentContainer]
  properties: {
    serverFarmId: appServicePlanFlex.id
    httpsOnly: true
    functionAppConfig: {
      runtime: { name: 'node', version: '22' }
      scaleAndConcurrency: {
        maximumInstanceCount: maximumFlexInstances
        instanceMemoryMB: instanceMemoryMB
        alwaysReady: alwaysReadyInstances > 0 ? [{ name: 'http', instanceCount: alwaysReadyInstances }] : []
      }
      deployment: {
        storage: {
          type: 'blobContainer'
          value: '${storageAccount.properties.primaryEndpoints.blob}${deploymentContainerName}'
          authentication: { type: 'SystemAssignedIdentity' }
        }
      }
    }
    siteConfig: {
      appSettings: effectiveSharedAppSettings
      cors: {
        allowedOrigins: ['https://${tenantName}.sharepoint.com']
        supportCredentials: false
      }
    }
  }
}

resource authSettingsFlex 'Microsoft.Web/sites/config@2023-12-01' = if (isFlexConsumption) {
  name: 'authsettingsV2'
  parent: functionAppFlex
  properties: easyAuthProperties
}

// ── Storage role assignments (identity-based, no key) ────────────────────────
// The current app uses only HTTP and timer triggers. With identity-based
// AzureWebJobsStorage, the Functions host needs blob access for timer locks and
// host artifacts. We also keep table access so host diagnostic events can still
// be persisted. Queue access is intentionally omitted because no queue- or
// blob-triggered workloads are deployed in this app.
//
// Each role is declared twice (Consumption / Flex) with mutually exclusive
// conditions because Bicep cannot reference a conditionally deployed resource
// outside its own condition branch without triggering BCP318.
//
// roleDefinition IDs:
//   b7e6dc6d-f1e8-4753-8033-0f276bb0955b  Storage Blob Data Owner
//   0a9a7e1f-b9d0-4cc4-a60d-0319b160aaa3  Storage Table Data Contributor

// BCP318 suppression: safe — variables are only accessed inside matching condition branches.
#disable-next-line BCP318
var consumptionPrincipalId = functionApp.identity.principalId
#disable-next-line BCP318
var flexPrincipalId = functionAppFlex.identity.principalId

var functionAppResourceId = resourceId('Microsoft.Web/sites', functionAppName)

resource storageBlobRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (!isFlexConsumption) {
  scope: storageAccount
  name: guid(storageAccount.id, functionAppResourceId, 'b7e6dc6d-f1e8-4753-8033-0f276bb0955b')
  properties: {
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      'b7e6dc6d-f1e8-4753-8033-0f276bb0955b'
    )
    #disable-next-line BCP318
    principalId: consumptionPrincipalId
    principalType: 'ServicePrincipal'
  }
}

resource storageTableRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (!isFlexConsumption) {
  scope: storageAccount
  name: guid(storageAccount.id, functionAppResourceId, '0a9a7e1f-b9d0-4cc4-a60d-0319b160aaa3')
  properties: {
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      '0a9a7e1f-b9d0-4cc4-a60d-0319b160aaa3'
    )
    #disable-next-line BCP318
    principalId: consumptionPrincipalId
    principalType: 'ServicePrincipal'
  }
}

resource storageBlobRoleFlex 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (isFlexConsumption) {
  scope: storageAccount
  name: guid(storageAccount.id, functionAppResourceId, 'b7e6dc6d-f1e8-4753-8033-0f276bb0955b')
  properties: {
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      'b7e6dc6d-f1e8-4753-8033-0f276bb0955b'
    )
    #disable-next-line BCP318
    principalId: flexPrincipalId
    principalType: 'ServicePrincipal'
  }
}

resource storageTableRoleFlex 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (isFlexConsumption) {
  scope: storageAccount
  name: guid(storageAccount.id, functionAppResourceId, '0a9a7e1f-b9d0-4cc4-a60d-0319b160aaa3')
  properties: {
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      '0a9a7e1f-b9d0-4cc4-a60d-0319b160aaa3'
    )
    #disable-next-line BCP318
    principalId: flexPrincipalId
    principalType: 'ServicePrincipal'
  }
}

// ── Outputs ───────────────────────────────────────────────────────────────────

@description('Object ID of the system-assigned Managed Identity — needed for Graph role assignments and setup-graph-permissions.ps1.')
output managedIdentityObjectId string = isFlexConsumption ? flexPrincipalId : consumptionPrincipalId

@description('The Function App name.')
output functionAppName string = functionAppName

@description('Base URL of the Function App.')
output functionAppUrl string = 'https://${functionAppName}.azurewebsites.net'

@description('Full endpoint URL for the getGuestSponsors function.')
output sponsorApiEndpointUrl string = 'https://${functionAppName}.azurewebsites.net/api/getGuestSponsors'

@description('Name of the Storage Account used by the Functions runtime.')
output deploymentStorageAccountName string = storageAccount.name

@description('The selected hosting plan.')
output hostingPlan string = hostingPlan
