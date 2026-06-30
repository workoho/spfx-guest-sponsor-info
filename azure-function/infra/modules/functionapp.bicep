// SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
// SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0
//
// Deploys the Azure Functions hosting stack:
//   - Storage Account (identity-based access, no keys)
//   - Blob container + native OneDeploy publishing for Flex Consumption
//   - App Service Plan (FC1 / Flex Consumption, Linux)
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

@description('Number of always-ready instances. 0 = on-demand (cold starts possible); 1 = one instance kept warm.')
@minValue(0)
param alwaysReadyInstances int

@description('Maximum scale-out instance count (Flex Consumption only).')
@minValue(1)
@maxValue(1000)
param maximumFlexInstances int

@description('Memory per instance in MB. 512 or 2048.')
@allowed([512, 2048])
param instanceMemoryMB int

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

// ── Blob container for deployment package ────────────────────────────────────
// Flex Consumption cannot pull a ZIP from a remote URL; the package must live in
// a blob container. The container is created here; the native OneDeploy
// extension publishes the actual ZIP into this container during deployment.
resource blobService 'Microsoft.Storage/storageAccounts/blobServices@2023-01-01' = {
  parent: storageAccount
  name: 'default'
}

resource deploymentContainer 'Microsoft.Storage/storageAccounts/blobServices/containers@2023-01-01' = {
  parent: blobService
  name: deploymentContainerName
  properties: { publicAccess: 'None' }
}

// ── App Service Plan ──────────────────────────────────────────────────────────
// FC1 / Flex Consumption (Linux) — requires API version 2023-12-01 or later.
// Not available in all Azure regions: https://aka.ms/flex-region
resource appServicePlan 'Microsoft.Web/serverfarms@2023-12-01' = {
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

// EasyAuth configuration.
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

// ── Function App ─────────────────────────────────────────────────────────────
resource functionApp 'Microsoft.Web/sites@2023-12-01' = {
  name: functionAppName
  location: location
  tags: tags
  kind: 'functionapp,linux'
  identity: { type: 'SystemAssigned' }
  dependsOn: [deploymentContainer]
  properties: {
    serverFarmId: appServicePlan.id
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
      http20Enabled: true
      minTlsVersion: '1.2'
      // Require ECDHE (forward secrecy) + AES-GCM (AEAD). Drops legacy RSA
      // key-exchange and CBC-mode suites. Safe because all clients are modern
      // browsers via SharePoint; HTTP/2 (RFC 7540) also mandates GCM suites.
      minTlsCipherSuite: 'TLS_ECDHE_RSA_WITH_AES_128_GCM_SHA256'
      scmMinTlsVersion: '1.2'
      // Deployment runs via native OneDeploy — FTP is never used.
      ftpsState: 'Disabled'
      cors: {
        allowedOrigins: ['https://${tenantName}.sharepoint.com']
        supportCredentials: false
      }
    }
  }
}

resource authSettings 'Microsoft.Web/sites/config@2023-12-01' = {
  name: 'authsettingsV2'
  parent: functionApp
  properties: easyAuthProperties
}

// ── Storage role assignments (identity-based, no key) ────────────────────────
// The current app uses only HTTP and timer triggers. With identity-based
// AzureWebJobsStorage, the Functions host needs blob access for timer locks,
// host artifacts, and the Flex deployment container. We also keep table access
// so host diagnostic events can still be persisted. Queue access is
// intentionally omitted because no queue- or blob-triggered workloads are
// deployed in this app.
//
// roleDefinition IDs:
//   b7e6dc6d-f1e8-4753-8033-0f276bb0955b  Storage Blob Data Owner
//   0a9a7e1f-b9d0-4cc4-a60d-0319b160aaa3  Storage Table Data Contributor

var functionAppResourceId = resourceId('Microsoft.Web/sites', functionAppName)

resource storageBlobRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  scope: storageAccount
  name: guid(storageAccount.id, functionAppResourceId, 'b7e6dc6d-f1e8-4753-8033-0f276bb0955b')
  properties: {
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      'b7e6dc6d-f1e8-4753-8033-0f276bb0955b'
    )
    principalId: functionApp.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

resource storageTableRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  scope: storageAccount
  name: guid(storageAccount.id, functionAppResourceId, '0a9a7e1f-b9d0-4cc4-a60d-0319b160aaa3')
  properties: {
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      '0a9a7e1f-b9d0-4cc4-a60d-0319b160aaa3'
    )
    principalId: functionApp.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

// ── Native Flex package publish (OneDeploy) ─────────────────────────────────
// Flex Consumption uses OneDeploy to copy a ready-to-run ZIP from a remote URL
// into the configured deployment container. This keeps code publishing in the
// native Microsoft.Web control plane instead of a custom deployment script.
resource functionAppOneDeploy 'Microsoft.Web/sites/extensions@2022-09-01' = {
  parent: functionApp
  name: 'onedeploy'
  dependsOn: [storageBlobRole]
  // The Azure Functions Flex docs require packageUri/remoteBuild here, but the
  // current Bicep type provider does not model onedeploy's request body yet.
  #disable-next-line BCP187
  properties: {
    packageUri: resolvedPackageUrl
    remoteBuild: false
  }
}

// ── Outputs ───────────────────────────────────────────────────────────────────

@description('Object ID of the system-assigned Managed Identity — needed for Graph role assignments and setup-graph-permissions.ps1.')
output managedIdentityObjectId string = functionApp.identity.principalId

@description('The Function App name.')
output functionAppName string = functionAppName

@description('Base URL of the Function App.')
output functionAppUrl string = 'https://${functionAppName}.azurewebsites.net'

@description('Full endpoint URL for the getGuestSponsors function.')
output sponsorApiEndpointUrl string = 'https://${functionAppName}.azurewebsites.net/api/getGuestSponsors'

@description('Name of the Storage Account used by the Functions runtime.')
output deploymentStorageAccountName string = storageAccount.name
