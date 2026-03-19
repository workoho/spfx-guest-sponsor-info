targetScope = 'resourceGroup'

metadata name = 'Guest Sponsor Info – Azure Function Proxy'
metadata description = 'Deploys an Azure Function App that acts as a Graph API proxy for the Guest Sponsor Info SharePoint web part. Includes a Storage Account, App Service Plan, EasyAuth configuration, Managed Identity role assignments, Log Analytics Workspace, and Application Insights.'

@description('Azure region for all resources.')
param location string = resourceGroup().location

@description('Entra tenant ID (GUID).')
param tenantId string

@description('Tenant name without domain suffix, e.g. "contoso".')
param tenantName string

@description('Globally unique name for the Function App (2–60 characters, letters, numbers, and hyphens only).')
@minLength(2)
@maxLength(60)
param functionAppName string

@description('Client ID of the App Registration created for EasyAuth.')
param functionClientId string

@description('URL to the pre-built function ZIP package (GitHub Release asset).')
param packageUrl string = 'https://github.com/jpawlowski/spfx-guest-sponsor-info/releases/latest/download/guest-sponsor-info-function.zip'

@description('Resource tags to apply to all deployed resources.')
param tags object = {}

var storageAccountName = toLower(replace(functionAppName, '-', ''))
var appServicePlanName = '${functionAppName}-plan'
var logAnalyticsWorkspaceName = '${functionAppName}-logs'
var appInsightsName = '${functionAppName}-insights'

// ── Storage Account (required by Azure Functions runtime) ────────────────────
resource storageAccount 'Microsoft.Storage/storageAccounts@2023-01-01' = {
  name: length(storageAccountName) > 24 ? substring(storageAccountName, 0, 24) : storageAccountName
  location: location
  tags: tags
  sku: {
    name: 'Standard_LRS'
  }
  kind: 'StorageV2'
  properties: {
    supportsHttpsTrafficOnly: true
    minimumTlsVersion: 'TLS1_2'
    allowBlobPublicAccess: false
    allowSharedKeyAccess: false
  }
}

// ── Log Analytics Workspace ──────────────────────────────────────────────────
// Backend for Application Insights. Workspace-based AppInsights is the modern
// approach (classic components are deprecated).
resource logAnalyticsWorkspace 'Microsoft.OperationalInsights/workspaces@2022-10-01' = {
  name: logAnalyticsWorkspaceName
  location: location
  tags: tags
  properties: {
    sku: {
      name: 'PerGB2018'
    }
    // 30 days is the minimum; first 5 GB/month per workspace is free.
    retentionInDays: 30
  }
}

// ── Application Insights ─────────────────────────────────────────────────────
// When APPLICATIONINSIGHTS_CONNECTION_STRING is set, the Azure Functions Node.js
// runtime instruments automatically — no code changes needed. Captured data:
//   • invocations as "requests" (duration, success, HTTP status)
//   • outbound Graph API calls as "dependencies" (URL, latency, status)
//   • context.log/warn/error() as "traces" (incl. Graph requestId)
//   • unhandled exceptions as "exceptions"
resource appInsights 'Microsoft.Insights/components@2020-02-02' = {
  name: appInsightsName
  location: location
  tags: tags
  kind: 'web'
  properties: {
    Application_Type: 'web'
    WorkspaceResourceId: logAnalyticsWorkspace.id
    RetentionInDays: 30
    IngestionMode: 'LogAnalytics'
    publicNetworkAccessForIngestion: 'Enabled'
    publicNetworkAccessForQuery: 'Enabled'
  }
}

// ── Consumption App Service Plan ─────────────────────────────────────────────
resource appServicePlan 'Microsoft.Web/serverfarms@2023-01-01' = {
  name: appServicePlanName
  location: location
  tags: tags
  sku: {
    name: 'Y1'
    tier: 'Dynamic'
  }
  properties: {}
}

// ── Function App ─────────────────────────────────────────────────────────────
resource functionApp 'Microsoft.Web/sites@2023-01-01' = {
  name: functionAppName
  location: location
  tags: tags
  kind: 'functionapp'
  identity: {
    type: 'SystemAssigned'
  }
  properties: {
    serverFarmId: appServicePlan.id
    httpsOnly: true
    siteConfig: {
      appSettings: [
        // Identity-based storage connection — no account key stored anywhere.
        // The three role assignments below grant the Managed Identity the minimum
        // required access to blob, queue, and table services.
        {
          name: 'AzureWebJobsStorage__accountName'
          value: storageAccount.name
        }
        {
          name: 'AzureWebJobsStorage__credential'
          value: 'managedidentity'
        }
        {
          name: 'FUNCTIONS_EXTENSION_VERSION'
          value: '~4'
        }
        {
          name: 'FUNCTIONS_WORKER_RUNTIME'
          value: 'node'
        }
        {
          name: 'WEBSITE_NODE_DEFAULT_VERSION'
          value: '~22'
        }
        {
          name: 'WEBSITE_RUN_FROM_PACKAGE'
          value: packageUrl
        }
        {
          name: 'TENANT_ID'
          value: tenantId
        }
        {
          name: 'ALLOWED_AUDIENCE'
          value: 'api://guest-sponsor-info-proxy/${functionClientId}'
        }
        {
          name: 'CORS_ALLOWED_ORIGIN'
          value: 'https://${tenantName}.sharepoint.com'
        }
        {
          name: 'SPONSOR_LOOKUP_TIMEOUT_MS'
          value: '5000'
        }
        {
          name: 'BATCH_TIMEOUT_MS'
          value: '4000'
        }
        {
          name: 'PRESENCE_TIMEOUT_MS'
          value: '2500'
        }
        {
          name: 'NODE_ENV'
          value: 'production'
        }
        {
          // When set, the Azure Functions Node.js runtime automatically sends all
          // invocation traces, outbound dependencies, and exceptions to AppInsights.
          // No application code changes are required.
          name: 'APPLICATIONINSIGHTS_CONNECTION_STRING'
          value: appInsights.properties.ConnectionString
        }
      ]
      cors: {
        allowedOrigins: [
          'https://${tenantName}.sharepoint.com'
        ]
        supportCredentials: false
      }
    }
  }
}

// ── EasyAuth – Microsoft Entra ID provider ───────────────────────────────────
resource authSettings 'Microsoft.Web/sites/config@2023-01-01' = {
  name: 'authsettingsV2'
  parent: functionApp
  properties: {
    globalValidation: {
      requireAuthentication: true
      unauthenticatedClientAction: 'Return401'
    }
    identityProviders: {
      azureActiveDirectory: {
        enabled: true
        registration: {
          clientId: functionClientId
          openIdIssuer: 'https://sts.windows.net/${tenantId}/'
        }
        validation: {
          allowedAudiences: [
            'api://guest-sponsor-info-proxy/${functionClientId}'
          ]
        }
      }
    }
    login: {
      tokenStore: {
        enabled: false
      }
    }
  }
}

// ── Storage role assignments (Managed Identity auth, no key required) ────────
// The Consumption plan runtime needs blob, queue, and table access.
// 'Owner' (or a custom role with Microsoft.Authorization/roleAssignments/write)
// is required on the deploying principal to create these assignments.
var storageBlobDataOwnerRoleId = 'b7e6dc6d-f1e8-4753-8033-0f276bb0955b'
var storageQueueDataContributorRoleId = '974c5e8b-45b9-4653-ba55-5f855dd0fb88'
var storageTableDataContributorRoleId = '0a9a7e1f-b9d0-4cc4-a60d-0319b160aaa3'

resource storageBlobRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  scope: storageAccount
  name: guid(storageAccount.id, functionApp.id, storageBlobDataOwnerRoleId)
  properties: {
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', storageBlobDataOwnerRoleId)
    principalId: functionApp.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

resource storageQueueRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  scope: storageAccount
  name: guid(storageAccount.id, functionApp.id, storageQueueDataContributorRoleId)
  properties: {
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', storageQueueDataContributorRoleId)
    principalId: functionApp.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

resource storageTableRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  scope: storageAccount
  name: guid(storageAccount.id, functionApp.id, storageTableDataContributorRoleId)
  properties: {
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', storageTableDataContributorRoleId)
    principalId: functionApp.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

// ── Outputs ──────────────────────────────────────────────────────────────────
@description('The URL of the deployed Function App.')
output functionAppUrl string = 'https://${functionApp.properties.defaultHostName}'

@description('The function endpoint URL to paste into the SPFx web part property pane.')
output sponsorApiUrl string = 'https://${functionApp.properties.defaultHostName}/api/getGuestSponsors'

@description('Object ID of the system-assigned Managed Identity — needed for setup-graph-permissions.ps1.')
output managedIdentityObjectId string = functionApp.identity.principalId

@description('Name of the Application Insights component — open in the Azure Portal for live telemetry.')
output appInsightsName string = appInsights.name
