targetScope = 'resourceGroup'

metadata name = 'Guest Sponsor Info – Azure Function Proxy'
metadata description = 'Deploys an Azure Function App that acts as a Graph API proxy for the Guest Sponsor Info SharePoint web part. Includes a Storage Account, App Service Plan, EasyAuth configuration, Managed Identity role assignments, Log Analytics Workspace, and Application Insights.'
metadata repository = 'https://github.com/jpawlowski/spfx-guest-sponsor-info'
metadata author = 'jpawlowski'
metadata license = 'MIT'

@description('Azure region for all resources.')
param location string = resourceGroup().location

@description('Entra tenant ID (GUID).')
param tenantId string

@description('Tenant name without domain suffix, e.g. "contoso".')
param tenantName string

@description('Globally unique name for the Function App (2–58 characters, letters, numbers, and hyphens only).')
@minLength(2)
@maxLength(58)
param functionAppName string

@description('Client ID of the App Registration created for EasyAuth.')
param functionClientId string

@description('Hosting plan for the Function App. "Consumption" = Y1/Dynamic (free tier included, cold starts after ~20 min idle, ZIP served directly from GitHub package URL). "FlexConsumption" = FC1/Linux-only (no free tier, cold starts greatly reduced — alwaysReadyInstances=1 eliminates them; ZIP is uploaded to blob storage by the provisioning script automatically during ARM deployment; "Deploy to Azure" button is supported). Not all Azure regions support Flex Consumption — check https://aka.ms/flex-region before choosing.')
@allowed([
  'Consumption'
  'FlexConsumption'
])
param hostingPlan string = 'Consumption'

@description('Number of always-ready (pre-warmed) instances for the Function App (Flex Consumption plan only). 0 = purely on-demand (cold starts possible). 1 = one instance kept warm — eliminates cold starts (~€2–5/month). Ignored when hostingPlan = "Consumption".')
@minValue(0)
param alwaysReadyInstances int = 1

@description('Hard upper bound on the number of instances the Flex Consumption plan may scale out to (Flex Consumption plan only). Acts as a cost ceiling — scale-out stops at this limit regardless of demand. Valid range: 1–1000. Default: 10. Ignored when hostingPlan = "Consumption".')
@minValue(1)
@maxValue(1000)
param maximumFlexInstances int = 10

@description('Memory allocated to each Flex Consumption instance in MB (Flex Consumption plan only). Valid values: 512 or 2048. Higher memory allows more concurrent requests per instance but costs more per GB-second. Default: 2048. Ignored when hostingPlan = "Consumption".')
@allowed([512, 2048])
param instanceMemoryMB int = 2048

@description('Function package version to deploy. "latest" (default) = always pull the newest GitHub Release at provisioning time. SemVer without "v" prefix, e.g. "1.4.2" = pin to that specific release. On Consumption: sets the WEBSITE_RUN_FROM_PACKAGE URL. On Flex Consumption: the provisioning script re-runs and re-uploads the ZIP whenever this value changes — set it on each redeployment to trigger a code update.')
param appVersion string = 'latest'

@description('Override the computed package ZIP URL. Leave empty (default) to auto-compute from appVersion. Use only when hosting the ZIP at a custom location.')
param packageUrl string = ''

@description('Additional resource tags to apply to all deployed resources. The tags "managed-by", "source", and "package-version" are always set automatically and cannot be overridden.')
param tags object = {}

@description('Deploy Azure Maps account for inline address map preview.')
param deployAzureMaps bool = true

@description('Optional custom Azure Maps account name. Leave empty to auto-generate.')
param azureMapsAccountName string = ''

@description('Azure region for the Azure Maps account. Must be one of the regions supported by Microsoft.Maps/accounts (westeurope, northeurope, westus2, eastus, westcentralus, global). Defaults to westeurope. Required when the resource group location is not supported by Azure Maps (e.g. germanywestcentral).')
@allowed([
  'westeurope'
  'northeurope'
  'westus2'
  'eastus'
  'westcentralus'
  'global'
])
param azureMapsLocation string = 'westeurope'

@description('Enable operational email alert for probable service outage (5xx/504 spike or low success rate).')
param enableServiceOutageAlert bool = true

@description('Enable operational email alert for auth/config regressions (AUTH_CONFIG_* reason codes).')
param enableAuthConfigRegressionAlert bool = true

@description('Enable info-only alert for likely attack/noise spikes (high 401/403 from many IPs).')
param enableLikelyAttackInfoAlert bool = true

@description('KQL alert evaluation frequency in minutes.')
@minValue(1)
param alertEvaluationFrequencyInMinutes int = 5

@description('KQL alert lookback window in minutes.')
@minValue(5)
param alertWindowInMinutes int = 15

@description('Minimum total requests in window before service outage alert can fire.')
@minValue(1)
param serviceOutageMinRequests int = 20

@description('5xx/504 count threshold for service outage alert.')
@minValue(1)
param serviceOutageFailureCountThreshold int = 10

@description('Success-rate percentage threshold below which service outage alert can fire.')
@minValue(1)
@maxValue(99)
param serviceOutageSuccessRatePercentThreshold int = 70

@description('AUTH_CONFIG_* trace count threshold for config-regression alert.')
@minValue(1)
param authConfigRegressionHitsThreshold int = 1

@description('401/403 count threshold for likely-attack info alert.')
@minValue(1)
param likelyAttackDeniedCountThreshold int = 50

@description('Unique client IP threshold for likely-attack info alert.')
@minValue(1)
param likelyAttackUniqueIpThreshold int = 20

@description('Denied-rate percentage threshold for likely-attack info alert.')
@minValue(1)
@maxValue(100)
param likelyAttackDenyRatePercentThreshold int = 80

@description('Minimum successful requests required before likely-attack info alert fires (avoid pure outage overlap).')
@minValue(0)
param likelyAttackMinSuccessThreshold int = 1

@description('Consumption plan (Y1) daily memory-time budget in GB-seconds. When hit, the Function App is suspended until midnight UTC — the primary cost guard. 0 = unlimited (not recommended). Default 10 000 GB-s ≈ 13 000 typical invocations/day, stays within the monthly free tier (400 000 GB-s/month). Ignored when hostingPlan = "FlexConsumption" (Flex has no daily GB-second budget concept).')
@minValue(0)
param dailyMemoryTimeQuotaGBs int = 10000

@description('Action group resource IDs for operational email alerts. Leave empty to create alert rules without notifications.')
param operationalActionGroupResourceIds array = []

@description('Action group resource IDs for info-only alerts. Leave empty to create alert rules without notifications.')
param infoActionGroupResourceIds array = []

@description('Optional notification email used to auto-create default operational/info action groups. Leave empty to skip auto-creation.')
param defaultAlertNotificationEmail string = ''

@description('Short name for the auto-created operational action group (max 12 chars).')
@maxLength(12)
param defaultOperationalActionGroupShortName string = 'GSIOps'

@description('Short name for the auto-created info action group (max 12 chars).')
@maxLength(12)
param defaultInfoActionGroupShortName string = 'GSIInfo'

var isFlexConsumption = hostingPlan == 'FlexConsumption'
var deploymentContainerName = 'app-package'
var storageAccountName = toLower(replace(functionAppName, '-', ''))
var baseReleaseUrl = 'https://github.com/jpawlowski/spfx-guest-sponsor-info/releases'
// Strip a leading 'v' from appVersion so both '1.4.2' and 'v1.4.2' work correctly.
var normalizedAppVersion = startsWith(appVersion, 'v') ? substring(appVersion, 1) : appVersion
var resolvedPackageUrl = !empty(packageUrl)
  ? packageUrl
  : appVersion == 'latest'
      ? '${baseReleaseUrl}/latest/download/guest-sponsor-info-function.zip'
      : '${baseReleaseUrl}/download/v${normalizedAppVersion}/guest-sponsor-info-function.zip'
var builtInTags = {
  'managed-by': 'bicep'
  source: 'https://github.com/jpawlowski/spfx-guest-sponsor-info'
  'package-version': appVersion
}
var effectiveTags = union(builtInTags, tags)
var appServicePlanName = '${functionAppName}-plan'
var logAnalyticsWorkspaceName = '${functionAppName}-logs'
var appInsightsName = '${functionAppName}-insights'
var azureMapsName = empty(azureMapsAccountName)
  ? toLower('maps${uniqueString(resourceGroup().id, functionAppName)}')
  : toLower(azureMapsAccountName)
// KQL queries use triple-quoted raw strings (no interpolation) with replace() for parameters.
var serviceOutageAlertQueryRaw = '''
let window = __WINDOW__m;
let req = requests
| where timestamp > ago(window)
| where name contains "getGuestSponsors";
let total = toscalar(req | count);
let failures5xx = toscalar(req | where resultCode startswith "5" or resultCode == "504" | count);
let success = toscalar(req | where resultCode startswith "2" | count);
print total=total, failures5xx=failures5xx, success=success,
      successRatePct = iff(total == 0, 100.0, todouble(success) * 100.0 / todouble(total))
| where total >= __MIN_REQUESTS__
| where failures5xx >= __FAILURE_COUNT__ or successRatePct < __SUCCESS_RATE_PCT__
'''
#disable-next-line prefer-interpolation
var serviceOutageAlertQuery = replace(
  replace(
    replace(
      replace(serviceOutageAlertQueryRaw, '__WINDOW__', string(alertWindowInMinutes)),
      '__MIN_REQUESTS__',
      string(serviceOutageMinRequests)
    ),
    '__FAILURE_COUNT__',
    string(serviceOutageFailureCountThreshold)
  ),
  '__SUCCESS_RATE_PCT__',
  string(serviceOutageSuccessRatePercentThreshold)
)

var authConfigRegressionAlertQueryRaw = '''
let window = __WINDOW__m;
traces
| where timestamp > ago(window)
| where message has "Client validation ("
| extend reasonCode = tostring(customDimensions.reasonCode)
| where reasonCode in ("AUTH_CONFIG_TENANT_MISSING", "AUTH_CONFIG_AUDIENCE_MISSING")
| summarize hits = count() by reasonCode
| where hits >= __HITS_THRESHOLD__
'''
#disable-next-line prefer-interpolation
var authConfigRegressionAlertQuery = replace(
  replace(authConfigRegressionAlertQueryRaw, '__WINDOW__', string(alertWindowInMinutes)),
  '__HITS_THRESHOLD__',
  string(authConfigRegressionHitsThreshold)
)

var likelyAttackInfoAlertQueryRaw = '''
let window = __WINDOW__m;
let req = requests
| where timestamp > ago(window)
| where name contains "getGuestSponsors";
let denied = req
| where resultCode in ("401", "403")
| summarize deniedCount = count(), uniqueIps = dcount(client_IP);
let total = toscalar(req | count);
let success = toscalar(req | where resultCode startswith "2" | count);
denied
| extend denyRatePct = iff(total == 0, 0.0, todouble(deniedCount) * 100.0 / todouble(total))
| where deniedCount >= __DENIED_COUNT__
| where uniqueIps >= __UNIQUE_IP__
| where denyRatePct >= __DENY_RATE_PCT__
| where success >= __MIN_SUCCESS__
'''
#disable-next-line prefer-interpolation
var likelyAttackInfoAlertQuery = replace(
  replace(
    replace(
      replace(
        replace(likelyAttackInfoAlertQueryRaw, '__WINDOW__', string(alertWindowInMinutes)),
        '__DENIED_COUNT__',
        string(likelyAttackDeniedCountThreshold)
      ),
      '__UNIQUE_IP__',
      string(likelyAttackUniqueIpThreshold)
    ),
    '__DENY_RATE_PCT__',
    string(likelyAttackDenyRatePercentThreshold)
  ),
  '__MIN_SUCCESS__',
  string(likelyAttackMinSuccessThreshold)
)
var createDefaultActionGroups = !empty(defaultAlertNotificationEmail)

resource defaultOperationalActionGroup 'Microsoft.Insights/actionGroups@2023-01-01' = if (createDefaultActionGroups) {
  name: '${functionAppName}-ops-ag'
  location: 'global'
  tags: effectiveTags
  properties: {
    groupShortName: defaultOperationalActionGroupShortName
    enabled: true
    emailReceivers: [
      {
        name: 'ops-email'
        emailAddress: defaultAlertNotificationEmail
        useCommonAlertSchema: true
      }
    ]
  }
}

resource defaultInfoActionGroup 'Microsoft.Insights/actionGroups@2023-01-01' = if (createDefaultActionGroups) {
  name: '${functionAppName}-info-ag'
  location: 'global'
  tags: effectiveTags
  properties: {
    groupShortName: defaultInfoActionGroupShortName
    enabled: true
    emailReceivers: [
      {
        name: 'info-email'
        emailAddress: defaultAlertNotificationEmail
        useCommonAlertSchema: true
      }
    ]
  }
}

var effectiveOperationalActionGroupResourceIds = createDefaultActionGroups
  ? concat(operationalActionGroupResourceIds, [defaultOperationalActionGroup.id])
  : operationalActionGroupResourceIds

var effectiveInfoActionGroupResourceIds = createDefaultActionGroups
  ? concat(infoActionGroupResourceIds, [defaultInfoActionGroup.id])
  : infoActionGroupResourceIds

// ── Storage Account (required by Azure Functions runtime) ────────────────────
resource storageAccount 'Microsoft.Storage/storageAccounts@2023-01-01' = {
  name: length(storageAccountName) > 24 ? substring(storageAccountName, 0, 24) : storageAccountName
  location: location
  tags: effectiveTags
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
  tags: effectiveTags
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
  tags: effectiveTags
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

// ── App Service Plan ─────────────────────────────────────────────────────────
// Only one of these two resources is deployed, depending on the hostingPlan parameter.

// Consumption (Y1 / Dynamic) — includes Azure free grant (1M executions + 400 000 GB-s/month).
resource appServicePlanConsumption 'Microsoft.Web/serverfarms@2023-01-01' = if (!isFlexConsumption) {
  name: appServicePlanName
  location: location
  tags: effectiveTags
  sku: {
    name: 'Y1'
    tier: 'Dynamic'
  }
  properties: {}
}

// Flex Consumption (FC1) — Linux-only, no free grant.
// Requires API version 2023-12-01 or later. Not available in all Azure regions.
resource appServicePlanFlex 'Microsoft.Web/serverfarms@2023-12-01' = if (isFlexConsumption) {
  name: appServicePlanName
  location: location
  tags: effectiveTags
  kind: 'functionapp'
  sku: {
    name: 'FC1'
    tier: 'FlexConsumption'
  }
  properties: {
    reserved: true // Linux
  }
}

// ── Blob container for Flex Consumption deployment package ───────────────────
// Flex Consumption deploys code from a blob container rather than a direct URL.
// The Managed Identity needs read access — granted via role assignment below.
resource blobService 'Microsoft.Storage/storageAccounts/blobServices@2023-01-01' = if (isFlexConsumption) {
  parent: storageAccount
  name: 'default'
}

resource deploymentContainer 'Microsoft.Storage/storageAccounts/blobServices/containers@2023-01-01' = if (isFlexConsumption) {
  parent: blobService
  name: deploymentContainerName
  properties: {
    publicAccess: 'None'
  }
}

// ── Deployment script — Flex Consumption initial ZIP upload ──────────────────
// Flex Consumption cannot use WEBSITE_RUN_FROM_PACKAGE with a remote URL.
// This Azure CLI container script runs during ARM provisioning (~2 min): it
// downloads the function ZIP from packageUrl and uploads it to the app-package
// blob container so the Function App mounts and serves it automatically.
//
// forceUpdateTag = hash of packageUrl: the script re-runs only when packageUrl
// changes (e.g. pinning a new release version), not on every redeployment.

// User-Assigned Managed Identity for the deployment script.
// Deployment scripts require a User-Assigned Identity — System-Assigned is not supported.
resource deployScriptIdentity 'Microsoft.ManagedIdentity/userAssignedIdentities@2023-01-31' = if (isFlexConsumption) {
  name: '${functionAppName}-deploy-id'
  location: location
  tags: effectiveTags
}

// Grant the deployment script identity write access to the deployment container.
// Storage Blob Data Contributor is sufficient — full owner access is not required.
resource deployScriptBlobRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (isFlexConsumption) {
  scope: storageAccount
  name: guid(storageAccount.id, '${functionAppName}-deploy-id', 'ba92f5b4-2d11-453d-a403-e96b0029c9fe')
  properties: {
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      'ba92f5b4-2d11-453d-a403-e96b0029c9fe' // Storage Blob Data Contributor
    )
    #disable-next-line BCP318 // Safe: deployScriptIdentity is always deployed when isFlexConsumption.
    principalId: deployScriptIdentity.properties.principalId
    principalType: 'ServicePrincipal'
  }
}

// Extracted to a variable so BCP318 is suppressed at declaration and not inside the resource body.
// Safe: deployScriptIdentity shares the same isFlexConsumption condition as deployZipScript.
#disable-next-line BCP318
var deployScriptIdentityResourceId = deployScriptIdentity.id

resource deployZipScript 'Microsoft.Resources/deploymentScripts@2023-08-01' = if (isFlexConsumption) {
  name: '${functionAppName}-deploy-zip'
  location: location
  kind: 'AzureCLI'
  tags: effectiveTags
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: {
      '${deployScriptIdentityResourceId}': {}
    }
  }
  dependsOn: [
    deployScriptBlobRole
    deploymentContainer
  ]
  properties: {
    azCliVersion: '2.60.0'
    retentionInterval: 'PT2H'
    // Re-run only when appVersion/packageUrl changes — not on every template redeployment.
    forceUpdateTag: uniqueString(resolvedPackageUrl)
    environmentVariables: [
      { name: 'STORAGE_ACCOUNT', value: storageAccount.name }
      { name: 'CONTAINER', value: deploymentContainerName }
      { name: 'PACKAGE_URL', value: resolvedPackageUrl }
    ]
    // Retry loop to handle RBAC propagation delay (typically 30–60 s after role assignment).
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

// ── Azure Maps account (optional; used by inline map preview in SPFx card) ───
resource azureMapsAccount 'Microsoft.Maps/accounts@2023-06-01' = if (deployAzureMaps) {
  name: azureMapsName
  location: azureMapsLocation
  tags: effectiveTags
  sku: {
    name: 'G2'
  }
  kind: 'Gen2'
  properties: {
    disableLocalAuth: false
  }
}

// ── Function App ─────────────────────────────────────────────────────────────
// Only ONE of the two resources below is deployed, depending on hostingPlan.

// App settings shared by both Consumption and Flex Consumption plans.
var sharedAppSettings = [
  // Identity-based storage connection — no account key stored anywhere.
  // The role assignments below grant the Managed Identity the minimum
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
    // Automatic instrumentation — no code changes required.
    name: 'APPLICATIONINSIGHTS_CONNECTION_STRING'
    value: appInsights.properties.ConnectionString
  }
  {
    // Visible in Application Insights telemetry and Azure Portal tags.
    name: 'APP_VERSION'
    value: appVersion
  }
]

// Additional app settings only needed on the Consumption (Y1) plan.
// Flex Consumption configures runtime, version, and deployment via functionAppConfig.
var consumptionOnlyAppSettings = [
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
    value: resolvedPackageUrl
  }
]

// Consumption plan — Y1 / Dynamic (free tier, cold starts, GitHub package URL).
resource functionApp 'Microsoft.Web/sites@2023-01-01' = if (!isFlexConsumption) {
  name: functionAppName
  location: location
  tags: effectiveTags
  kind: 'functionapp'
  identity: {
    type: 'SystemAssigned'
  }
  properties: {
    serverFarmId: appServicePlanConsumption.id
    httpsOnly: true
    // Suspend the Function App when the daily GB-second budget is exhausted.
    // This is the primary cost-containment guard for the Consumption plan.
    // Set dailyMemoryTimeQuotaGBs=0 to disable (not recommended).
    dailyMemoryTimeQuota: dailyMemoryTimeQuotaGBs
    siteConfig: {
      appSettings: concat(sharedAppSettings, consumptionOnlyAppSettings)
      cors: {
        allowedOrigins: [
          'https://${tenantName}.sharepoint.com'
        ]
        supportCredentials: false
      }
    }
  }
}

// Flex Consumption plan — FC1 / Linux (no free tier, fast cold starts).
resource functionAppFlex 'Microsoft.Web/sites@2023-12-01' = if (isFlexConsumption) {
  name: functionAppName
  location: location
  tags: effectiveTags
  kind: 'functionapp,linux'
  identity: {
    type: 'SystemAssigned'
  }
  dependsOn: [
    deploymentContainer // Ensure the blob container exists before the deployment config references it.
  ]
  properties: {
    serverFarmId: appServicePlanFlex.id
    httpsOnly: true
    functionAppConfig: {
      runtime: {
        name: 'node'
        version: '22'
      }
      scaleAndConcurrency: {
        maximumInstanceCount: maximumFlexInstances
        instanceMemoryMB: instanceMemoryMB
        alwaysReady: alwaysReadyInstances > 0
          ? [
              {
                name: 'http'
                instanceCount: alwaysReadyInstances
              }
            ]
          : []
      }
      deployment: {
        storage: {
          type: 'blobContainer'
          value: '${storageAccount.properties.primaryEndpoints.blob}${deploymentContainerName}'
          authentication: {
            type: 'SystemAssignedIdentity'
          }
        }
      }
    }
    siteConfig: {
      appSettings: sharedAppSettings
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
// Shared auth configuration — identical for both hosting plans.
var easyAuthProperties = {
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

resource authSettings 'Microsoft.Web/sites/config@2023-01-01' = if (!isFlexConsumption) {
  name: 'authsettingsV2'
  parent: functionApp
  properties: easyAuthProperties
}

resource authSettingsFlex 'Microsoft.Web/sites/config@2023-12-01' = if (isFlexConsumption) {
  name: 'authsettingsV2'
  parent: functionAppFlex
  properties: easyAuthProperties
}

// ── Storage role assignments (Managed Identity auth, no key required) ────────
// Both plans need blob, queue, and table access for the Azure Functions runtime.
// 'Owner' (or a custom role with Microsoft.Authorization/roleAssignments/write)
// is required on the deploying principal to create these assignments.
//
// Each role is duplicated with mutually exclusive conditions because the
// principalId comes from the conditionally deployed Function App resource.
var storageBlobDataOwnerRoleId = 'b7e6dc6d-f1e8-4753-8033-0f276bb0955b'
var storageQueueDataContributorRoleId = '974c5e8b-45b9-4653-ba55-5f855dd0fb88'
var storageTableDataContributorRoleId = '0a9a7e1f-b9d0-4cc4-a60d-0319b160aaa3'
var functionAppResourceId = resourceId('Microsoft.Web/sites', functionAppName)

// ── Consumption plan role assignments ────────────────────────────────────────
resource storageBlobRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (!isFlexConsumption) {
  scope: storageAccount
  name: guid(storageAccount.id, functionAppResourceId, storageBlobDataOwnerRoleId)
  properties: {
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', storageBlobDataOwnerRoleId)
    #disable-next-line BCP318 // Safe: functionApp is always deployed when !isFlexConsumption.
    principalId: functionApp.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

resource storageQueueRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (!isFlexConsumption) {
  scope: storageAccount
  name: guid(storageAccount.id, functionAppResourceId, storageQueueDataContributorRoleId)
  properties: {
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      storageQueueDataContributorRoleId
    )
    #disable-next-line BCP318
    principalId: functionApp.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

resource storageTableRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (!isFlexConsumption) {
  scope: storageAccount
  name: guid(storageAccount.id, functionAppResourceId, storageTableDataContributorRoleId)
  properties: {
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      storageTableDataContributorRoleId
    )
    #disable-next-line BCP318
    principalId: functionApp.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

// ── Flex Consumption plan role assignments ────────────────────────────────────
resource storageBlobRoleFlex 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (isFlexConsumption) {
  scope: storageAccount
  name: guid(storageAccount.id, functionAppResourceId, storageBlobDataOwnerRoleId)
  properties: {
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', storageBlobDataOwnerRoleId)
    #disable-next-line BCP318 // Safe: functionAppFlex is always deployed when isFlexConsumption.
    principalId: functionAppFlex.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

resource storageQueueRoleFlex 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (isFlexConsumption) {
  scope: storageAccount
  name: guid(storageAccount.id, functionAppResourceId, storageQueueDataContributorRoleId)
  properties: {
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      storageQueueDataContributorRoleId
    )
    #disable-next-line BCP318
    principalId: functionAppFlex.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

resource storageTableRoleFlex 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (isFlexConsumption) {
  scope: storageAccount
  name: guid(storageAccount.id, functionAppResourceId, storageTableDataContributorRoleId)
  properties: {
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      storageTableDataContributorRoleId
    )
    #disable-next-line BCP318
    principalId: functionAppFlex.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

// ── Optional KQL alerts (low false-positive model) ───────────────────────────
resource serviceOutageAlert 'Microsoft.Insights/scheduledQueryRules@2021-08-01' = if (enableServiceOutageAlert) {
  name: '${functionAppName}-service-outage-kql'
  location: location
  tags: effectiveTags
  properties: {
    description: 'Operational email alert for probable service outage (5xx/504 spike or low success rate).'
    enabled: true
    scopes: [
      appInsights.id
    ]
    evaluationFrequency: 'PT${alertEvaluationFrequencyInMinutes}M'
    windowSize: 'PT${alertWindowInMinutes}M'
    severity: 2
    criteria: {
      allOf: [
        {
          query: serviceOutageAlertQuery
          timeAggregation: 'Count'
          operator: 'GreaterThan'
          threshold: 0
        }
      ]
    }
    actions: {
      actionGroups: effectiveOperationalActionGroupResourceIds
    }
    autoMitigate: true
  }
}

resource authConfigRegressionAlert 'Microsoft.Insights/scheduledQueryRules@2021-08-01' = if (enableAuthConfigRegressionAlert) {
  name: '${functionAppName}-auth-config-regression-kql'
  location: location
  tags: effectiveTags
  properties: {
    description: 'Operational email alert for auth/config regressions (AUTH_CONFIG_* reason codes).'
    enabled: true
    scopes: [
      appInsights.id
    ]
    evaluationFrequency: 'PT${alertEvaluationFrequencyInMinutes}M'
    windowSize: 'PT${alertWindowInMinutes}M'
    severity: 2
    criteria: {
      allOf: [
        {
          query: authConfigRegressionAlertQuery
          timeAggregation: 'Count'
          operator: 'GreaterThan'
          threshold: 0
        }
      ]
    }
    actions: {
      actionGroups: effectiveOperationalActionGroupResourceIds
    }
    autoMitigate: true
  }
}

resource likelyAttackInfoAlert 'Microsoft.Insights/scheduledQueryRules@2021-08-01' = if (enableLikelyAttackInfoAlert) {
  name: '${functionAppName}-likely-attack-info-kql'
  location: location
  tags: effectiveTags
  properties: {
    description: 'Info-only alert for likely attack/noise spikes (high 401/403 from many IPs).'
    enabled: true
    scopes: [
      appInsights.id
    ]
    evaluationFrequency: 'PT${alertEvaluationFrequencyInMinutes}M'
    windowSize: 'PT${alertWindowInMinutes}M'
    severity: 4
    criteria: {
      allOf: [
        {
          query: likelyAttackInfoAlertQuery
          timeAggregation: 'Count'
          operator: 'GreaterThan'
          threshold: 0
        }
      ]
    }
    actions: {
      actionGroups: effectiveInfoActionGroupResourceIds
    }
    autoMitigate: true
  }
}

// ── Outputs ──────────────────────────────────────────────────────────────────
// The hostname follows a fixed pattern — no runtime reference needed.
var functionAppHostName = '${functionAppName}.azurewebsites.net'

@description('The base URL of the deployed Function App. Paste this into the SPFx web part property pane (Azure Function Base URL field).')
output functionAppUrl string = 'https://${functionAppHostName}'

@description('The full function endpoint URL — use this for curl/Postman testing or health checks. The web part property pane only needs the base URL (functionAppUrl).')
output sponsorApiUrl string = 'https://${functionAppHostName}/api/getGuestSponsors'

// Safe: exactly one of the two conditional Function App resources is always deployed.
#disable-next-line BCP318
var flexPrincipalId = functionAppFlex.identity.principalId
#disable-next-line BCP318
var consumptionPrincipalId = functionApp.identity.principalId

@description('Object ID of the system-assigned Managed Identity — needed for setup-graph-permissions.ps1.')
output managedIdentityObjectId string = isFlexConsumption ? flexPrincipalId : consumptionPrincipalId

@description('Name of the Application Insights component — open in the Azure Portal for live telemetry.')
output appInsightsName string = appInsights.name

@description('The selected hosting plan — included in outputs for operational visibility.')
output hostingPlan string = hostingPlan

@description('Azure Maps account name (empty when deployAzureMaps=false).')
output azureMapsAccountName string = deployAzureMaps ? azureMapsAccount.name : ''

@description('Azure CLI command to fetch the Azure Maps primary key (empty when deployAzureMaps=false).')
output azureMapsKeyCommand string = deployAzureMaps
  ? 'az maps account keys list -g ${resourceGroup().name} -n ${azureMapsAccount.name} --query primaryKey -o tsv'
  : ''

@description('Name of the Storage Account. For Flex Consumption: upload updated ZIPs to the app-package container here to trigger a redeployment (use --auth-mode login). For Consumption: the runtime uses this account for trigger state and blob/queue/table operations.')
output deploymentStorageAccountName string = storageAccount.name

@description('The function package version deployed. "latest" = newest release at provisioning time; otherwise the pinned SemVer tag without "v" prefix.')
output deployedAppVersion string = appVersion
