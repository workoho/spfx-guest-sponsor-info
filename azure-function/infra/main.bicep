extension microsoftGraphV1

targetScope = 'resourceGroup'

metadata name = 'Guest Sponsor API for Microsoft Entra B2B'
metadata description = 'Deploys an Azure Function App that acts as a Graph API proxy for the Guest Sponsor Info for Microsoft Entra B2B SharePoint web part. Includes a Storage Account, App Service Plan, EasyAuth configuration, Log Analytics Workspace, and Application Insights. Always creates the Entra App Registration. Graph application role assignments to the Managed Identity are included by default and can be deferred to setup-graph-permissions.ps1 via skipGraphRoleAssignments=true.'
metadata repository = 'https://github.com/workoho/spfx-guest-sponsor-info'
metadata author = 'Workoho GmbH'
metadata license = 'PolyForm-Shield-1.0.0'

@metadata({ category: 'Basics' })
@description('Azure region for all resources.')
param location string = resourceGroup().location

@metadata({ category: 'Basics' })
@description('Entra tenant ID (GUID).')
param tenantId string = tenant().tenantId

@metadata({ category: 'Basics' })
@description('Tenant name without domain suffix, e.g. "contoso".')
param tenantName string

@metadata({ category: 'Basics' })
@description('Name for the Function App (2–58 characters, letters, numbers, and hyphens only). Leave empty to auto-generate a deterministic name scoped to this resource group — re-deployments always produce the same name, and the URL is not easily guessable.')
@minLength(2)
@maxLength(58)
param functionAppName string = 'gsi-${uniqueString(resourceGroup().id)}'

@metadata({ category: 'Hosting' })
@description('Hosting plan for the Function App. "Consumption" = Y1/Dynamic (free tier included, cold starts after ~20 min idle, ZIP served directly from GitHub package URL). "FlexConsumption" = FC1/Linux-only (no free tier, cold starts greatly reduced — alwaysReadyInstances=1 eliminates them; ZIP is uploaded to blob storage by the provisioning script automatically during Bicep/azd deployment). Not all Azure regions support Flex Consumption — check https://aka.ms/flex-region before choosing.')
@allowed([
  'Consumption'
  'FlexConsumption'
])
param hostingPlan string = 'Consumption'

@metadata({ category: 'Hosting' })
@description('Number of always-ready (pre-warmed) instances for the Function App (Flex Consumption plan only). 0 = purely on-demand (cold starts possible). 1 = one instance kept warm — eliminates cold starts (~€2-5/month). Ignored when hostingPlan = "Consumption".')
param alwaysReadyInstances string = '1'

@metadata({ category: 'Hosting' })
@description('Hard upper bound on the number of instances the Flex Consumption plan may scale out to (Flex Consumption plan only). Acts as a cost ceiling — scale-out stops at this limit regardless of demand. Valid range: 1-1000. Default: 10. Ignored when hostingPlan = "Consumption".')
param maximumFlexInstances string = '10'

@metadata({ category: 'Hosting' })
@description('Memory allocated to each Flex Consumption instance in MB (Flex Consumption plan only). Valid values: 512 or 2048. Higher memory allows more concurrent requests per instance but costs more per GB-second. Default: 2048. Ignored when hostingPlan = "Consumption".')
@allowed(['512', '2048'])
param instanceMemoryMB string = '2048'

@metadata({ category: 'Deployment' })
@description('Function package version to deploy. "latest" (default) = always pull the newest GitHub Release at provisioning time. SemVer without "v" prefix, e.g. "1.4.2" = pin to that specific release. On Consumption: sets the WEBSITE_RUN_FROM_PACKAGE URL. On Flex Consumption: the provisioning script re-runs and re-uploads the ZIP whenever this value changes — set it on each redeployment to trigger a code update.')
param appVersion string = 'latest'

@metadata({ category: 'Deployment' })
@description('Override the computed package ZIP URL. Leave empty (default) to auto-compute from appVersion. Use only when hosting the ZIP at a custom location.')
param packageUrl string = ''

@metadata({ category: 'Tags' })
@description('Optional deployment environment tag applied to the resource group and all deployed resources. Leave empty to omit. Examples: prod, production, dev, nonprod.')
param environment string = ''

@metadata({ category: 'Tags' })
@description('Optional workload criticality tag applied to the resource group and all deployed resources. Leave empty to omit. Examples: low, medium, high, mission-critical, business-critical.')
param criticality string = ''

@metadata({ category: 'Tags' })
@description('Additional resource tags to apply to all deployed resources. The built-in tags "application", optional "environment" and "criticality", "managed-by", "repository-url", "documentation-url", "support-url", "publisher", "license", "package-version", and the legacy compatibility tag "source" are always set automatically and cannot be overridden when present.')
param tags object = {}

@metadata({ category: 'Azure Maps' })
@description('Deploy Azure Maps account for inline address map preview.')
@allowed(['true', 'false'])
param deployAzureMaps string = 'true'

@metadata({ category: 'Azure Maps' })
@description('Optional custom Azure Maps account name. Leave empty to auto-generate.')
param azureMapsAccountName string = ''

@metadata({ category: 'Azure Maps' })
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

@metadata({ category: 'Monitoring' })
@description('Enable operational email alert for probable service outage (5xx/504 spike or low success rate).')
param enableServiceOutageAlert bool = true

@metadata({ category: 'Monitoring' })
@description('Deploy the monitoring stack (Log Analytics, Application Insights, action groups, and KQL alerts). Default: true.')
@allowed(['true', 'false'])
param enableMonitoring string = 'true'

@metadata({ category: 'Monitoring' })
@description('Enable operational email alert for auth/config regressions (AUTH_CONFIG_* reason codes).')
param enableAuthConfigRegressionAlert bool = true

@metadata({ category: 'Monitoring' })
@description('Enable info-only alert for likely attack/noise spikes (high 401/403 from many IPs).')
param enableLikelyAttackInfoAlert bool = true

@metadata({ category: 'Monitoring' })
@description('KQL alert evaluation frequency in minutes.')
@minValue(1)
param alertEvaluationFrequencyInMinutes int = 5

@metadata({ category: 'Monitoring' })
@description('KQL alert lookback window in minutes.')
@minValue(5)
param alertWindowInMinutes int = 15

@metadata({ category: 'Monitoring' })
@description('Minimum total requests in window before service outage alert can fire.')
@minValue(1)
param serviceOutageMinRequests int = 20

@metadata({ category: 'Monitoring' })
@description('5xx/504 count threshold for service outage alert.')
@minValue(1)
param serviceOutageFailureCountThreshold int = 10

@metadata({ category: 'Monitoring' })
@description('Success-rate percentage threshold below which service outage alert can fire.')
@minValue(1)
@maxValue(99)
param serviceOutageSuccessRatePercentThreshold int = 70

@metadata({ category: 'Monitoring' })
@description('AUTH_CONFIG_* trace count threshold for config-regression alert.')
@minValue(1)
param authConfigRegressionHitsThreshold int = 1

@metadata({ category: 'Monitoring' })
@description('401/403 count threshold for likely-attack info alert.')
@minValue(1)
param likelyAttackDeniedCountThreshold int = 50

@metadata({ category: 'Monitoring' })
@description('Unique client IP threshold for likely-attack info alert.')
@minValue(1)
param likelyAttackUniqueIpThreshold int = 20

@metadata({ category: 'Monitoring' })
@description('Denied-rate percentage threshold for likely-attack info alert.')
@minValue(1)
@maxValue(100)
param likelyAttackDenyRatePercentThreshold int = 80

@metadata({ category: 'Monitoring' })
@description('Minimum successful requests required before likely-attack info alert fires (avoid pure outage overlap).')
@minValue(0)
param likelyAttackMinSuccessThreshold int = 1

@metadata({ category: 'Monitoring' })
@description('Enable info-only alert when a newer GitHub release of the function is available.')
param enableNewReleaseAlert bool = true

@metadata({ category: 'Monitoring' })
@description('Enable operational alert when a hard-deleted Entra object remains referenced as a sponsor (Graph 404).')
param enableBrokenSponsorAlert bool = false

@metadata({ category: 'Monitoring' })
@description('Enable the Application Insights Failure Anomalies smart detector alert rule. Default: false, so the rule stays disabled unless explicitly activated.')
@allowed(['true', 'false'])
param enableFailureAnomaliesAlert string = 'false'

@metadata({ category: 'Monitoring' })
@description('KQL evaluation frequency for the new-release alert in minutes.')
@minValue(5)
param newReleaseAlertEvaluationFrequencyInMinutes int = 60

@metadata({ category: 'Monitoring' })
@description('KQL lookback window for the new-release alert in minutes (default 720 = 12 h covers two 6-hour timer intervals).')
@minValue(60)
param newReleaseAlertWindowInMinutes int = 720

@metadata({ category: 'Hosting' })
@description('Consumption plan (Y1) daily memory-time budget in GB-seconds. When hit, the Function App is suspended until midnight UTC — the primary cost guard. 0 = unlimited (not recommended). Default 10 000 GB-s ≈ 13 000 typical invocations/day, stays within the monthly free tier (400 000 GB-s/month). Ignored when hostingPlan = "FlexConsumption" (Flex has no daily GB-second budget concept).')
@minValue(0)
param dailyMemoryTimeQuotaGBs int = 10000

@metadata({ category: 'Monitoring' })
@description('Action group resource IDs for operational email alerts. Leave empty to create alert rules without notifications.')
param operationalActionGroupResourceIds array = []

@metadata({ category: 'Monitoring' })
@description('Action group resource IDs for info-only alerts. Leave empty to create alert rules without notifications.')
param infoActionGroupResourceIds array = []

@metadata({ category: 'Monitoring' })
@description('Optional notification email used to auto-create default operational/info action groups. Leave empty to skip auto-creation.')
param defaultAlertNotificationEmail string = ''

@metadata({ category: 'Monitoring' })
@description('Short name for the auto-created operational action group (max 12 chars).')
@maxLength(12)
param defaultOperationalActionGroupShortName string = 'GSIOps'

@metadata({ category: 'Monitoring' })
@description('Short name for the auto-created info action group (max 12 chars).')
@maxLength(12)
param defaultInfoActionGroupShortName string = 'GSIInfo'

@metadata({ category: 'Telemetry' })
@description('Enable Customer Usage Attribution (CUA): an empty nested deployment named pid-18fb4033-c9f3-41fa-a5db-e3a03b012939 is created in your resource group. Microsoft forwards aggregated Azure consumption figures for that GUID to Workoho via Partner Center — no personal data or resource details ever leave your subscription. Set to false to opt out. See https://aka.ms/partnercenter-attribution')
param enableTelemetry bool = true

@metadata({ category: 'Entra' })
@description('Skip Microsoft Graph application role assignments to the Managed Identity. Set to "true" when setup-graph-permissions.ps1 will be run separately by an admin with Privileged Role Administrator rights. The App Registration and its configuration are always managed by Bicep regardless of this setting. Default: "false" — Bicep assigns all required Graph permissions in the same deployment.')
param skipGraphRoleAssignments string = 'false'

var skipRoleAssignments = skipGraphRoleAssignments == 'true'
var deployAzureMapsEnabled = deployAzureMaps == 'true'
var enableMonitoringEnabled = enableMonitoring == 'true'
var enableFailureAnomaliesAlertEnabled = enableFailureAnomaliesAlert == 'true'
var alwaysReadyInstancesValue = int(alwaysReadyInstances)
var maximumFlexInstancesValue = int(maximumFlexInstances)
var instanceMemoryMBValue = instanceMemoryMB == '512' ? 512 : 2048
var repositoryUrl = 'https://github.com/workoho/spfx-guest-sponsor-info'
var docsUrl = 'https://guest-sponsor-info.workoho.cloud'
var licenseId = 'PolyForm-Shield-1.0.0'
var baseReleaseUrl = '${repositoryUrl}/releases'
// Strip a leading 'v' from appVersion so both '1.4.2' and 'v1.4.2' work correctly.
var normalizedAppVersion = startsWith(appVersion, 'v') ? substring(appVersion, 1) : appVersion
var resolvedPackageUrl = !empty(packageUrl)
  ? packageUrl
  : appVersion == 'latest'
      ? '${baseReleaseUrl}/latest/download/guest-sponsor-info-function.zip'
      : '${baseReleaseUrl}/download/v${normalizedAppVersion}/guest-sponsor-info-function.zip'
var baseBuiltInTags = {
  application: 'guest-sponsor-info'
  'managed-by': 'bicep'
  'repository-url': repositoryUrl
  'documentation-url': docsUrl
  'support-url': '${repositoryUrl}/issues'
  publisher: 'Workoho GmbH'
  license: licenseId
  source: repositoryUrl
  'package-version': appVersion
}
var environmentTag = empty(environment)
  ? {}
  : {
      environment: environment
    }
var criticalityTag = empty(criticality)
  ? {}
  : {
      criticality: criticality
    }
var optionalBuiltInTags = union(environmentTag, criticalityTag)
var builtInTags = union(baseBuiltInTags, optionalBuiltInTags)
var effectiveTags = union(builtInTags, tags)
var resourceGroupEffectiveTags = union(resourceGroup().tags, effectiveTags)
var azureMapsName = empty(azureMapsAccountName)
  ? toLower('maps${uniqueString(resourceGroup().id, functionAppName)}')
  : toLower(azureMapsAccountName)
var appRegDescription = 'EasyAuth identity provider for the "Guest Sponsor Info" SharePoint Online web part (SPFx). Authenticates requests from the web part to the Azure Function proxy, which calls Microsoft Graph on behalf of signed-in guest users to retrieve their Entra sponsor information. Tokens are acquired silently via pre-authorized SharePoint Online Web Client Extensibility. Docs: ${docsUrl}'
var appRegNotes = 'Do not delete — the "Guest Sponsor Info" SharePoint web part depends on this for guest sponsor lookups via Microsoft Graph. The associated Azure Function uses a system-assigned Managed Identity for Graph API calls (User.Read.All, Presence.Read.All, MailboxSettings.Read, TeamMember.Read.All). Docs: ${docsUrl}'
var appRegInfo = {
  termsOfServiceUrl: '${repositoryUrl}/blob/main/docs/terms-of-use.md'
  privacyStatementUrl: '${repositoryUrl}/blob/main/docs/privacy-policy.md'
  supportUrl: '${repositoryUrl}/issues'
  marketingUrl: docsUrl
}
var appRegLogo = loadFileAsBase64('./assets/icon-300.png')
var userImpersonationScopeId = guid(functionAppName, 'user-impersonation')

// ── Customer Usage Attribution (Partner Center tracking) ─────────────────────
// Empty nested deployment whose name carries the Partner Center GUID. Azure
// records this GUID against every resource group deployment that includes this
// template, allowing Workoho to see adoption metrics in Partner Center without
// collecting any customer data. See https://learn.microsoft.com/partner-center/marketplace-offers/azure-partner-customer-usage-attribution
// The no-deployments-resources rule is suppressed: Microsoft's CUA pattern
// intentionally requires a named nested deployment — a Bicep module cannot
// carry the pid- prefix required for attribution.
#disable-next-line no-deployments-resources
resource partnerAttribution 'Microsoft.Resources/deployments@2021-04-01' = if (enableTelemetry) {
  name: 'pid-18fb4033-c9f3-41fa-a5db-e3a03b012939'
  properties: {
    mode: 'Incremental'
    template: {
      '$schema': 'https://schema.management.azure.com/schemas/2019-04-01/deploymentTemplate.json#'
      contentVersion: '1.0.0.0'
      resources: []
    }
  }
}

// Keep azd-owned resource group tags such as azd-env-name and add the same
// workload metadata that is applied to the deployed resources.
resource resourceGroupTags 'Microsoft.Resources/tags@2021-04-01' = {
  name: 'default'
  properties: {
    tags: resourceGroupEffectiveTags
  }
}

// ── Monitoring module ────────────────────────────────────────────────────────
// Log Analytics, Application Insights, Action Groups, and KQL alert rules are
// managed in a dedicated module to keep this orchestration template focused.
module monitoring './modules/monitoring.bicep' = if (enableMonitoringEnabled) {
  name: 'monitoring'
  params: {
    location: location
    functionAppName: functionAppName
    tags: effectiveTags
    enableServiceOutageAlert: enableServiceOutageAlert
    enableAuthConfigRegressionAlert: enableAuthConfigRegressionAlert
    enableLikelyAttackInfoAlert: enableLikelyAttackInfoAlert
    alertEvaluationFrequencyInMinutes: alertEvaluationFrequencyInMinutes
    alertWindowInMinutes: alertWindowInMinutes
    serviceOutageMinRequests: serviceOutageMinRequests
    serviceOutageFailureCountThreshold: serviceOutageFailureCountThreshold
    serviceOutageSuccessRatePercentThreshold: serviceOutageSuccessRatePercentThreshold
    authConfigRegressionHitsThreshold: authConfigRegressionHitsThreshold
    likelyAttackDeniedCountThreshold: likelyAttackDeniedCountThreshold
    likelyAttackUniqueIpThreshold: likelyAttackUniqueIpThreshold
    likelyAttackDenyRatePercentThreshold: likelyAttackDenyRatePercentThreshold
    likelyAttackMinSuccessThreshold: likelyAttackMinSuccessThreshold
    enableNewReleaseAlert: enableNewReleaseAlert
    newReleaseAlertEvaluationFrequencyInMinutes: newReleaseAlertEvaluationFrequencyInMinutes
    newReleaseAlertWindowInMinutes: newReleaseAlertWindowInMinutes
    enableBrokenSponsorAlert: enableBrokenSponsorAlert
    enableFailureAnomaliesAlert: enableFailureAnomaliesAlertEnabled
    operationalActionGroupResourceIds: operationalActionGroupResourceIds
    infoActionGroupResourceIds: infoActionGroupResourceIds
    defaultAlertNotificationEmail: defaultAlertNotificationEmail
    defaultOperationalActionGroupShortName: defaultOperationalActionGroupShortName
    defaultInfoActionGroupShortName: defaultInfoActionGroupShortName
  }
}

// ── Function App module ───────────────────────────────────────────────────────
// Storage Account, App Service Plan, Function App (both variants), EasyAuth,
// Flex Consumption deployment script, and storage role assignments are all
// managed in a dedicated module.
module functionApp './modules/functionapp.bicep' = {
  params: {
    location: location
    functionAppName: functionAppName
    hostingPlan: hostingPlan
    alwaysReadyInstances: alwaysReadyInstancesValue
    maximumFlexInstances: maximumFlexInstancesValue
    instanceMemoryMB: instanceMemoryMBValue
    dailyMemoryTimeQuotaGBs: dailyMemoryTimeQuotaGBs
    resolvedPackageUrl: resolvedPackageUrl
    appVersion: appVersion
    tenantId: tenantId
    tenantName: tenantName
    appClientId: effectiveAppClientId
    appInsightsConnectionString: enableMonitoringEnabled
      #disable-next-line BCP318 // Safe: monitoring module is always deployed when enableMonitoringEnabled=true.
      ? monitoring.outputs.appInsightsConnectionString
      : ''
    tags: effectiveTags
  }
}

// ── Azure Maps account (optional; used by inline map preview in SPFx card) ───
resource azureMapsAccount 'Microsoft.Maps/accounts@2023-06-01' = if (deployAzureMapsEnabled) {
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

// ── Outputs ──────────────────────────────────────────────────────────────────

@description('The base URL of the deployed Function App. Paste this into the SPFx web part property pane (Azure Function Base URL field).')
output functionAppUrl string = functionApp.outputs.functionAppUrl

@description('The full function endpoint URL — use this for curl/Postman testing or health checks. The web part property pane only needs the base URL (functionAppUrl).')
output sponsorApiEndpointUrl string = functionApp.outputs.sponsorApiEndpointUrl

@description('Object ID of the system-assigned Managed Identity — needed for setup-graph-permissions.ps1.')
output managedIdentityObjectId string = functionApp.outputs.managedIdentityObjectId

@description('Name of the Application Insights component (empty when enableMonitoring=false).')
#disable-next-line BCP318
output appInsightsName string = enableMonitoringEnabled ? monitoring.outputs.appInsightsName : ''

@description('The selected hosting plan — included in outputs for operational visibility.')
output hostingPlan string = functionApp.outputs.hostingPlan

@description('Azure Maps account name (empty when deployAzureMaps=false).')
output azureMapsAccountName string = deployAzureMapsEnabled ? azureMapsAccount.name : ''

@description('Azure CLI command to fetch the Azure Maps primary key (empty when deployAzureMaps=false).')
output azureMapsKeyCommand string = deployAzureMapsEnabled
  ? 'az maps account keys list -g ${resourceGroup().name} -n ${azureMapsAccount.name} --query primaryKey -o tsv'
  : ''

@description('Name of the Storage Account. For Flex Consumption: upload updated ZIPs to the app-package container here to trigger a redeployment (use --auth-mode login). For Consumption: the runtime uses this account for trigger state and blob/queue/table operations.')
output deploymentStorageAccountName string = functionApp.outputs.deploymentStorageAccountName

@description('The function package version deployed. "latest" = newest release at provisioning time; otherwise the pinned SemVer tag without "v" prefix.')
output deployedAppVersion string = appVersion

@description('The Function App name (auto-generated or explicitly supplied).')
output functionAppName string = functionApp.outputs.functionAppName

// ── Entra App Registration ────────────────────────────────────────────────────
// Creates (or idempotently updates) the App Registration used by EasyAuth as
// the identity provider. The web part acquires delegated tokens against this
// audience using the pre-authorized SharePoint Online Web Client Extensibility
// client — no user consent prompt is shown.
//
// uniqueName acts as the stable key across re-deployments; it must be globally
// unique in the tenant.  We incorporate functionAppName which is already
// required to be globally unique across Azure.
//
// NOTE: migration from the old pre-provision-hook-created App Registration:
// That App Registration had no uniqueName set, so Bicep will create a new one
// alongside it on the first run with this template.  Delete the old registration
// in the Entra admin center after confirming the new one is working.
resource appReg 'Microsoft.Graph/applications@v1.0' = {
  displayName: 'Guest Sponsor Info - SharePoint Web Part Auth'
  uniqueName: 'guest-sponsor-info-proxy-${functionAppName}'
  description: appRegDescription
  notes: appRegNotes
  signInAudience: 'AzureADMyOrg'
  info: appRegInfo
  logo: appRegLogo
  serviceManagementReference: '${repositoryUrl}/issues'
  web: {
    homePageUrl: repositoryUrl
  }
  // identifierUri uses functionAppName (globally unique) instead of appId to avoid
  // a circular reference (appId is read-only and not available inside the same resource).
  identifierUris: [
    'api://guest-sponsor-info-${functionAppName}'
  ]
  api: {
    requestedAccessTokenVersion: 2
    oauth2PermissionScopes: [
      {
        // Deterministic GUID based on the Function App name — stable across re-deployments.
        id: userImpersonationScopeId
        adminConsentDescription: 'Allows the SharePoint web part to call the Azure Function proxy on behalf of the signed-in user.'
        adminConsentDisplayName: 'Access Guest Sponsor Info web part proxy as the signed-in user'
        isEnabled: true
        type: 'User'
        userConsentDescription: 'Allows the app to call the Azure Function proxy on your behalf.'
        userConsentDisplayName: 'Access Guest Sponsor Info web part proxy'
        value: 'user_impersonation'
      }
    ]
    // Pre-authorize the SharePoint Online Web Client Extensibility client app
    // so the web part can acquire tokens silently without a user consent prompt.
    preAuthorizedApplications: [
      {
        appId: '08e18876-6177-487e-b8b5-cf950c1e598c' // SharePoint Online Web Client Extensibility
        delegatedPermissionIds: [
          userImpersonationScopeId
        ]
      }
    ]
  }
}

var effectiveAppClientId = appReg.appId

// ── EasyAuth App Registration Service Principal ───────────────────────────────
// Configure the Enterprise Application (Service Principal) created automatically
// when the App Registration is deployed. Sets appRoleAssignmentRequired=false so
// all users (including guests) can acquire tokens without individual assignment,
// and hides the app from the My Apps portal (HideApp tag).
// Requires: Application.ReadWrite.All (Cloud Application Administrator).
resource easyAuthSp 'Microsoft.Graph/servicePrincipals@v1.0' = {
  appId: appReg.appId
  tags: ['HideApp']
  appRoleAssignmentRequired: false
  description: appRegDescription
  notes: appRegNotes
}

// ── Graph permissions module ──────────────────────────────────────────────────
// Assigns the four required Microsoft Graph application permissions to the
// Managed Identity. Isolated in a module so security reviewers can audit the
// exact permission grants independently of the rest of the deployment.
module graphPermissions './modules/graph-permissions.bicep' = {
  params: {
    managedIdentityObjectId: functionApp.outputs.managedIdentityObjectId
    skipRoleAssignments: skipRoleAssignments
  }
}

@description('Client ID (appId) of the Entra App Registration used for EasyAuth. Paste this into the SPFx web part property pane (Application (client) ID field).')
output webPartClientId string = effectiveAppClientId
