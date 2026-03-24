metadata name = 'Guest Sponsor Info – Monitoring'
metadata description = 'Log Analytics Workspace, Application Insights, Action Groups, and KQL alert rules.'

@description('Azure region for all monitoring resources.')
param location string

@description('Base name used to derive resource names (must match functionAppName in the parent template).')
param functionAppName string

@description('Resource tags to apply to all monitoring resources.')
param tags object

// ── Alert feature flags ──────────────────────────────────────────────────────

@description('Enable operational email alert for probable service outage (5xx/504 spike or low success rate).')
param enableServiceOutageAlert bool = true

@description('Enable operational email alert for auth/config regressions (AUTH_CONFIG_* reason codes).')
param enableAuthConfigRegressionAlert bool = true

@description('Enable info-only alert for likely attack/noise spikes (high 401/403 from many IPs).')
param enableLikelyAttackInfoAlert bool = true

@description('Enable info-only alert when a newer GitHub release of the function is available.')
param enableNewReleaseAlert bool = true

@description('Enable operational alert when a hard-deleted Entra object remains referenced as a sponsor (Graph 404).')
param enableBrokenSponsorAlert bool = false

// ── Alert timing ─────────────────────────────────────────────────────────────

@description('KQL alert evaluation frequency in minutes.')
@minValue(1)
param alertEvaluationFrequencyInMinutes int = 5

@description('KQL alert lookback window in minutes.')
@minValue(5)
param alertWindowInMinutes int = 15

// ── Service outage thresholds ────────────────────────────────────────────────

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

// ── Auth config regression threshold ────────────────────────────────────────

@description('AUTH_CONFIG_* trace count threshold for config-regression alert.')
@minValue(1)
param authConfigRegressionHitsThreshold int = 1

// ── Likely-attack thresholds ─────────────────────────────────────────────────

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

// ── New-release check alert thresholds ───────────────────────────────────────

@description('KQL evaluation frequency for the new-release alert in minutes.')
@minValue(5)
param newReleaseAlertEvaluationFrequencyInMinutes int = 60

@description('KQL lookback window for the new-release alert in minutes. Must be at least twice the function timer interval (6 h) to tolerate one missed timer invocation.')
@minValue(60)
param newReleaseAlertWindowInMinutes int = 720

// ── Action groups ────────────────────────────────────────────────────────────

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

// ── Derived resource names ───────────────────────────────────────────────────

var logAnalyticsWorkspaceName = '${functionAppName}-logs'
var appInsightsResourceName = '${functionAppName}-insights'

// ── KQL queries ──────────────────────────────────────────────────────────────
// Triple-quoted raw strings (no interpolation); placeholders are substituted
// with replace() so that the final query is a plain string in the ARM template.

// Matches the structured WARNING logged by the checkGitHubRelease timer trigger.
// The `latestVersion` column is used as an alert dimension so that each distinct
// GitHub release version creates an independent alert instance.
// autoMitigate: true — the instance resolves automatically once the function is
// updated (no more [NEW_RELEASE_AVAILABLE] traces for that latestVersion value).
// When the function is then updated but a subsequent newer release appears, the
// same mechanism fires a fresh notification — one email per unique latestVersion.
var newReleaseAlertQueryRaw = '''
let window = __WINDOW__m;
traces
| where timestamp > ago(window)
| where message has "[NEW_RELEASE_AVAILABLE]"
| extend latestVersion = extract(@"latestVersion=(\\d+\\.\\d+\\.\\d+)", 1, message)
| where isnotempty(latestVersion)
| project latestVersion
'''
#disable-next-line prefer-interpolation
var newReleaseAlertQuery = replace(newReleaseAlertQueryRaw, '__WINDOW__', string(newReleaseAlertWindowInMinutes))

// Matches the structured WARNING logged by getGuestSponsors when a sponsor object
// returns Graph HTTP 404 (hard-deleted Entra account still referenced as a sponsor).
// Dimension split on sponsorId so each distinct broken reference fires and
// auto-mitigates independently — no alert noise when unrelated guests are affected.
var brokenSponsorAlertQueryRaw = '''
let window = __WINDOW__m;
traces
| where timestamp > ago(window)
| where message has "[BROKEN_SPONSOR_REF]"
| extend sponsorId = extract(@"sponsorId=([a-f0-9]{8}\\.\\.\\.[a-f0-9]{4})", 1, message)
| where isnotempty(sponsorId)
| project sponsorId
'''
#disable-next-line prefer-interpolation
var brokenSponsorAlertQuery = replace(brokenSponsorAlertQueryRaw, '__WINDOW__', string(alertWindowInMinutes))

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

// ── Default Action Groups (auto-created when defaultAlertNotificationEmail is set) ──

resource defaultOperationalActionGroup 'Microsoft.Insights/actionGroups@2023-01-01' = if (createDefaultActionGroups) {
  name: '${functionAppName}-ops-ag'
  location: 'global'
  tags: tags
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
  tags: tags
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

var effectiveOperationalActionGroupIds = createDefaultActionGroups
  ? concat(operationalActionGroupResourceIds, [defaultOperationalActionGroup.id])
  : operationalActionGroupResourceIds

var effectiveInfoActionGroupIds = createDefaultActionGroups
  ? concat(infoActionGroupResourceIds, [defaultInfoActionGroup.id])
  : infoActionGroupResourceIds

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
  name: appInsightsResourceName
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

// ── KQL Alert Rules ──────────────────────────────────────────────────────────

resource serviceOutageAlert 'Microsoft.Insights/scheduledQueryRules@2021-08-01' = if (enableServiceOutageAlert) {
  name: '${functionAppName}-service-outage-kql'
  location: location
  tags: tags
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
      actionGroups: effectiveOperationalActionGroupIds
    }
    autoMitigate: true
  }
}

resource authConfigRegressionAlert 'Microsoft.Insights/scheduledQueryRules@2021-08-01' = if (enableAuthConfigRegressionAlert) {
  name: '${functionAppName}-auth-config-regression-kql'
  location: location
  tags: tags
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
      actionGroups: effectiveOperationalActionGroupIds
    }
    autoMitigate: true
  }
}

resource likelyAttackInfoAlert 'Microsoft.Insights/scheduledQueryRules@2021-08-01' = if (enableLikelyAttackInfoAlert) {
  name: '${functionAppName}-likely-attack-info-kql'
  location: location
  tags: tags
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
      actionGroups: effectiveInfoActionGroupIds
    }
    autoMitigate: true
  }
}

resource newReleaseAlert 'Microsoft.Insights/scheduledQueryRules@2021-08-01' = if (enableNewReleaseAlert) {
  name: '${functionAppName}-new-release-kql'
  location: location
  tags: tags
  properties: {
    description: 'Info-only alert when a newer GitHub release of the Guest Sponsor Info function is available. Fires once per unique version; auto-mitigates when the function is updated.'
    enabled: true
    scopes: [
      appInsights.id
    ]
    evaluationFrequency: 'PT${newReleaseAlertEvaluationFrequencyInMinutes}M'
    windowSize: 'PT${newReleaseAlertWindowInMinutes}M'
    // Severity 4 = Verbose / informational — lowest possible severity.
    severity: 4
    criteria: {
      allOf: [
        {
          query: newReleaseAlertQuery
          timeAggregation: 'Count'
          operator: 'GreaterThan'
          threshold: 0
          // Dimension split: one alert instance per unique latestVersion value.
          // Azure Monitor fires a new notification whenever a previously unseen
          // latestVersion appears and auto-mitigates each instance independently.
          dimensions: [
            {
              name: 'latestVersion'
              operator: 'Include'
              values: ['*']
            }
          ]
          failingPeriods: {
            numberOfEvaluationPeriods: 1
            minFailingPeriodsToAlert: 1
          }
        }
      ]
    }
    actions: {
      actionGroups: effectiveInfoActionGroupIds
    }
    autoMitigate: true
  }
}

resource brokenSponsorAlert 'Microsoft.Insights/scheduledQueryRules@2021-08-01' = if (enableBrokenSponsorAlert) {
  name: '${functionAppName}-broken-sponsor-ref-kql'
  location: location
  tags: tags
  properties: {
    description: 'Operational alert when a hard-deleted Entra object remains referenced as a sponsor (Graph 404). Fires per unique pseudonymized sponsor OID; auto-mitigates when the reference is resolved.'
    enabled: true
    scopes: [
      appInsights.id
    ]
    evaluationFrequency: 'PT${alertEvaluationFrequencyInMinutes}M'
    windowSize: 'PT${alertWindowInMinutes}M'
    // Severity 3 = Warning — operational issue requiring admin attention but not an outage.
    severity: 3
    criteria: {
      allOf: [
        {
          query: brokenSponsorAlertQuery
          timeAggregation: 'Count'
          operator: 'GreaterThan'
          threshold: 0
          // Dimension split: one alert instance per unique pseudonymized sponsorId.
          // Alerts resolve independently once the broken reference is cleaned up.
          dimensions: [
            {
              name: 'sponsorId'
              operator: 'Include'
              values: ['*']
            }
          ]
          failingPeriods: {
            numberOfEvaluationPeriods: 1
            minFailingPeriodsToAlert: 1
          }
        }
      ]
    }
    actions: {
      actionGroups: effectiveOperationalActionGroupIds
    }
    autoMitigate: true
  }
}

// ── Outputs ──────────────────────────────────────────────────────────────────

@description('Application Insights connection string for use in the Function App settings.')
output appInsightsConnectionString string = appInsights.properties.ConnectionString

@description('Resource ID of the Application Insights component (used as alert scope).')
output appInsightsId string = appInsights.id

@description('Name of the Application Insights component — open in the Azure Portal for live telemetry.')
output appInsightsName string = appInsights.name
