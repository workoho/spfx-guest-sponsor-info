# Architecture and Design Decisions

Project-specific decisions and known limitations.
For installation and build instructions see the [README](../README.md).
For deployment details see [deployment.md](deployment.md).
For development setup see [development.md](development.md).
For a visual system overview see [architecture-diagram.md](architecture-diagram.md).

## SPFx Lifecycle and Non-Blocking Initialization

`onInit()` resolves immediately after `super.onInit()` and icon registration.
Graph and AAD HTTP client acquisition runs in the background via `_acquireClientsInBackground()`.

**Why:** SPFx awaits the `onInit()` Promise before rendering any web part on the page.
Blocking here with `getClient()` calls would delay the entire page layout, not just this
web part. By resolving `onInit()` immediately, the page renders all web parts in parallel.

**Render sequence for a guest user (view mode):**

1. SPFx calls `render()` right after `onInit()` resolves — both clients are still `undefined`.
   The React component initialises `loading = true` and the shimmer is immediately visible.
2. `_acquireClientsInBackground()` uses `Promise.allSettled` to wait for both clients
   concurrently, then calls `render()` once — passing real clients in a single props update.
3. The `useEffect` in `GuestSponsorInfo` detects the new client props and starts the data fetch.
4. Sponsors load; shimmer replaced by sponsor cards.

For non-guests `render()` returns `null` immediately in both step 1 and step 2 — no visible
effect. Edit-mode shows the placeholder in step 1 and re-renders with the proxy health-check
client in step 2.

## Guest Detection

Combined check: `isGuest = isExternalGuestUser || loginName.includes('#EXT#')`.

`isExternalGuestUser` is the primary signal (from the Entra token, synchronous, no Graph
call). The `#EXT#` fallback covers edge cases where the flag is not yet populated.

**Known limitation:** On a guest's very first visit, `loginName` can be empty because
the SharePoint user profile hasn't been provisioned yet. `isExternalGuestUser` is
unaffected.

## Data Paths

### Guest Sponsor API (recommended)

```text
[Guest Browser]
      │  ① acquires Bearer token from Entra ID
      │     (scoped to the EasyAuth App Registration)
      ▼
[Entra ID]
      │  returns signed token (identifies the guest)
      ▼
[Guest Browser]
      │  ② calls Guest Sponsor API with Bearer token attached
      ▼
[Azure Function – EasyAuth gate]
  - Validates token before any function code runs
  - Invalid / missing token → HTTP 401, request rejected here
  - Valid token → injects caller OID as X-MS-CLIENT-PRINCIPAL-ID
      ▼
[Azure Function – business logic]
  - Calls Graph via Managed Identity (app permissions)
  - Returns { activeSponsors, unavailableCount }
      │
[Guest Browser – SPFx Web Part]
  - Renders sponsor cards (one-time on page load)
  - Loads manager photos via Azure Function photo proxy (`/api/getPhoto`)
    (sponsor photos are already embedded in the function response)

[Presence polling – ongoing, separate from initial load]
  - Web part polls Guest Sponsor API at adaptive intervals:
      30 s while a card is hovered · 2 min tab visible · 5 min tab hidden
  - Same Bearer token (silently refreshed) · same EasyAuth gate
  - Function returns presence status only — sponsor list is never re-fetched
```

No Entra directory role needed for the guest. The Azure Function is the only party that
holds `User.Read.All`; the guest never sees that permission.

## Graph Permissions

### Function (application, via Managed Identity)

| Permission | Purpose |
|---|---|
| `User.Read.All` | `/users/{oid}/sponsors`, `$batch` profile checks, `accountEnabled` |
| `Presence.Read.All` | **Optional.** `/communications/getPresencesByUserId`. Requires Teams licensing. Skipped when absent — sponsors render without presence indicator. |
| `MailboxSettings.Read` | **Optional.** Filter shared/room/equipment mailboxes. Skipped when absent. |
| `TeamMember.Read.All` | **Optional.** Detect whether the guest has a Teams account via `/users/{id}/joinedTeams`. Skipped when absent — Teams chat/call buttons default to enabled. |

## Profile Photos

Sponsor photos are returned inline by the Azure Function as part of the
`getSponsorsViaProxy` response. Manager photos are fetched progressively
via the Function's `/api/getPhoto` proxy endpoint — the web part calls this
with the same `AadHttpClient` used for all other API requests. Returned as
a base64 data URL. Failed photo requests fall back to initials silently.

## Presence Display

Both `availability` and `activity` are read from Graph. Display labels follow
[Microsoft's documented combination table](https://learn.microsoft.com/graph/cloud-communications-manage-presence-state):
`activity` takes priority when it differs from `availability` (e.g. `Busy`/`InAMeeting` → "In a meeting").
All documented tokens are resolved via localised strings; undocumented tokens fall back to a
PascalCase word-splitter (English only).

**OutOfOffice as suffix modifier.** When `activity === 'OutOfOffice'`, the label is the
base `availability` label plus a localised suffix (e.g. "Available, out of office").
If no base availability is set, it falls back to the standalone "Out of office" string.
The dot colour uses the OOF magenta (`#B4009E`) regardless of the base availability.

**Focusing colour.** `Focusing` uses Teams purple (`#6264A7`), not the generic DND red.
This matches the colour Teams displays for focus sessions.

Presence is polled with adaptive intervals: 30 s when a sponsor card is actively hovered,
2 min when the browser tab is visible, 5 min when hidden.

## Azure Function

### Why the proxy exists

`/me/sponsors` requires the calling user to hold an Entra directory role. Assigning roles
to guests at scale is impractical (role-assignable groups have no dynamic membership,
automation requires `RoleManagement.ReadWrite.Directory`, and the sponsor-read permission
is not self-scoped — GDPR concern). The function sidesteps all of this.

### Security

- **User identification:** Caller OID comes from the
  [EasyAuth](https://learn.microsoft.com/azure/app-service/overview-authentication-authorization)-validated
  token — callers cannot query other users. The function never accepts a user ID
  from the request body or query string.
- **Client authorization:** In production ( `NODE_ENV === 'production'`), the function validates
  EasyAuth principal claims (not raw Authorization headers):
  1. `tid` (tenant ID) — must match our tenant (`TENANT_ID` env var)
  2. `aud` (audience) — must match our API client ID (`ALLOWED_AUDIENCE` env var)
  3. `appid` or `azp` (calling application) — must match the SharePoint Online Web Client
     Extensibility application
  This is defense-in-depth on top of EasyAuth issuer/audience validation.
- **Scoped follow-up authorization:** `getPresence` and `getPhoto` only serve sponsor or manager IDs
  authorized for the current caller. Preferred path is a short-lived HMAC-signed `X-Presence-Token`;
  fallback is a live sponsor lookup and fail-safe drop or deny behavior.
- **CORS** restricted to the tenant's SharePoint origin.
- **No secrets stored;** Managed Identity for Graph and storage access (RBAC, no keys).
- **Client bundle posture:** The downloaded web part bundle contains no secrets. Treat optional
  client-side configuration such as the Azure Maps subscription key as visible to page viewers.
- **Caller OID redacted** in function logs; client validation failures include structured reason
  codes and diagnostics without exposing full GUIDs.
- **Rate limit:** Two tiers — anonymous callers (no OID, e.g. dev without EasyAuth) are always
  rate-limited at 5 req / 60 s per IP. In production EasyAuth blocks anonymous callers at the
  infra level before function code runs. Authenticated callers are not rate-limited by default;
  set `RATE_LIMIT_ENABLED=true` (with optional `RATE_LIMIT_MAX_REQUESTS` / `RATE_LIMIT_WINDOW_MS`)
  to activate per-user limiting as an incident-response measure.

For trust assumptions, residual risk, and hardening guidance, see
[security-assessment.md](security-assessment.md).

### Silent Token Acquisition and Pre-Authorization

SPFx web parts do not have their own OAuth client identity. When a web part
calls `AadHttpClientFactory.getClient()` to obtain a token for a custom API,
SharePoint delegates this to the
[SharePoint Online Web Client Extensibility](https://learn.microsoft.com/sharepoint/dev/spfx/use-aadhttpclient#known-issues)
first-party application (two well-known app IDs across SharePoint Online
environments). This application acts as the MSAL confidential client that
requests tokens from Entra ID on behalf of the signed-in user (OAuth 2.0
On-Behalf-Of flow).

For this to work **silently** (no consent prompt, no page redirect), the
target API — in our case the EasyAuth App Registration — must
[pre-authorize](https://learn.microsoft.com/entra/identity-platform/permissions-consent-overview#preauthorization)
the SharePoint client application for its `user_impersonation` scope.
Pre-authorization tells Entra: *"This client is already trusted; issue tokens
without prompting the user for consent."*

Without pre-authorization, Entra would require an interactive consent flow.
Inside an embedded web part iframe, an interactive prompt either triggers a
full page reload (to break out of the iframe for the redirect) or a blocked
popup — both of which break the user experience.

The `setup-graph-permissions.ps1` script (and the `azd` post-provision hook)
configure this automatically:

1. Expose a `user_impersonation` delegated scope on the App Registration.
2. Add the SharePoint Online Web Client Extensibility app ID as a
  pre-authorized application for that scope.
3. Set `appRoleAssignmentRequired = false` on the Service Principal so all
   tenant users (including guests) can acquire tokens without individual
   assignment.

### Data Filtering

- Sponsors: must resolve in Graph active view, `accountEnabled !== false`.
- Managers: same rules.
- Mailbox filter (when `MailboxSettings.Read` granted): exclude `userPurpose` other
  than `user` / `linked`. Without the permission the filter is skipped (fail-open).
- Excluded sponsors/managers are counted in `unavailableCount`.
- Max 5 sponsors enforced at the Graph query level.
- `sponsorOrder` — the array of all sponsor IDs in the original Graph response
  order — is always returned alongside `activeSponsors` and
  `unavailableSponsors`. The client uses it to reconstruct the full
  priority-ordered list and implement automatic delegation: active sponsors
  step into vacated visible slots while unavailable ones are still rendered as
  read-only tiles in their original position.

### Runtime

Three concurrent Graph requests per invocation: sponsor lookup, presence, `$batch`
(profile + manager via `$expand=manager`). Photos are not fetched by the function.

Timeout app settings: `SPONSOR_LOOKUP_TIMEOUT_MS`, `BATCH_TIMEOUT_MS`,
`PRESENCE_TIMEOUT_MS`. Presence/manager degrade gracefully; sponsor lookup failure → 504.

### Deployment (`azd up`)

1. **Pre-provision** — validates required subscription resource providers,
   registers missing ones when the caller has subscription-level register
   permission, creates/reuses EasyAuth App Registration, detects SharePoint tenant.
2. **Bicep** — provisions storage (RBAC, no keys), Function App with MI + EasyAuth,
   Log Analytics, App Insights.
3. **Post-provision** — grants Graph permissions to MI, prints API URL + Client ID.

> RBAC propagation can take 1–2 min after deploy. Wait and retry if errors appear
> immediately. Azure deployment needs **Contributor** plus **Owner** (or **User
> Access Administrator**) on the resource group. Entra permissions are split:
> **Cloud Application Administrator** for the App Registration and
> **Privileged Role Administrator** for Graph app-role assignment. See
> [deployment.md](deployment.md) for the operator-facing workflow.

Manual fallback: `infra/setup-graph-permissions.ps1` (for role assignment only; App Registration is always created by Bicep).

### Hosting Plan Options

The `hostingPlan` parameter controls the Azure Functions pricing tier:

| | **Consumption** (default) | **Flex Consumption** (opt-in) |
|---|---|---|
| SKU | Y1 / Dynamic | FC1 / FlexConsumption |
| Free tier | 1M exec + 400K GB-s/month | None |
| Cold starts | ~2–5 s after ~20 min idle | Greatly reduced; eliminated with `alwaysReadyInstances=1` |
| OS | Windows | Linux only |
| ZIP deployment | `WEBSITE_RUN_FROM_PACKAGE` (GitHub URL) | Blob container (AZD / az CLI) |
| Cost guard | `dailyMemoryTimeQuota` (GB-s budget) | `maximumFlexInstances` (hard instance cap, required) |
| Estimated cost | Free (within grant) | ~€2–5/month with 1 warm instance |

**Default is Consumption** — it covers the typical use case (internal SPFx tool, hundreds of
guest users/day) at zero cost and supports the simplest deployment paths.

Choose **Flex Consumption** when cold-start latency is unacceptable for your users and you
are deploying via AZD or Azure CLI. Set `alwaysReadyInstances=1` (the default for Flex) to
keep one instance warm.

For day-2 code updates, the two plans diverge operationally: Consumption can
switch or restart `WEBSITE_RUN_FROM_PACKAGE`, while Flex updates the package in
its blob-backed deployment container. See [operations.md](operations.md).

### Debugging Client Authorization Failures

When a client authorization failure occurs (HTTP 403), the function logs a structured warning
containing diagnostic details without exposing sensitive GUIDs. The logs include:

- `reasonCode` — stable machine-readable code for dashboards/alerts
- Reason — human-readable rejection detail
- `tid` when available
- `callerOid` (redacted) — first 8 + last 4 chars of the user OID

Function error responses also include a safe troubleshooting contract for the web part:

- `reasonCode` — machine-readable classification
- `retryable` — hint whether client retry is reasonable
- `referenceId` — correlation ID (`x-correlation-id` header) for support tickets

The web part logs these fields in browser console and appends `referenceId` to the user-facing
error text so operations can jump directly into backend logs without exposing sensitive internals.

In Application Insights, search for `"client-validation-failed"` in the custom dimensions to
find all authorization rejections.

**Common reasons:**

- `AUTH_PRINCIPAL_MISSING` — EasyAuth principal header missing/invalid
- `AUTH_CLAIM_MISSING_TID` — principal does not contain tenant claim
- `AUTH_TENANT_MISMATCH` — token was issued for another tenant
- `AUTH_CLAIM_MISSING_AUD` — principal does not contain audience claim
- `AUTH_AUDIENCE_MISMATCH` — token audience does not match `ALLOWED_AUDIENCE`
- `AUTH_CONFIG_TENANT_MISSING` / `AUTH_CONFIG_AUDIENCE_MISSING` — server misconfiguration

**Operational alerting:** Bicep deploys four optional Azure Monitor Scheduled Query alerts
(`Microsoft.Insights/scheduledQueryRules`):

- Service outage operational email alert (`enableServiceOutageAlert`)
- Auth/config regression operational email alert (`enableAuthConfigRegressionAlert`)
- Likely attack/noise info alert (`enableLikelyAttackInfoAlert`)
- New GitHub release available info alert (`enableNewReleaseAlert`) — see
  [Update and Release Management](#update-and-release-management)

Action-group wiring is parameterized via `operationalActionGroupResourceIds` and
`infoActionGroupResourceIds`.

For small deployments, Bicep can auto-create default action groups when
`defaultAlertNotificationEmail` is provided. In that case:

- `${functionAppName}-ops-ag` is added to operational email alerts
- `${functionAppName}-info-ag` is added to info alerts

Short names are configurable via `defaultOperationalActionGroupShortName` and
`defaultInfoActionGroupShortName`.

### Low-Cost KQL Alert Strategy (No WAF / Front Door)

To reduce false positives, we intentionally do not page on generic 4xx spikes alone.
Instead, we split alerting by actionability:

1. **Service outage (operational email)**
2. **Configuration/auth regression (operational email)**
3. **Likely attack/noise spike (info channel only)**

The implemented Bicep rules follow these design goals:

- **Service outage (operational email):** alerts on clear availability degradation (5xx/504 spike
  and/or significant success-rate drop).
- **Config/auth regression (operational email):** alerts when `AUTH_CONFIG_*` reason codes appear,
  because this indicates deployment/config drift that is immediately actionable.
- **Likely attack/noise (info):** alerts on strong 401/403 denial patterns from many
  sources, but routes to a non-paging channel.

Concrete KQL and thresholds are maintained in infrastructure code (`main.bicep`), not in
this architecture decision document.

This pattern avoids waking admins for events they cannot immediately mitigate without
edge protection, while still surfacing real service breakages.

**To bypass client validation in development,** set `NODE_ENV` to a value other than `production`
(e.g., `development`). In that mode, client authorization is skipped entirely, but user
identification via EasyAuth headers is still enforced.

### Local Development

```bash
cd azure-function
cp local.settings.json.example local.settings.json  # fill in TENANT_ID, ALLOWED_AUDIENCE etc.
npm install && npm start
# Pass guest OID via X-Dev-User-OID header (only accepted when NODE_ENV !== 'production').
# For bypass (dev mode), ensure NODE_ENV is NOT 'production':
NODE_ENV=development npm start
```

To test client authorization failures, call the deployed Function App (with EasyAuth enabled)
using a real access token acquired by SharePoint/SPFx:

```curl
curl -X GET https://<your-function>.azurewebsites.net/api/getGuestSponsors \
  -H "Authorization: Bearer <your-access-token>"
```

Do not send `x-ms-client-principal` or `x-ms-client-principal-id` manually.
In production these headers are emitted by EasyAuth after token validation.

If tenant or audience validation fails, the function returns HTTP 403 with `reasonCode`
and diagnostic details in logs/Application Insights.

## App Catalog Guest Access

Assets are bundled inside the `.sppkg`. Guest users cannot access re-hosted assets by
default (HTTP 403). Two solutions — see
[Deployment Guide – Guest Access Requirements](deployment.md#guest-access-requirements):

- **Public CDN (recommended):** Assets served from `publiccdn.sharepointonline.com`,
  no auth needed. Simpler, faster, no ongoing permission management.
- **App Catalog permissions:** Grant Read to *All Users (membership)* + enable external
  sharing on the App Catalog site. *Everyone except external users* does **not** work.

Sponsor photos are fetched in parallel by the Azure Function and bundled in the
`getGuestSponsors` response. Manager photos are fetched lazily via the `/api/getPhoto`
endpoint using `AadHttpClient`. No direct Microsoft Graph calls remain in the web
part — all data flows through the Azure Function proxy.

## UI Behaviour

| User type | Edit mode | View mode |
|---|---|---|
| Guest | Text placeholder (no Graph calls) | Sponsor cards |
| Non-guest | Text placeholder (different message) | Hidden (`null`) |

Initials avatars use Fluent UI `Persona` component with a project-specific 12-colour palette
(hex strings passed via `initialsColor`). Presence indicators are rendered natively by Persona's
`presence` and `isOutOfOffice` props for all standard Graph availability states; only *Focusing*
(no Fluent enum equivalent) uses a custom-positioned `<span>` in Teams purple (`#6264A7`).
Photo fade-in is handled by `Persona`'s `imageShouldFadeIn` prop.

Action buttons (Chat, Email, Call) use Fluent `ActionButton` with a column-stacked layout
via the `styles` prop. Copy-to-clipboard buttons use `IconButton`. Contact value links
(mailto/tel) use Fluent `Link` with the `::before` full-row click overlay retained in CSS.
Rich contact card shown via Callout (desktop) or Panel (mobile).

**Retained custom CSS** covers structural layout only: grid dimensions, card sizing,
rich card header/section layout, info row positioning, and the Focusing presence dot.
All colours reference SPFx theme tokens via `var()` + `"[theme:]"` dual declarations.

## Development Testing

### Hosted workbench

- **As member:** `SPFX_SERVE_TENANT_DOMAIN` in `.env` (or set on host OS) +
  `./scripts/dev-webpart.sh`. Verifies non-guest path.
- **As guest:** Requires a second M365 tenant with your account as guest, sponsors assigned,
  API permissions consented, `.sppkg` deployed or localhost script loading enabled.

### Unit tests

`npm test` — covers guest detection, Graph service calls (mocked), and component rendering.

### Demo mode

Property pane toggle. Shows two fictitious sponsors without Graph calls.
Development/visual review only — disable before production.

## Update and Release Management

### Web Part — Version Check in the Property Pane

When a site editor opens the property pane for the first time in a browser session,
the web part calls the GitHub Releases API in the background (fire-and-forget):

```text
GET https://api.github.com/repos/workoho/spfx-guest-sponsor-info/releases/latest
```

If the API returns a newer semver tag than the currently deployed web part version,
a dismissable badge appears in the property pane header linking directly to the
release notes page on GitHub.

**Design decisions:**

- **Lazy check on `onPropertyPaneOpened`.** The check is deferred until the property
  pane is opened for the first time. Non-editors (view-only page visitors) never
  trigger a GitHub API call.
- **In-memory flag (`_githubCheckDone`).** Re-opening the property pane within the
  same browser-tab session does not repeat the network call.
- **sessionStorage cache (1 h TTL).** The result (`{ version, url, ts }`) is persisted
  in `sessionStorage` so that navigating between edit/view mode or refreshing the
  page within the same tab does not re-fetch. The cache expires after 1 hour.
- **Namespaced key.** The sessionStorage key is prefixed with `this.manifest.id`
  (the web part component UUID) to avoid collisions with other web parts or
  extensions that may also use sessionStorage.
- **Only published releases.** The `/releases/latest` endpoint returns only published
  GitHub Releases with release notes — bare git tags are never included.
- **Link target.** The badge `href` uses the `html_url` field from the API response
  (the canonical release notes page) rather than a constructed tag URL.

### Azure Function — Periodic GitHub Release Check

A dedicated Timer-triggered Azure Function (`checkGitHubRelease`) runs every 6 hours
and immediately on every deployment or restart (`runOnStartup: true`). It calls the
same GitHub Releases API as the web part and compares the response against the
function's own `package.json` version.

**When the function is up to date** it logs an info-level trace:

```text
[checkGitHubRelease] Function is up to date (vX.Y.Z).
```

**When a newer release is found** it logs a structured WARNING to Application Insights:

```text
[NEW_RELEASE_AVAILABLE] currentVersion=X latestVersion=Y url=Z
```

This trace is the trigger signal for the Azure Monitor KQL alert rule described below.
The key=value tokens in this message are intentionally stable — changing them requires
a corresponding update to the `extract()` expression in `monitoring.bicep`.

API errors (network timeouts, HTTP 404, rate-limit responses) are logged at info level
only and do not fail the invocation or cause retries.

The timer always records the fetched GitHub version in a shared in-memory variable
(`releaseState.ts`). The request handler reads this cache to enrich version-mismatch
traces without making an additional GitHub API call.

### Request-Time Version-Mismatch Logging

The web part sends its own version as `X-Client-Version` on every API request.
When the function detects that this differs from its own version it emits a structured
WARNING — but **throttled to once per hour per function instance** to avoid flooding
Application Insights when many guests are active simultaneously. One trace per hour
per instance is sufficient for KQL alert windows of ≥ 2 h.

By the time the first mismatch trace is written, the timer (`runOnStartup: true`) has
almost always already completed its GitHub check and populated the in-memory cache.
The request handler cross-references the cached GitHub version to emit the most
specific token it can:

| Token | Condition | Meaning |
|---|---|---|
| `[WEBPART_UPDATE_AVAILABLE]` | function = GitHub latest, web part < latest | S3: only the web part is outdated |
| `[FUNCTION_UPDATE_AVAILABLE]` | web part = GitHub latest, function < latest | S4: only the function is outdated |
| `[VERSION_MISMATCH]` | GitHub cache empty, or both behind latest | Mismatch with `olderComponent` hint |

**Trace format:**

```text
[WEBPART_UPDATE_AVAILABLE] webPartVersion=X functionVersion=Y latestVersion=Y
[FUNCTION_UPDATE_AVAILABLE] functionVersion=X webPartVersion=Y latestVersion=Y
[VERSION_MISMATCH] functionVersion=X webPartVersion=Y olderComponent=webpart|function [latestVersion=Y]
```

`latestVersion` is omitted from `[VERSION_MISMATCH]` only in the rare cold-start window
between instance launch and the first completed timer run (typically < 10 s).

**Hierarchy of severity (informational, highest to lowest):**

1. `[VERSION_MISMATCH]` / `[WEBPART_UPDATE_AVAILABLE]` / `[FUNCTION_UPDATE_AVAILABLE]` —
   the two running components are incompatible *right now*; the admin knows which to update.
2. `[NEW_RELEASE_AVAILABLE]` / `[WEBPART_UPDATE_AVAILABLE]` / `[FUNCTION_UPDATE_AVAILABLE]` —
   a newer GitHub release exists but nothing is broken yet.

These tokens are not currently backed by dedicated KQL alert rules (no Bicep resource).
They are visible in Application Insights logs and can be used for ad-hoc KQL queries or
as the basis for future alert rules if needed.

### Azure Monitor KQL Alert Rule (New-Release Notification)

The Bicep module `monitoring.bicep` provisions an optional
`Microsoft.Insights/scheduledQueryRules` resource that watches Application Insights for
`[NEW_RELEASE_AVAILABLE]` traces. Its purpose is to send **one informational notification
per GitHub release version** to the configured info action group, with automatic resolution
once the function is updated.

**KQL query (simplified):**

```kql
let window = 720m;
traces
| where timestamp > ago(window)
| where message has "[NEW_RELEASE_AVAILABLE]"
| extend latestVersion = extract(@"latestVersion=(\d+\.\d+\.\d+)", 1, message)
| where isnotempty(latestVersion)
| project latestVersion
```

**Alert properties:**

| Property | Value | Rationale |
|---|---|---|
| Severity | 4 (Verbose / Informational) | Lowest severity; informational only, no paging |
| Evaluation frequency | 60 min | One notification quickly after a new release appears |
| Lookback window | 720 min (12 h) | Covers two 6-hour timer intervals to tolerate one missed run |
| `autoMitigate` | `true` | Alert instance resolves automatically when traces stop |
| Dimension split | `latestVersion` | One independent alert instance per GitHub release version |

**Why `autoMitigate: true` + dimension split is the right architecture:**

Each unique `latestVersion` value produces an independent alert instance in Azure Monitor.
This means:

1. **No duplicate notifications.** While v1.5.0 is pending, the same `latestVersion`
   dimension value keeps **the existing** instance "fired" — Azure Monitor suppresses
   further notifications for it.
2. **Automatic resolution after update.** Once the function is updated to v1.5.0+, the
   timer stops emitting `[NEW_RELEASE_AVAILABLE]` with `latestVersion=1.5.0`. After the
   12-hour lookback window elapses with no matching traces, `autoMitigate` resolves the
   alert instance for that version — no manual intervention needed.
3. **Cascading release handling.** If v1.6.0 is published while the v1.5.0 notification
   is still pending, the function starts emitting `latestVersion=1.6.0` immediately. The
   v1.6.0 dimension creates a **new** alert instance and fires its own notification
   independently. The v1.5.0 instance auto-mitigates after its window elapses.

**Bicep parameters:**

| Parameter | Default | Description |
|---|---|---|
| `enableNewReleaseAlert` | `true` | Deploy the alert rule |
| `newReleaseAlertEvaluationFrequencyInMinutes` | `60` | How often Azure Monitor evaluates the KQL |
| `newReleaseAlertWindowInMinutes` | `720` | Lookback window for the KQL query |

The alert is routed to the info action group (`infoActionGroupResourceIds` /
`defaultAlertNotificationEmail`), the same group used by the
likely-attack info alert.
