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
  - Loads profile photos directly from Graph (delegated token, progressive)

[Presence polling – ongoing, separate from initial load]
  - Web part polls Guest Sponsor API at adaptive intervals:
      30 s while a card is hovered · 2 min tab visible · 5 min tab hidden
  - Same Bearer token (silently refreshed) · same EasyAuth gate
  - Function returns presence status only — sponsor list is never re-fetched
```

No Entra directory role needed for the guest. The function is the only party that
holds `User.Read.All`; the guest never sees that permission.

### Direct Graph (legacy fallback)

When no Guest Sponsor API URL is configured, the web part calls `GET /v1.0/me/sponsors` directly.
Requires the guest to hold an Entra directory role (e.g. Directory Readers) — see README.

## Graph Permissions

### Function (application, via Managed Identity)

| Permission | Purpose |
|---|---|
| `User.Read.All` | `/users/{oid}/sponsors`, `$batch` profile checks, `accountEnabled` |
| `Presence.Read.All` | **Optional.** `/communications/getPresencesByUserId`. Requires Teams licensing. Skipped when absent — sponsors render without presence indicator. |
| `MailboxSettings.Read` | **Optional.** Filter shared/room/equipment mailboxes. Skipped when absent. |

### Direct path (delegated, via SharePoint API access panel)

| Permission | Purpose |
|---|---|
| `User.Read` | `/me/sponsors`. Also requires Directory Readers role. |
| `User.ReadBasic.All` | Existence checks (`/users/{id}`), profile photos |
| `Presence.Read.All` | **Optional.** Presence status for sponsor cards. Skipped when not consented. |

### Why no `User.Read.All` on the delegated path

Reading `accountEnabled` requires `User.Read.All`. On the delegated path we avoid this
scope and instead probe with `GET /users/{id}?$select=id` — HTTP 404 = deleted, 200 =
still exists. Disabled-but-not-deleted sponsors remain visible until hard-deleted.
The Guest Sponsor API path does not have this limitation.

## Profile Photos

Always fetched client-side (delegated token) via `/users/{id}/photo/$value`, even when
the Guest Sponsor API is used for sponsor/presence data. Returned as `ArrayBuffer` → base64
data URL to avoid `Blob` URL leaks. Failed photo requests fall back to initials silently.

## Presence Display

Both `availability` and `activity` are read from Graph. Display labels follow
[Microsoft's documented combination table](https://learn.microsoft.com/en-us/graph/cloud-communications-manage-presence-state):
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
  This is defense-in-depth on top of EasyAuth issuer/audience validation.
- **CORS** restricted to the tenant's SharePoint origin.
- **No secrets stored;** Managed Identity for Graph and storage access (RBAC, no keys).
- **Caller OID redacted** in function logs; client validation failures include structured reason
  codes and diagnostics without exposing full GUIDs.
- **Rate limit:** Two tiers — anonymous callers (no OID, e.g. dev without EasyAuth) are always
  rate-limited at 10 req / 60 s per IP. In production EasyAuth blocks anonymous callers at the
  infra level before function code runs. Authenticated callers are not rate-limited by default;
  set `RATE_LIMIT_ENABLED=true` (with optional `RATE_LIMIT_MAX_REQUESTS` / `RATE_LIMIT_WINDOW_MS`)
  to activate per-user limiting as an incident-response measure.

### Data Filtering

- Sponsors: must resolve in Graph active view, `accountEnabled !== false`.
- Managers: same rules.
- Mailbox filter (when `MailboxSettings.Read` granted): exclude `userPurpose` other
  than `user` / `linked`. Without the permission the filter is skipped (fail-open).
- Excluded sponsors/managers are counted in `unavailableCount`.
- Max 5 sponsors enforced at the Graph query level.

### Runtime

Three concurrent Graph requests per invocation: sponsor lookup, presence, `$batch`
(profile + manager via `$expand=manager`). Photos are not fetched by the function.

Timeout app settings: `SPONSOR_LOOKUP_TIMEOUT_MS`, `BATCH_TIMEOUT_MS`,
`PRESENCE_TIMEOUT_MS`. Presence/manager degrade gracefully; sponsor lookup failure → 504.

### Deployment (`azd up`)

1. **Pre-provision** — creates/reuses EasyAuth App Registration, detects SharePoint tenant.
2. **Bicep** — provisions storage (RBAC, no keys), Function App with MI + EasyAuth,
   Log Analytics, App Insights.
3. **Post-provision** — grants Graph permissions to MI, prints API URL + Client ID.

> RBAC propagation can take 1–2 min after deploy. Wait and retry if errors appear
> immediately. Deployer needs **Owner** role on the resource group.

Manual fallback: `infra/setup-app-registration.ps1` + `infra/setup-graph-permissions.ps1`.

### Hosting Plan Options

The `hostingPlan` parameter controls the Azure Functions pricing tier:

| | **Consumption** (default) | **Flex Consumption** (opt-in) |
|---|---|---|
| SKU | Y1 / Dynamic | FC1 / FlexConsumption |
| Free tier | 1M exec + 400K GB-s/month | None |
| Cold starts | ~2–5 s after ~20 min idle | Greatly reduced; eliminated with `alwaysReadyInstances=1` |
| OS | Windows | Linux only |
| ZIP deployment | `WEBSITE_RUN_FROM_PACKAGE` (GitHub URL) | Blob container (AZD / az CLI) |
| "Deploy to Azure" button | Supported | Not supported |
| Cost guard | `dailyMemoryTimeQuota` (GB-s budget) | `maximumFlexInstances` (hard instance cap, required) |
| Estimated cost | Free (within grant) | ~€2–5/month with 1 warm instance |

**Default is Consumption** — it covers the typical use case (internal SPFx tool, hundreds of
guest users/day) at zero cost and supports the simplest deployment paths.

Choose **Flex Consumption** when cold-start latency is unacceptable for your users and you
are deploying via AZD or Azure CLI. Set `alwaysReadyInstances=1` (the default for Flex) to
keep one instance warm.

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

**Operational alerting:** Bicep deploys three optional Azure Monitor Scheduled Query alerts
(`Microsoft.Insights/scheduledQueryRules`):

- Service outage operational email alert (`enableServiceOutageAlert`)
- Auth/config regression operational email alert (`enableAuthConfigRegressionAlert`)
- Likely attack/noise info alert (`enableLikelyAttackInfoAlert`)

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

`isDomainIsolated: true` is intentionally not used — it has known issues with guest token
acquisition.

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

- **As member:** `SPFX_TENANT` in `.env` + `./scripts/dev-webpart.sh`. Verifies non-guest path.
- **As guest:** Requires a second M365 tenant with your account as guest, sponsors assigned,
  API permissions consented, `.sppkg` deployed or localhost script loading enabled.

### Unit tests

`npm test` — covers guest detection, Graph service calls (mocked), and component rendering.

### Demo mode

Property pane toggle. Shows two fictitious sponsors without Graph calls.
Development/visual review only — disable before production.
