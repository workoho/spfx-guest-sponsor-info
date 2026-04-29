# Guest Sponsor API for Microsoft Entra B2B

<p align="center">
  <img src="../images/icon-rounded.svg" width="96" height="96" alt="Guest Sponsor Info icon" />
</p>

Part of [**Guest Sponsor Info**](../README.md) — a SharePoint Online web part
for guest landing pages in Microsoft Entra resource tenants.

HTTP-triggered Azure Function (Node.js 22, Azure Functions v4) that acts as a
proxy between the SharePoint web part and Microsoft Graph.

For deployment instructions and administration, see
[docs/deployment.md](../docs/deployment.md#guest-sponsor-api).
For architecture decisions, see
[docs/architecture.md](../docs/architecture.md#azure-function).
For a visual system overview, see
[docs/architecture-diagram.md](../docs/architecture-diagram.md).

## Why this exists

The Microsoft Graph `/me/sponsors` endpoint requires the calling user to hold
an Entra directory role. Assigning a directory role to every B2B guest at
scale is impractical:

- Role-assignable groups have no dynamic membership support.
- The `microsoft.directory/users/sponsors/read` permission is not self-scoped
  — a guest with that role can also read other guests' sponsor relationships.
- Automation requires `RoleManagement.ReadWrite.Directory`.

The function sidesteps all of this: it calls Graph with **application
permissions** via a Managed Identity, enforces server-side that it returns
only the calling user's own sponsors, and guests never need any directory role.

## Permissions

### Caller → Function (delegated, via [EasyAuth](https://learn.microsoft.com/azure/app-service/overview-authentication-authorization) App Registration)

The web part acquires a token for the EasyAuth App Registration and passes it
as a Bearer token. EasyAuth validates the token before function code runs.

| Scope | Type | Purpose |
|---|---|---|
| `user_impersonation` | Delegated | Identifies the calling user. Exposed on the App Registration; pre-authorized for *SharePoint Online Web Client Extensibility* so the web part can acquire tokens silently. |

The function never trusts anything from the request body or query string for
identity. The caller OID is read exclusively from the EasyAuth-validated
`X-MS-CLIENT-PRINCIPAL-ID` header.

### Function → Microsoft Graph (application, via Managed Identity)

No stored credentials. The Function App uses `DefaultAzureCredential` which
resolves to the system-assigned Managed Identity in Azure. For local
development it falls back to Azure CLI credentials (`az login`) or
environment-variable service-principal credentials.

#### Required

| Permission | Purpose |
|---|---|
| `User.Read.All` | Read any user's profile (`/users/{oid}`), sponsor list (`/users/{oid}/sponsors`), and `accountEnabled` status. |

#### Optional

| Permission | Purpose | Behaviour when absent |
|---|---|---|
| `Presence.Read.All` | Fetch online presence status for sponsor cards. Requires Microsoft Teams licensing in the tenant. | Sponsors are shown without presence indicator. |
| `MailboxSettings.Read` | Read `userPurpose` to filter out shared/room/equipment mailboxes from sponsor cards. | Filter is skipped — all mailbox types are included (fail-open). |
| `TeamMember.Read.All` | Check if the guest user is a member of any Microsoft Team (used to show/hide Teams action buttons on sponsor cards). | Falls back to presence check without Teams membership signal. |

The function detects which optional permissions are granted at cold-start by
inspecting the Managed Identity token and logs the result to Application
Insights. Missing optional permissions degrade gracefully; a missing required
permission produces an error log entry and all sponsor lookups return HTTP 503.

### Validation checks on every request

In production (`NODE_ENV=production`) the function validates two claims from
the EasyAuth principal before executing any Graph call:

| Claim | Check | Failure code |
|---|---|---|
| `tid` | Must match `TENANT_ID` env var | `AUTH_TENANT_MISMATCH` |
| `aud` | Must match `ALLOWED_AUDIENCE` env var | `AUTH_AUDIENCE_MISMATCH` |

This is defence-in-depth on top of EasyAuth's own issuer/audience validation.

## Security design

| Control | Detail |
|---|---|
| **No stored credentials** | DefaultAzureCredential resolves to Managed Identity in Azure; Azure CLI or env vars locally. |
| **Auth at the perimeter** | EasyAuth rejects unauthenticated requests before function code runs. |
| **Caller identity from validated header** | OID comes from `X-MS-CLIENT-PRINCIPAL-ID` (set by EasyAuth). Never from the request body or query string. |
| **Defence-in-depth tenant/audience check** | `tid` and `aud` claims re-validated in code even after EasyAuth passes the request. |
| **GUID validation** | Every OID is validated against a strict GUID regex before being embedded in a Graph API URL. |
| **CORS** | Restricted to `CORS_ALLOWED_ORIGIN` (the tenant's SharePoint origin). |
| **Caller OID redacted in logs** | Only the first 8 and last 4 characters are logged. |
| **Rate limiting** | Anonymous callers: 5 req / 60 s per IP (always active). Authenticated callers: disabled by default; when enabled, each endpoint is rate-limited independently per user. |
| **Response data validation** | Graph response fields (e.g. `availability`) are matched against an allowlist regex before being forwarded to the client. |

**Overall risk level: Low.**

## Environment variables

| Variable | Required | Default | Description |
|---|---|---|---|
| `TENANT_ID` | ✅ | — | Entra tenant ID (GUID). |
| `ALLOWED_AUDIENCE` | ✅ | — | Client ID (GUID) of the EasyAuth App Registration. |
| `CORS_ALLOWED_ORIGIN` | ✅ | — | SharePoint tenant origin, e.g. `https://contoso.sharepoint.com`. |
| `SPONSOR_LOOKUP_TIMEOUT_MS` | | `5000` | Timeout in ms for the sponsor Graph call. |
| `BATCH_TIMEOUT_MS` | | `4000` | Timeout in ms for the Graph `$batch` call (profiles + manager). |
| `PRESENCE_TIMEOUT_MS` | | `2500` | Timeout in ms for the presence Graph call. |
| `RATE_LIMIT_ENABLED` | | `false` | Enable per-user rate limiting for authenticated callers. |
| `RATE_LIMIT_MAX_REQUESTS` | | `12` | Max requests per user and endpoint per window when rate limiting is enabled. |
| `RATE_LIMIT_WINDOW_MS` | | `60000` | Sliding window in ms for rate limiting. |
| `MOCK_MODE` | | `false` | Return demo data without Graph credentials. Only accepted when `NODE_ENV` is not `production`. |

See `local.settings.json.example` for a ready-to-use local development template.

## Graph calls per invocation

The function is called **once per page load** to retrieve sponsor data. Each
invocation runs at most three concurrent Graph requests:

1. **Sponsor lookup** — `GET /users/{oid}/sponsors?$select=...&$top=5`
2. **`$batch`** — profile fields + manager (`$expand=manager`) for each
   sponsor, combined into a single batch request.
3. **Presence** — `POST /communications/getPresencesByUserId` (only when
   `Presence.Read.All` is granted).

**Presence refresh polls bypass the function entirely.** After the initial
load, the web part knows which sponsors exist and polls only for presence
updates at adaptive intervals (30 s while a card is open, 2 min when tab is
visible, 5 min when hidden). These polls call Graph directly with the user's
delegated token (`Presence.Read.All`) — no further Function invocations.

Profile photos are also **not** fetched by the function. The web part fetches
them directly from Graph using the user's own delegated token
(`GET /users/{id}/photo/$value`).

## Response contract

```jsonc
{
  "activeSponsors": [
    {
      "id": "<entra-oid>",
      "displayName": "Jane Smith",
      "mail": "jane.smith@contoso.com",
      "jobTitle": "Senior Manager",
      "department": "Engineering",
      "officeLocation": "Building A",
      "businessPhones": ["+49 30 ..."],
      "mobilePhone": "+49 172 ...",
      "presence": "Available",        // omitted when Presence.Read.All not granted
      "presenceActivity": "Available",
      "managerDisplayName": "Bob Jones",
      "managerId": "<entra-oid>",
      "hasTeams": true                // omitted when TeamMember.Read.All not granted
    }
  ],
  "unavailableCount": 1,              // deleted / disabled sponsors
  "guestHasTeamsAccess": true         // omitted when neither TeamMember nor Presence granted
}
```

The response includes `X-Api-Version` with the deployed function version so
the web part can detect version mismatches and surface a warning banner.

## Local development

See [docs/development.md](../docs/development.md#azure-function-development)
for the full guide. Quick start:

```bash
az login                                      # authenticate for Graph
cp local.settings.json.example local.settings.json
# Fill in TENANT_ID, ALLOWED_AUDIENCE, CORS_ALLOWED_ORIGIN
../scripts/dev-function.sh                    # build + start
```

Or manually:

```bash
npm install
npm run build
func start
```

Local development bypasses EasyAuth. Pass `X-Dev-User-OID: <your-oid>` in
requests to simulate an authenticated caller. This header is **only accepted
when `NODE_ENV !== production`**.

The function uses `DefaultAzureCredential` which in local environments resolves
to Azure CLI (`az login`) or a service principal when `AZURE_CLIENT_ID`,
`AZURE_TENANT_ID`, and `AZURE_CLIENT_SECRET` environment variables are set.

## Source

| File | Purpose |
|---|---|
| `src/getGuestSponsors.ts` | Single function entry point — all logic |
| `infra/setup-graph-permissions.ps1` | Assigns Graph app roles to the Managed Identity |
| `infra/main.bicep` | Azure infrastructure (Function App, Storage, MI, EasyAuth, App Registration, Alerts) |
| `local.settings.json.example` | Template for local development settings |

## License

PolyForm Shield License 1.0.0 — see [LICENSE](../LICENSE.md) for details.

Copyright © 2026 [Workoho GmbH](https://workoho.com/?utm_source=gsiw&utm_medium=docs&utm_campaign=repo&utm_content=af-readme)

Author: [Julian Pawlowski](https://github.com/jpawlowski)

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS
OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR
PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
