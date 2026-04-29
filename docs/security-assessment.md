# Security Assessment

This document summarizes the current security posture, trust assumptions, and
residual risk of the recommended Guest Sponsor Info deployment model:
SharePoint Framework web part plus the companion Azure Function proxy.

It is aligned with the current implementation in the repository, especially the
Azure Function request validation, follow-up presence and photo authorization,
and the documented deployment flow.

## Executive Summary

Overall risk level for the recommended production deployment: Low.

The main reason the risk stays low is that privileged Microsoft Graph access is
kept server-side inside your Azure subscription. Guest users and browsers do
not receive Graph application permissions, stored secrets, or arbitrary read
access to directory data. The Azure Function returns only caller-scoped sponsor
data after Azure App Service Authentication validates the request.

Residual risk is not zero. The main remaining exposure is the Azure Function's
server-side application permissions and any admin changes that weaken EasyAuth,
CORS, SharePoint page access, or Azure RBAC around the deployed resources.

## Security Model

- The SPFx web part has no Microsoft Graph permissions of its own.
- The browser acquires a bearer token only for the EasyAuth App Registration
  behind the Azure Function.
- Azure App Service Authentication rejects unauthenticated requests before
  function code runs.
- In production, the function performs additional claim checks on the validated
  EasyAuth principal:
  - `tid` must match `TENANT_ID`
  - `aud` must match `ALLOWED_AUDIENCE`
  - `appid` or `azp` must match the SharePoint Online Web Client
    Extensibility application
- The function identifies the caller only from EasyAuth headers. It does not
  accept a user ID for sponsor lookup from query parameters or request bodies.
- Microsoft Graph is called only from the Azure Function by using its
  system-assigned Managed Identity.

## Primary Controls

- No stored credentials: Graph access uses Managed Identity, not client
  secrets or certificates.
- Caller-scoped sponsor lookup: the function returns only sponsors of the
  EasyAuth-authenticated caller OID.
- Defense in depth on client authorization: production requests are validated
  not only for tenant and audience, but also for the expected Microsoft-managed
  SharePoint client application.
- Scoped follow-up authorization: presence and photo requests use a short-lived
  HMAC-signed `X-Presence-Token` when available. If the token is missing or
  invalid, the function falls back to a live sponsor lookup and fails safe.
- Unauthorized ID filtering: the presence endpoint silently drops IDs that are
  outside the caller's authorized sponsor set. The photo endpoint denies access
  when the requested `userId` is not authorized.
- CORS restriction: responses are scoped to the tenant's SharePoint origin.
- Redacted logging: caller OIDs and client IPs are redacted in diagnostics.
  Correlation IDs and reason codes support troubleshooting without exposing raw
  tokens or full GUIDs.
- Rate limiting: anonymous callers are always limited per IP. Authenticated
  per-user limits can be enabled per endpoint as an incident-response measure.
- Safe degradation: missing optional Graph permissions do not widen access.
  The function returns empty presence data or skips optional filtering instead
  of leaking additional information.

## Web Part Delivery Surface

How the bundle is delivered does not materially change the runtime trust model.
Whether you install through AppSource, a Tenant App Catalog, or a Site
Collection App Catalog, the sensitive data path still begins only at the Azure
Function boundary.

- The web part bundle is client-side JavaScript and is expected to be
  downloadable by users who can render the page. When guest access requires the
  Office 365 Public CDN, the compiled bundle is intentionally served as a
  public static asset.
- This is acceptable because the bundle contains no secrets, Graph application
  permissions, or tenant data by design.
- This also means the client bundle must never be treated as secret storage.
  Future changes should not place secrets, internal-only endpoints, or private
  tokens in web part properties or front-end code.

## Client-Side Configuration Caveat

The optional Azure Maps subscription key for inline address previews is used
from the browser and is therefore visible to page viewers and browser tooling.

- Treat that key as a client-distributed billing token, not as a confidentiality
  boundary.
- If this is unacceptable in your environment, leave inline map preview
  disabled and rely on the external map link fallback instead.

## Residual Risks And Assumptions

### Managed Identity Permissions

`User.Read.All` is required on the Azure Function's Managed Identity. Optional
permissions such as `Presence.Read.All`, `MailboxSettings.Read`, and
`TeamMember.Read.All` increase feature coverage, but they also widen the data
the function could query if the function code or Azure control plane were
compromised.

This is still preferable to granting guests direct directory roles, but it is
the main residual privilege concentration in the design.

### Production Configuration Must Stay Intact

The documented authorization model assumes all of the following remain true in
production:

- EasyAuth stays enabled on the Function App and bound to the intended App
  Registration.
- `NODE_ENV=production`, `TENANT_ID`, and `ALLOWED_AUDIENCE` are configured
  correctly.
- SharePoint pre-authorization for silent token acquisition remains configured.
- Azure RBAC on the Function App, Storage Account, App Registration, Log
  Analytics workspace, and Application Insights stays restricted to trusted
  admins.

`PRESENCE_TOKEN_SECRET` is recommended because it keeps follow-up presence and
photo authorization stateless and tightly scoped. If it is absent, the design
does not become insecure, but the function must re-check sponsor membership via
Graph for those follow-up calls.

### SharePoint Access Still Matters

This solution does not replace SharePoint access control. Anyone who can load
the guest landing page can load the client bundle. Whether a user can see the
rendered sponsor data still depends on the page's actual SharePoint permission
model and external sharing configuration.

### Logs Stay In Your Subscription

Application Insights and Log Analytics stay in your Azure subscription and are
not accessible to Workoho, but they still contain operational metadata such as
redacted caller identifiers, reason codes, version mismatch warnings, and rate
limit events. Access to those workspaces should therefore be limited and their
retention should follow your internal policy.

### Rate Limiting Is Not DDoS Protection

The built-in throttling is an in-memory, per-instance safeguard. It helps with
accidental hammering, anonymous probing, and targeted abuse response, but it is
not a replacement for broader network-layer protections such as WAF, Front
Door, or upstream rate controls.

## Recommended Hardening Actions

- Deploy through the provided Bicep and installer flow rather than replacing
  EasyAuth with ad-hoc request handling.
- Keep CORS narrowed to the tenant's SharePoint origin.
- Grant only the optional Graph permissions you actually use.
- Restrict Azure RBAC around the Function App, App Registration, and log
  resources to a small set of trusted admins.
- Treat all client-side configuration as public, especially the optional Azure
  Maps key.
- Monitor for repeated `401`, `403`, `429`, `PHOTO_ACCESS_DENIED`, and
  `client-validation-failed` events.
- Reassess the deployment after tenant-wide authentication, App Registration,
  or external-sharing changes.

## Related Documents

- [architecture.md](architecture.md#azure-function) for the detailed runtime
  flow and authorization behavior
- [privacy-policy.md](privacy-policy.md) for data processing and Graph
  permission scope
- [deployment.md](deployment.md) for admin setup responsibilities and required
  roles
- [telemetry.md](telemetry.md) for Customer Usage Attribution and opt-out

## Report A Security Issue

For responsible disclosure of potential vulnerabilities in this solution,
contact [security@workoho.com](mailto:security@workoho.com).

For non-security privacy inquiries, use
[privacy@workoho.com](mailto:privacy@workoho.com).
