# Architecture and Design Decisions

This document explains the technical decisions made in the Guest Sponsor Info web part.
For installation and build instructions see the [README](../README.md).

## Guest Detection

**Primary signal:** `pageContext.user.isExternalGuestUser` — a boolean property on the SPFx
page context that is set directly from the Entra authentication token and is available
synchronously without any Graph call. This is the authoritative indicator for B2B guests.

**Fallback signal:** The `#EXT#` marker in `pageContext.user.loginName`. Entra guest UPNs
follow the form `user_externaldomain.com#EXT#@tenant.onmicrosoft.com`. This heuristic covers
edge cases where `isExternalGuestUser` might not be populated (e.g. during first-load before
the SPFx context is fully hydrated).

The combined check is: `isGuest = isExternalGuestUser || loginName.includes('#EXT#')`.

**Known limitation:** When a guest user visits a SharePoint site for the first time, the call
to `SP.UserProfile.ReadCacheOrCreate` may return HTTP 500 because the guest profile has not
been created yet. This can cause `pageContext.user.loginName` to be empty. The
`isExternalGuestUser` flag is populated from the Entra token and is unaffected by this
user-profile failure, which is why it is used as the primary signal.

**If the `#EXT#` marker is absent:** In unusual tenant configurations where a guest UPN has
been reset or renamed, the Graph `userType` property (`GET /v1.0/me?$select=userType`)
is the ground truth. This would require an async Graph call that is out of scope for the
current implementation; `isExternalGuestUser` covers this gap for all standard Entra B2B
configurations.

## Sponsor Retrieval

**Endpoint used:** `GET /v1.0/me/sponsors`

**Assumption:** The `/me/sponsors` endpoint is available and returns the list of assigned
sponsors in the resource tenant. This endpoint was added to Microsoft Graph in 2023 and is
generally available in v1.0.

**Properties selected:** `id`, `displayName`, `mail`, `jobTitle`, `department`, `officeLocation`, `businessPhones`, `mobilePhone`

**Risk:** The `/me/sponsors` relationship may be empty if no sponsors have been assigned in
Entra. The web part handles this gracefully and shows a "no sponsors" message.

## Profile Photos

**Endpoint used:** `GET /v1.0/users/{id}/photo/$value`

**Assumption:** The signed-in guest user has read access to the profile photos of their
sponsors. The `User.ReadBasic.All` delegated permission grants this access.

**Implementation detail:** Photos are fetched as `ArrayBuffer` and converted to base64 data
URLs. This avoids `Blob` URL leaks that would require explicit cleanup.

**Risk:** Profile photos may be restricted by tenant policies. A failed photo request is
silently ignored and the initials-based fallback is shown instead.

## Required Microsoft Graph Permissions

The following delegated permissions must be granted by a tenant administrator in the
SharePoint Admin Center → API access panel after the `.sppkg` is uploaded.

| Permission | Why it is needed | Can it be reduced? |
|---|---|---|
| `User.Read` | Read the signed-in user's own `/me/sponsors` relationship. | No – this is the smallest delegated scope in Graph. |
| `User.ReadBasic.All` | Read whether each sponsor's directory object still exists (HTTP 404 detection) and load their profile photos via `/users/{id}/photo/$value`. | No – there is no narrower scope that covers reading *other* users' objects. `ReadBasic` exposes only: displayName, first/last name, mail, and photo. It does **not** grant access to sensitive properties like `accountEnabled`, licences, group memberships, etc. |

### Why `*.All` does not mean "full access"

The `All` suffix means the permission applies to *all users* in the tenant, not that it
grants access to *all data* about a user. `User.ReadBasic.All` is deliberately limited
to a very small set of identity attributes. Microsoft provides no narrower alternative
for reading another user's directory objects.

### Why `User.Read.All` is intentionally absent

Reading the `accountEnabled` flag on other users requires `User.Read.All` (or
`Directory.Read.All`), both of which expose far more tenant data than needed.
We therefore do **not** rely on `accountEnabled`.

Instead, we detect that a sponsor is no longer available by making a lightweight probe
request (`GET /users/{id}?$select=id`) for each sponsor:

- **HTTP 200** → sponsor's directory object still exists. The web part shows the card.
  (If the account is merely *disabled* rather than *deleted*, it will still appear –
  see limitation below.)
- **HTTP 404** → sponsor's directory object is gone (hard-deleted or past the 30-day
  soft-delete window). The web part counts this sponsor as unavailable.
- **Any other error** → treated as "still exists" so that a transient Graph outage does
  not incorrectly hide a sponsor card.

### Known limitation: disabled-but-not-deleted accounts

A sponsor whose Entra account has been *disabled* (but not yet deleted) will still show
in the web part, because confirming `accountEnabled === false` on another user's object
is blocked by the least-privilege boundary above. The account will disappear automatically
once it is hard-deleted (typically 30 days after disabling, or immediately if
force-deleted by an admin).

If this limitation becomes critical, a future iteration can:

- Upgrade to `User.Read.All` and re-introduce the `accountEnabled` check, **or**
- Use an admin-side automation to remove the sponsor assignment from the guest user
  object at the time of account deactivation.

## App Catalog Access for Guest Users

All JavaScript and CSS assets are bundled *inside* the `.sppkg` file
(`includeClientSideAssets: true`). SharePoint re-hosts those assets on the **App Catalog
site collection** and serves them from there at runtime.

When a guest user loads a page that contains this web part, the browser requests the bundle
from the App Catalog origin. Because guest users have no default access to the App Catalog,
the request fails with HTTP 403 and the web part never initialises.

**Required one-time tenant configuration (admin):**

1. Verify that `ShowAllUsersClaim` is `$true` (it is the default on all modern tenants —
   only needs attention on tenants provisioned before ~2018 or if it was explicitly disabled).
2. **Enable external sharing on the App Catalog site** (SharePoint Admin Center → Sites →
   Active sites → App Catalog → Policies → External sharing → *Existing guests*).
   Without this, guests receive HTTP 403 even when Read permission has been granted.
3. Grant *Read* permission on the App Catalog site to **All Users (membership)** or to a
   specific security group that contains the guest accounts.
   Note: *Everyone except external users* explicitly excludes B2B guests and will not work.

This is entirely a SharePoint/Entra configuration concern and requires no code changes.
See [README – Guest Access Requirements](../README.md#guest-access-requirements) for the
full step-by-step procedure.

**Why not `isDomainIsolated: true`?**
Domain-isolated web parts use a separate Microsoft Entra application and an isolated frame.
That model adds complexity and has known limitations with guest-user token acquisition;
`isDomainIsolated: false` (the default) is the recommended approach for guest-facing
solutions.

## Edit Mode Behaviour

In edit mode, the web part always shows a lightweight text placeholder. No Graph calls
are made in edit mode. This keeps the authoring experience fast and avoids permission
prompts for page editors who are not guest users.

## Display Mode Rendering

| User type | Edit mode | View mode |
|---|---|---|
| Guest user | Placeholder text | Full sponsor card grid |
| Non-guest user | Placeholder text (different message) | Web part hidden (`null`) |

## Profile Photo Colour Palette

Initials avatars use a deterministic colour derived from the sponsor's display name.
The colour set matches the Fluent UI persona colours used by SharePoint people experiences.

## Hover / Focus Interaction

The contact details overlay currently appears to the **right** of the card using
`position: absolute`. For cards near the right viewport edge, the overlay may be clipped.
A future iteration could use the Fluent UI `Callout` component, which handles viewport
boundary flipping automatically.

## Development Testing

Testing this web part end-to-end is more involved than typical SPFx components because the
core feature — showing a guest user's sponsors — requires the signed-in user to actually be
a guest in the target tenant, with sponsors assigned.

### Hosted workbench — as a regular member

Set `SPFX_TENANT` in your `.env` file and run `./scripts/start.sh`. SPFx serves the local
bundle from `https://localhost:4321` and embeds it in the real SharePoint page. You will
have a real `pageContext`, but because `#EXT#` is absent from your login name you will be
treated as a non-guest → the web part renders nothing in view mode. Useful for verifying
the non-guest code path.

### Hosted workbench — as a guest (full integration test)

This is the only complete end-to-end test. Prerequisites:

1. **A second Microsoft 365 tenant** (e.g. a partner or dedicated test tenant) where your
   account has been invited as a guest user.
2. **Sponsors must be explicitly assigned** to your guest account in that tenant:
   Microsoft Entra admin center → Users → your guest account → Sponsors → assign one or more
   internal users as sponsors.
3. **API permissions must be consented** in that tenant's SharePoint Admin Center → API
   access panel (`User.Read` and `User.ReadBasic.All`).
4. **The `.sppkg` must be deployed** in that tenant's App Catalog, or the tenant must be
   configured to load scripts from `localhost:4321` during local dev mode.

Set `SPFX_TENANT=<partner-tenant>.sharepoint.com` in your `.env`, run `./scripts/start.sh`,
and open the hosted workbench URL in a browser session where you are signed in as the guest
account.

### Unit tests (Jest)

`npm test` covers guest detection logic and Graph service calls with mocked API responses.
This is the practical substitute for integration testing when a guest-enabled second tenant
is not available. Run the tests after every change to `SponsorService.ts` or the main
component.

### Demo mode

The web part includes a **Demo mode** toggle in the property pane (edit the web part →
property pane → "Demo mode (mock data)"). When enabled:

- The user is treated as a guest regardless of the actual login name.
- Two fictitious sponsor records from `MockSponsorService.ts` are rendered without any
  Graph calls.
- Useful for visual review of sponsor card layout and styles on the hosted workbench
  without needing a genuine guest account with sponsors assigned.

Demo mode is intended for development and visual review only. It is not a security
boundary; disable it before deploying to production. The property is stored in the web
part instance properties and defaults to `false`.
