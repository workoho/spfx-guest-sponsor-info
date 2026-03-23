# Features in Depth

High-level overview: [README](../README.md)
For deployment details see [deployment.md](deployment.md).
For development setup see [development.md](development.md).
For design decisions see [architecture.md](architecture.md).

---

## The problem this solves

When employees from partner organisations land on your SharePoint tenant as
B2B guests, they often have no idea who their internal sponsor is or how to
reach them. There is no out-of-the-box SharePoint component that surfaces this
information, and even if you knew the Graph API existed, calling `/me/sponsors`
from a delegated context requires the guest to hold an Entra directory role —
completely impractical to assign at scale.

The *Guest Sponsor Info* web part closes that gap: it shows every guest exactly
who to call, chat with, or email, in a polished UI that feels native to
SharePoint.

---

## Sponsor cards — live, not static

Each assigned sponsor is rendered as a Fluent UI persona card with their photo,
name, and job title — modelled after the Microsoft Teams profile card and
SharePoint's built-in People web part, so the result feels native to the
platform rather than like a custom add-on.

Profile photos are fetched live from Microsoft Graph using progressive loading:
cards appear immediately with name and title, and photos fill in behind them
without causing a layout shift.
When a photo is unavailable, a **deterministic colour avatar** is generated from
the sponsor's name — so the display always looks intentional rather than broken.

Two layout modes are available, switchable per web part instance:

- **Full** — 136 px tile grid, matches the SharePoint People web part style
- **Compact** — horizontal name list, good for narrow columns or sidebar zones
- **Auto** — switches between Full and Compact based on available column width

---

## Rich contact details on hover or focus

Hovering or tabbing to a card opens a contact panel with everything a guest
needs to actually reach their sponsor:

- **Email**, **work phone**, and **mobile** — each with a one-click copy button
  that shows a checkmark confirmation for 1.5 s
- **Office location** and **full address** (street, city, state, ZIP, country),
  configurable field by field so admins can expose exactly what is relevant
- **Inline map preview** via Azure Maps (when a subscription key is configured),
  or an external map link to Bing, Google, Apple Maps, OpenStreetMap, or HERE —
  whichever fits the organisation
- **Teams Chat** and **Audio Call** buttons that open directly in the native
  Teams client, built with the correct guest tenant context so they work even
  when the guest belongs to a different home tenant

The panel also shows the sponsor's **manager** (photo, name, title, department)
so the guest understands the reporting chain at a glance.

Every contact field shown in the panel is individually toggleable in the web
part property pane, giving page authors full control over what is displayed:
phone numbers, full address broken down by street, city, state, ZIP, and
country, map preview, manager section, presence status indicators, and
sponsor/manager photos can all be turned on or off independently.

---

## Real-time presence — without extra permission wrangling

The web part polls Microsoft Teams presence for all displayed sponsors and
renders it as a live availability indicator using the same colour semantics as
Teams itself:

| State | Colour |
|---|---|
| Available | Green |
| Away / Be Right Back | Amber |
| Busy / In a call / In a meeting / Presenting | Red |
| Do Not Disturb | Dark red |
| Focusing | Teams purple (#6264A7) |
| Out of Office | Magenta |
| Offline | Grey |

Activity tokens (`InAMeeting`, `InACall`, `Presenting`, …) take priority over
the base availability token, matching Microsoft's profile card display behaviour.
The "Out of Office" state is rendered as a suffix modifier (e.g.
*"Available, out of office"*) when the guest is also reachable.

Polling adapts to user behaviour to minimise unnecessary Graph traffic:

- **30 seconds** — while a sponsor card is actively open (hovered or focused)
- **2 minutes** — while the browser tab is visible but no card is active
- **5 minutes** — while the browser tab is hidden

Presence is entirely optional. If `Presence.Read.All` has not been consented,
the presence indicator is omitted and nothing breaks — no error, no empty space.

---

## Teams integration details

The Chat and Call buttons in the contact panel generate Teams deep links with
the `tenantId` query parameter set to the host tenant's Entra ID, so clicking
them opens the correct guest context even when the guest is signed into their
own home tenant in the Teams client.

### "Teams not provisioned yet" state

A guest account can exist in Entra before it has been added to any Team. In
this state, the guest cannot use Teams features in the host tenant. The web part
detects this condition (via the Azure Function proxy) and reacts gracefully:

- The Chat and Call buttons are **disabled** with explanatory tooltips
- An **informational banner** appears below the sponsor grid, guiding the guest
  to ask their sponsor to add them to a Team

No error message, no broken buttons — just a clear, actionable explanation.

---

## The Azure Function proxy — the part that makes it actually work

The central architectural challenge is that the Graph `/me/sponsors` relationship
requires the calling user to hold an Entra directory role. Assigning roles to
guests at scale is not viable:

- Role-assignable groups have no dynamic membership
- Automating role assignment requires `RoleManagement.ReadWrite.Directory`
- The sponsor-read permission is not self-scoped — a potential GDPR concern
  because directory roles typically grant broader read access

The included **Azure Function proxy** sidesteps all of this. The function:

1. Receives the guest's request authenticated by **EasyAuth** (no secret in the
   client, no custom auth header handling)
2. Reads the caller's Object ID exclusively from the EasyAuth-validated principal
   claims — callers cannot query other users' sponsors by manipulating a parameter
3. Calls Graph with its own **Managed Identity** (application permissions, no
   stored secrets, RBAC-based key access)
4. Returns only the sponsor data for the authenticated caller

Additional security safeguards:

- **Tenant ID and audience** validated in production as defence-in-depth on top
  of EasyAuth's issuer/audience validation
- **CORS** locked to the host tenant's SharePoint origin
- **Rate limiting** — two tiers: IP-based for unauthenticated callers (always
  on), optional per-user limiting via env vars for incident response
- **Structured error logging** with reference IDs so support tickets contain
  enough context to diagnose failures without exposing sensitive GUIDs

> Full function internals: [architecture.md](architecture.md#azure-function)
> Deployment steps: [deployment.md](deployment.md)

---

## Guest-only rendering

The web part renders **nothing** for member accounts — no empty shell, no
"nothing to see here" message. This is intentional: member users do not have
sponsors, and a blank space is cleaner than a visible but empty widget.

### Preview mode for page editors

Page authors see **realistic sample cards** in edit mode — no Graph calls, no
live photos, no network traffic. The mock cards show a plausible sponsor name,
title, department, and initials avatar, giving editors an accurate picture of
how the web part will look for real guests without requiring anyone to sign in
as a guest account. When the page is published and viewed by an actual guest,
the real sponsor data takes over seamlessly.

Guest detection uses two signals in combination:

- `isExternalGuestUser` from `pageContext.user` (primary — derived from the
  Entra token, synchronous, always available)
- `#EXT#` in the login name UPN (fallback — covers the edge case where the
  SharePoint user profile is not yet provisioned on a guest's very first visit)

---

## Resilience

The web part is designed to degrade gracefully rather than show errors unless
something is genuinely unrecoverable:

- **Exponential backoff** — transient errors retry up to 3 times with increasing
  delays (3 s → 9 s → 27 s, capped at 30 s) before surfacing an error message
- **Progressive photo loading** — sponsor names and titles appear in the first
  render; photos fill in asynchronously without blocking the initial display
- **Active-sponsor filtering** — three categories of non-reachable accounts are
  silently excluded before rendering:
  - **Disabled accounts** — sponsors whose Entra account has been deactivated
    (e.g. departed employees still in the soft-delete grace period)
  - **Shared and room mailboxes** — system accounts that are technically valid
    directory objects but have no real person behind them
  - **Deleted sponsors** — accounts whose directory object no longer exists
    (hard-deleted or past the soft-delete period)

  When all assigned sponsors fall into one of these categories, the guest
  receives a clear informational notice rather than an empty page
- **Non-blocking SPFx initialisation** — Graph and AAD client acquisition runs
  in the background after `onInit()` resolves, so the host page layout is never
  held up waiting for this web part to finish initialising

---

## Privacy-first permissions

| Data path | Permissions used |
|---|---|
| Azure Function (recommended) | `User.Read.All` · `Presence.Read.All` *(optional)* — on the Managed Identity, guests never hold these |
| Direct Graph (legacy fallback) | `User.Read` · `User.ReadBasic.All` — avoids `User.Read.All` entirely |

On the delegated (direct Graph) path, `User.ReadBasic.All` is chosen over the
broader `User.Read.All` deliberately. `ReadBasic` exposes only:
display name, given name, surname, mail, and profile photo. It does **not**
expose `accountEnabled` or other sensitive directory properties.

---

## Multilingual — 14 locales

Every user-facing string ships in 14 languages:

English · German · French · Spanish · Italian · Danish · Finnish · Norwegian ·
Swedish · Japanese · Chinese (Simplified) · Portuguese (Brazil) · Polish · Dutch

### Informal salutation mode

Languages with a T–V distinction (German `du/Sie`, French `tu/vous`, Spanish
`tú/usted`, Italian `tu/Lei`, Dutch `je/u`) support an optional **informal
salutation mode**, configurable per web part instance in the property pane.

When enabled, all user-facing messages and banners that address the guest
directly in full sentences (error messages, Teams access notice, etc.) switch
to the informal variant. Card labels and section headings are not affected.

---

## Theme-aware rendering

The web part inherits the active SharePoint site theme automatically via CSS
custom properties — no extra configuration required. All interactive elements
(hover states, focus rings, callout borders) adapt to both light and dark themes.

---

## Automating sponsor assignments

This web part handles the *display* side — it shows guests who their sponsor is.
If you also want to **automate who gets assigned as a sponsor** and keep those
assignments current as people join, move teams, or leave, that is a separate
lifecycle problem.

[EasyLife 365 Collaboration](https://www.easylife365.cloud) is purpose-built for
exactly that: it manages the full lifecycle of Microsoft 365 collaboration
workspaces — Teams, SharePoint sites, Viva Engage communities, and more —
including guest onboarding workflows and sponsor assignments. The two tools are
designed to complement each other: EasyLife 365 keeps the directory data
current; this web part surfaces it to the guest.

[Workoho](https://www.workoho.com), the author of this web part, is a Platinum
sales and implementation partner of EasyLife 365 and can help with both the
technical rollout and the collaboration governance design around it.
