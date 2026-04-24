# Privacy Policy

**Product:** Guest Sponsor Info for Microsoft Entra B2B\
**Publisher:** Workoho GmbH\
**Effective date:** 2026-03-25

---

## Overview

This Privacy Policy describes how the **Guest Sponsor Info for Microsoft Entra B2B**
SharePoint web part and its companion **Guest Sponsor API for Microsoft Entra B2B** (Azure Function, "the Solution")
handle personal data while running inside your Microsoft 365 and Azure tenant.

**The Solution does not collect, store, or transmit personal data to Workoho
or any third party.** All data remains within your Microsoft 365 and Azure
environment.

### Who is affected by this policy?

The Solution involves three distinct parties:

- **The organisation** — the company or institution that deploys the Solution
  in its Microsoft 365 and Azure tenant ("your organisation" or "the tenant
  admin").
- **Member employees of your organisation** — whose profiles are displayed
  as sponsor cards inside the web part.
- **Guest users** — external individuals (from another company or as private
  persons) who sign in to your Microsoft 365 tenant under a Microsoft Entra
  External Identity (guest) account. Guest accounts are identified by the
  `#EXT#` marker in their User Principal Name. The web part is visible
  exclusively to those guest users.

From a data-protection perspective, the guest user is in a special position:
they belong to a third-party organisation and have no direct contractual
relationship with Workoho. Their personal data (UPN, Entra object ID) is
processed solely to render the sponsor contact cards that your organisation
has assigned to them. Workoho has no access to this data. The organisation
that deploys the Solution acts as the data controller; Workoho provides the
software tool only.

---

## Data Processed by the Solution

All personal data is processed at runtime exclusively within your tenant. No
data is stored beyond the current browser session or function invocation.

### Web Part (runs in the guest user's browser)

| Data | Source | Purpose | Stored? |
|-|-|-|-|
| Guest user's UPN / `loginName` | SharePoint page context | Detect the `#EXT#` marker to identify guest accounts | No — evaluated in memory only, never transmitted |
| Guest user's Entra object ID (OID) | Entra ID token (via MSAL) | Authenticate requests to the Guest Sponsor API for Microsoft Entra B2B | No — present only in the short-lived Bearer token |
| Sponsor display name, given name, surname | Guest Sponsor API for Microsoft Entra B2B | Render the sponsor name on the card | No — held in browser memory for the page lifetime |
| Sponsor job title, department | Guest Sponsor API for Microsoft Entra B2B | Display role context on the card | No |
| Sponsor profile photo | Microsoft Graph CDN | Show a visual identifier on the card; initials fallback when absent | No — decoded in browser memory |
| Sponsor email address | Guest Sponsor API for Microsoft Entra B2B | Render mailto link on the card | No |
| Sponsor phone numbers (business, mobile) | Guest Sponsor API for Microsoft Entra B2B | Render click-to-call links on the card | No |
| Sponsor office location, city, country, address | Guest Sponsor API for Microsoft Entra B2B | Render address and map hint on the card | No |
| Sponsor Teams presence (availability, activity) | Guest Sponsor API for Microsoft Entra B2B | Show presence indicator on the card | No — polled periodically, held in browser memory |
| Sponsor's manager: display name, job title, department, photo | Guest Sponsor API for Microsoft Entra B2B | Render manager context on the card | No |
| Guest's own Teams provisioning status | Azure Function (via Microsoft Graph) | Enable/disable Teams chat and call buttons | No |

### Guest Sponsor API for Microsoft Entra B2B (Azure Function, runs in your Azure subscription)

The Azure Function processes personal data only for the duration of a single
HTTP request:

| Data | Source | Purpose | Stored? |
|-|-|-|-|
| Guest user's Entra OID (from EasyAuth header `X-MS-CLIENT-PRINCIPAL-ID`) | Azure App Service EasyAuth | Identify the caller; look up their sponsors via Graph | No — discarded after the request |
| Guest user's tenant ID and token audience (from EasyAuth claims) | Azure App Service EasyAuth | Validate that the caller belongs to the correct tenant | No |
| Sponsor profile fields (same set as above) | Microsoft Graph application call | Construct the JSON response | No — not persisted |
| Sponsor account status (`accountEnabled`, `isResourceAccount`, `assignedPlans`) | Microsoft Graph | Filter out disabled and resource accounts (Teams Room devices, Common Area Phones, etc.) | No |
| Sponsor mailbox settings (`mailboxSettings.userPurpose`) | Microsoft Graph | Filter out shared, room, and equipment mailboxes (requires `MailboxSettings.Read`) | No |
| Guest's joined Teams (`joinedTeams`) | Microsoft Graph | Determine Teams provisioning status | No |
| Guest's own Teams presence | Microsoft Graph | Used as fallback Teams provisioning signal | No |
| Redacted IP address (last octet masked for IPv4 / last 64 bits for IPv6) | HTTP request | Anonymous rate-limiting; partial security logging | Application Insights in **your** subscription — not accessible by Workoho |
| Redacted caller OID (first 8 and last 4 hex chars only) | Derived from EasyAuth OID | Structured logging / audit traces | Application Insights in **your** subscription |
| Web part version (`X-Client-Version` request header) | HTTP request header | Detect version mismatches; log update-available warnings | Application Insights in **your** subscription |

**Application Insights** receives structured traces, warnings, and error events
from the Function App. This data is stored in a Log Analytics workspace inside
**your own Azure subscription**. Workoho has no access to it.

---

## Microsoft Graph Permissions

The web part has **no Microsoft Graph permissions of its own**. It authenticates
exclusively to the Azure Function (using the Function's App Registration as the
token audience). All Graph calls are made server-side by the Azure Function
through its Managed Identity.

### Azure Function — Application Permissions (acting as its own Managed Identity)

These permissions are granted to the Function App's system-assigned Managed
Identity by running `infra/setup-graph-permissions.ps1`. They allow the
function to query Microsoft Graph server-side, independent of the guest user's
own consent.

| Permission | Purpose | Required? |
|-|-|-|
| `User.Read.All` | Read sponsor profiles, sponsor list (`/users/{id}/sponsors`), and `accountEnabled` status to filter disabled accounts | **Required** — assigned by the setup script. The function also accepts `User.ReadBasic.All` or `Directory.Read.All` as minimal alternatives (lose `accountEnabled` filtering); these require manual assignment. |
| `Presence.Read.All` | Read real-time Teams presence status for sponsors and detect guest Teams provisioning | Optional — presence indicators disabled without it |
| `MailboxSettings.Read` | Read `mailboxSettings.userPurpose` to filter shared mailboxes, room accounts, and equipment accounts from the sponsor list | Optional — filter is skipped without it |
| `TeamMember.Read.All` | Read the guest's joined Teams to determine whether their Teams account has been provisioned | Optional — Teams chat/call buttons default to enabled without it |

By default, the `setup-graph-permissions.ps1` script grants **all four
permissions** (one required, three optional). A tenant administrator may choose
to omit the optional permissions; doing so reduces functionality as described
in the table above.

---

## Telemetry & Customer Usage Attribution

The Bicep template for the Azure Function includes a
[Customer Usage Attribution (CUA)](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/azure-partner-customer-usage-attribution)
tracking resource. When the template is deployed, Azure creates an empty
nested deployment named `pid-18fb4033-c9f3-41fa-a5db-e3a03b012939` in your
resource group.

Microsoft uses this GUID to forward **aggregated Azure consumption figures**
(compute hours, storage, etc.) for that resource group to Workoho via Partner
Center. **No personal data, tenant IDs, user names, or resource configurations
are shared.** See [telemetry.md](telemetry.md) for details and opt-out
instructions.

### GitHub Release Check

The Azure Function checks the GitHub Releases API once every six hours (and
on every cold start) to detect whether a newer version of the Solution is
available. This outbound HTTPS request is made **from the Azure Function
runtime inside your Azure subscription** and contains:

- The current Function version in the `User-Agent` header
  (e.g. `guest-sponsor-info-function/1.2.3`)

The result is cached in the Function's process memory for up to six hours.
When the SPFx web part is opened in edit mode and an Azure Function URL is
configured, the web part fetches this cached result from the Function's
`/api/getLatestRelease` endpoint — **not directly from GitHub**. No GitHub
API calls are made from the browser.

If no Azure Function is configured for the web part, no GitHub release check
is performed at all and the update notification in the property pane remains
hidden.

No personal data, tenant IDs, or user information are transmitted to GitHub.
The check is a standard read-only GitHub public API call and is subject to
GitHub's
[Privacy Statement](https://docs.github.com/site-policy/privacy-policies/github-general-privacy-statement).

---

## Microsoft AppSource / SharePoint Store

The Solution is submitted to the Microsoft AppSource marketplace and the
SharePoint Store. By installing it, your organisation accepts these policies
in addition to the
[Microsoft Marketplace Terms](https://learn.microsoft.com/en-us/legal/marketplace/marketplace-terms).

The following commitments apply to the Solution as a marketplace offering:

- **No hidden data collection.** The Solution does not phone home, embed
  analytics SDKs, or transmit telemetry to any Workoho-controlled endpoint.
- **Least-privilege permissions.** Only the permissions described in this
  policy are declared or requested. No permissions are silently elevated.
- **No token forwarding.** Bearer tokens issued by Entra ID are validated
  server-side by Azure App Service EasyAuth. The function code never parses
  or forwards raw token strings.
- **No stored credentials.** The Azure Function uses a system-assigned Managed
  Identity. No client secrets, certificates, or passwords are stored in code or
  configuration.
- **GDPR/data residency.** All personal data processed by the Solution stays
  within the tenant's Microsoft 365 and Azure regions. Workoho has no ability
  to access or export it.

---

## Third-Party Services

| Service | Used by | Data sent | Link |
|-|-|-|-|
| Microsoft Graph API | Web part, Azure Function | See permission tables above | [Privacy Statement](https://privacy.microsoft.com/privacystatement) |
| Microsoft Graph CDN | Web part browser | Sponsor/manager profile photo requests | [Privacy Statement](https://privacy.microsoft.com/privacystatement) |
| GitHub Releases API | Azure Function (timer trigger, every 6 h) | `User-Agent` with function version — no browser calls | [Privacy Statement](https://docs.github.com/site-policy/privacy-policies/github-general-privacy-statement) |
| Azure Application Insights | Azure Function | Structured traces (redacted IDs only) | Stored in **your** subscription |

No third-party analytics, advertising, or tracking services are used.

### UTM Parameters on Admin and Setup Links

Links displayed in the **setup wizard** (Welcome Dialog) and **property pane**
(the SharePoint admin configuration area of the web part) include static UTM
parameters (`utm_source`, `utm_medium`, `utm_campaign`, `utm_content`). These
parameters contain only campaign metadata — for example, which area of the
web part the admin clicked from — and do **not** contain tenant IDs, user IDs,
email addresses, or any other personal identifiers.

These UTM parameters are **not** added to links opened by end users (for
example, links to map providers such as Bing, Google Maps, or OpenStreetMap
are passed without any UTM parameters).

---

## Data Subject Rights

### Sponsor employees (member users of your organisation)

Profile data (name, photo, job title, etc.) is mastered in your organisation's
Microsoft Entra ID directory. To exercise data-subject rights (access,
rectification, erasure, objection), contact your organisation's Microsoft 365
administrator or refer to Microsoft's privacy documentation.

### Guest users

Guest accounts are managed in your organisation's Entra ID tenant as External
Identities. The guest user's home organisation retains control over the
personal data held in their home tenant (e.g. display name, email). Your
organisation controls the guest object in your tenant (e.g. the sponsor
assignment). To exercise data-subject rights, the guest should contact:

1. **Your organisation's data protection officer or Microsoft 365 admin** — for
   the Entra guest object and sponsor assignments stored in your tenant.
2. **Their own (home) organisation** — for personal data mastered there.
3. **Microsoft** — for data processing by Microsoft 365 and Azure services.
   Refer to the [Microsoft Privacy Statement](https://privacy.microsoft.com/privacystatement).

---

## Contact

For privacy-related questions about this Solution:

**Workoho GmbH**\
[privacy@workoho.com](mailto:privacy@workoho.com)\
[workoho.com](https://workoho.com/?utm_source=gsiw&utm_medium=docs&utm_campaign=repo&utm_content=privacy-policy)

For responsible disclosure of security vulnerabilities related to this
Solution, contact [security@workoho.com](mailto:security@workoho.com).

For questions about Microsoft's data processing, refer to the
[Microsoft Privacy Statement](https://privacy.microsoft.com/privacystatement).
