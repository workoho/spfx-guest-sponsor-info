# Privacy Policy

**Product:** Guest Sponsor Info for SharePoint Online\
**Publisher:** Workoho GmbH\
**Effective date:** 2026-03-25

---

## Overview

This Privacy Policy describes how the **Guest Sponsor Info** SharePoint web
part and its companion Azure Function API ("the Solution") handle data while
running inside your Microsoft 365 and Azure tenant.

**The Solution does not collect, store, or transmit personal data to Workoho
or any third party.** All data remains within your Microsoft 365 and Azure
environment.

---

## Data Processed by the Solution

The Solution processes the following data at runtime, entirely within your
tenant:

| Data | Purpose | Stored? |
|-|-|-|
| Signed-in user's `loginName` | Detect whether the user is a guest (`#EXT#`) | No — checked in memory only |
| Sponsor display name, photo, job title, email, phone, office | Render sponsor cards in the web part | No — fetched per page load, cached in browser memory for the page lifetime only |
| Azure Function request/response logs | Diagnostics | Application Insights in your own subscription — not accessible by Workoho |

### Microsoft Graph

The Solution accesses Microsoft Graph exclusively with these scopes:

- `User.Read` — read the signed-in user's own profile
- `User.ReadBasic.All` — read basic profile fields of the sponsor users

No broader scopes (`User.Read.All`, `Directory.Read.All`, etc.) are used or
requested.

---

## Telemetry & Customer Usage Attribution

The ARM template for the Azure Function includes a
[Customer Usage Attribution (CUA)](https://aka.ms/partnercenter-attribution)
tracking resource. When the template is deployed, Azure creates an empty
nested deployment named `pid-18fb4033-c9f3-41fa-a5db-e3a03b012939` in your
resource group.

Microsoft uses this GUID to forward **aggregated Azure consumption figures**
(compute hours, storage, etc.) for that resource group to Workoho via Partner
Center. **No personal data, tenant IDs, user names, or resource configurations
are shared.** See the
[Data Collection and Telemetry](deployment.md#data-collection-and-telemetry)
section of the deployment guide for details and opt-out instructions.

---

## Third-Party Services

The web part requests sponsor profile photos directly from Microsoft Graph CDN
endpoints. No third-party analytics, tracking scripts, or advertising services
are used.

---

## Your Rights

All personal data visible in the web part (sponsor profiles) is mastered in
your Microsoft Entra ID tenant. To exercise data-subject rights (access,
rectification, erasure), contact your tenant's Microsoft 365 administrator or
refer to Microsoft's privacy documentation at
[https://www.microsoft.com/privacy](https://www.microsoft.com/privacy).

---

## Contact

For privacy-related questions about this Solution:

**Workoho GmbH**\
<https://workoho.com>

For questions about Microsoft's data processing, refer to the
[Microsoft Privacy Statement](https://privacy.microsoft.com/privacystatement).
