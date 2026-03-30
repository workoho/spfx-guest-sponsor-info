---
layout: page
lang: en
title: Privacy Policy
permalink: /en/privacy/
description: >-
  Privacy policy for Guest Sponsor Info for Microsoft Entra B2B and the
  Guest Sponsor API for Microsoft Entra B2B — how data is handled, what
  permissions are used, and where data is stored.
---

This policy covers both **Guest Sponsor Info for Microsoft Entra B2B** (the SharePoint web
part) and the **Guest Sponsor API for Microsoft Entra B2B** (the companion Azure Function).
Together they are designed with a **privacy-first architecture**. All data
processing happens within your own Microsoft 365 and Azure tenant boundaries.

## Key Principles

- **No data sent to Workoho or third parties** — the web part and Azure
  Function operate entirely within your tenant.
- **Browser memory only** — the web part holds sponsor data (name, title,
  email, phone, Teams presence) in browser memory during the page session.
  Nothing is persisted to disk or sent elsewhere.
- **Azure Function is stateless** — each request is processed and discarded.
  No sponsor or guest data is stored.
- **Your Application Insights** — if enabled, telemetry goes to your own
  Azure subscription. Workoho has no access.

## Permissions Used

All Microsoft Graph permissions are held exclusively by the
**Guest Sponsor API for Microsoft Entra B2B** — the web part itself has none.

| Scope | Required? | Purpose |
|-------|-----------|---------|
| [`User.Read.All`](https://learn.microsoft.com/en-us/graph/permissions-reference#userreadall) | Required | Read sponsor profiles and filter disabled accounts |
| [`Presence.Read.All`](https://learn.microsoft.com/en-us/graph/permissions-reference#presencereadall) | Optional | Teams presence indicators |
| [`MailboxSettings.Read`](https://learn.microsoft.com/en-us/graph/permissions-reference#mailboxsettingsread) | Optional | Filter shared/room/equipment mailboxes |
| [`TeamMember.Read.All`](https://learn.microsoft.com/en-us/graph/permissions-reference#teammemberreadall) | Optional | Detect guest Teams account provisioning |

## Full Policy

For the complete privacy policy including data subject rights, GitHub release
checks, and Customer Usage Attribution details, see the
[full privacy policy on GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/privacy-policy.md).
