---
layout: page
lang: en
title: Privacy Policy
meta_title: Privacy Policy for Guest Sponsor Info
permalink: /en/privacy/
description: >-
  Privacy policy for the Guest Sponsor Info SharePoint web part and Guest
  Sponsor API for Microsoft Entra B2B — how data is handled, which Microsoft
  Graph permissions are used, and where data stays.
lead: >-
  Plain-language overview of what stays inside your tenant, which Microsoft
  Graph permissions are used, and why no runtime data is sent to Workoho.
intro_badges:
  - Data stays in your tenant
  - Graph via Azure Function
  - Application Insights optional
intro_actions:
  - label: See telemetry details
    href: /en/telemetry/
    style: btn-secondary
  - label: Read full policy on GitHub
    href: https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/privacy-policy.md
    style: btn-primary
    external: true
---

This policy covers both **Guest Sponsor Info for Microsoft Entra B2B** (the SharePoint web
part) and the **Guest Sponsor API for Microsoft Entra B2B** (the companion Azure Function).
Together they are designed with a **privacy-first architecture**. All data
processing happens within your own Microsoft 365 and Azure tenant boundaries.

If you are looking specifically for Azure deployment attribution and opt-out,
see the [Telemetry page](/en/telemetry/). For the operational rollout steps,
see the [Setup Guide](/en/setup/).

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

For security posture and trust assumptions, see the
[security assessment on GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/security-assessment.md).

For support options around deployment, customization, or operations, see
[Support](/en/support/).
