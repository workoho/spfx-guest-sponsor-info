---
layout: page
lang: en
title: Trust & Security
meta_title: Trust & Security for Guest Sponsor Info
permalink: /en/security/
description: >-
  Trust and security overview for Guest Sponsor Info for Microsoft Entra B2B —
  deployment boundaries, security controls, disclosure process, and what
  enterprise customers can review today.
lead: >-
  A practical overview for IT admins, security reviewers, and procurement
  teams who want to understand where the trust boundary sits and which
  assurance signals are available today.
intro_badges:
  - Customer-controlled Azure deployment
  - No Graph permissions in the SPFx package
  - Public security and privacy documentation
intro_actions:
  - label: Read the full security assessment
    href: https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/security-assessment.md
    style: btn-primary
    external: true
  - label: See support options
    href: /en/support/
    style: btn-secondary
---

This page summarizes the current trust posture of **Guest Sponsor Info for
Microsoft Entra B2B** and its companion **Guest Sponsor API**.

It is intended as a practical entry point for enterprise review. For the full
technical detail, see the linked GitHub documents.

## Why the Deployment Model Matters

The most important trust characteristic of this solution is where privileged
processing happens.

- **Microsoft Graph application permissions stay server-side** in the Azure
  Function, not in the SharePoint Framework package
- **The Azure Function runs in your own Azure subscription** with your own RBAC,
  monitoring, retention, and compliance boundaries
- **The web part itself requests no Microsoft Graph permissions** in SharePoint
- **Workoho does not operate a shared multi-tenant backend** for sponsor lookups

For many enterprise customers, this matters more than a marketing badge: the
customer controls the Azure resources, identity configuration, logging, and
runtime access model.

## Security Controls in Place

The current recommended deployment uses a layered model:

- Azure App Service Authentication blocks unauthenticated requests before the
  function code runs
- Production requests are validated for tenant, audience, and expected calling
  application claims
- Microsoft Graph access uses a system-assigned Managed Identity, not stored
  secrets
- Presence and photo follow-up requests are scoped to the caller's authorized
  sponsor set
- CORS is restricted to the tenant's SharePoint origin
- Logs are redacted and remain in the customer's Azure subscription

For the detailed trust boundaries, residual risks, and hardening guidance, see
the [full security assessment on GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/security-assessment.md).

## Transparency Instead of Black Boxes

Enterprise security reviews move faster when the architecture is easy to
inspect. The following materials are public and versioned:

- [Security assessment on GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/security-assessment.md)
- [Privacy policy on GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/privacy-policy.md)
- [Architecture documentation on GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/architecture.md)
- [Repository security policy on GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/SECURITY.md)

These materials give IT admins and security teams concrete input for an
informed review instead of relying on generic vendor statements.

## Current Assurance Signals

Today, the strongest trust signals include:

- Transparent architecture and security documentation
- Customer-controlled deployment into the customer's Microsoft 365 and Azure
  environment
- Responsible disclosure channels for vulnerabilities
- A clear separation between the client-side web part, the Azure Function, and
  the customer's own Azure control plane
- Public documentation that supports internal security, compliance, and
  procurement review before deployment

Together, these signals make the solution easier to assess, easier to approve,
and easier to operate within enterprise governance processes.

Where relevant, distribution through the Microsoft commercial marketplace can
also make discovery and internal procurement alignment easier.

## What Enterprise Customers Can Review

If you are evaluating the solution for internal use, the most relevant review
questions are usually:

- Which Microsoft Graph permissions are required, and where are they held?
- Which Azure resources exist, and who administers them?
- How are requests authenticated and scoped to the signed-in guest?
- Where do logs and telemetry go?
- Which parts are optional, such as Teams presence or Azure Maps?

These answers are documented publicly and can be reviewed before deployment.

## Reporting Security Issues

Potential vulnerabilities should not be reported as public GitHub issues.

Use the disclosure guidance in the
[repository security policy on GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/SECURITY.md),
including the private advisory path and the dedicated security contact.

## Current Support Options

The free distribution does not include an SLA or guaranteed response time.
Current rollout, deployment, and operations support options are described on
the [Support page](/en/support/).
