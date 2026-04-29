---
layout: doc
lang: en
title: Telemetry
meta_title: Telemetry & Usage Attribution
permalink: /en/telemetry/
description: >-
  Telemetry and Customer Usage Attribution for the Guest Sponsor Info Azure
  deployment — what Azure usage data is shared, what is not collected, and how
  to opt out.
lead: >-
  A plain-language summary of the telemetry built into the Azure
  deployment. No personal data is ever collected or shared.
github_doc: telemetry.md
---

## What happens when you deploy

When you deploy the Azure Function using the included Bicep template and azd,
a small tracking marker is added to your resource group:

```text
pid-18fb4033-c9f3-41fa-a5db-e3a03b012939
```

This is an empty, harmless nested deployment. Microsoft uses this GUID to
forward **aggregated Azure consumption figures** (compute hours, storage
transactions, and similar billing signals) for that resource group to
[Workoho](https://workoho.com) via Partner Center.

This mechanism is called
[Customer Usage Attribution (CUA)](https://aka.ms/partnercenter-attribution)
and helps Workoho understand how the solution is used and justify continued
development.

If you are looking for Microsoft Graph permissions, runtime data handling, or
tenant-boundary details instead, see the [Privacy Policy](/en/privacy/).

## What is NOT collected

- **No personal data** — no user names, email addresses, or tenant IDs
- **No resource names**, configurations, or secrets
- **No data leaves your Azure subscription** — Microsoft only shares summary
  consumption figures with Workoho using existing billing data

For information about personal data processed within your tenant at runtime,
see the [Privacy Policy](/en/privacy/).

## What you will see in Azure Portal

In **Resource Group → Deployments** you will see a deployment named
`pid-18fb4033-c9f3-41fa-a5db-e3a03b012939`. It is an empty nested deployment.
Deleting it has no effect on running resources but stops future attribution
for that resource group.

## How to opt out

Pass `-EnableTelemetry $false` to the deployment wizard:

```powershell
& ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.ps1').Content)) -EnableTelemetry $false
```

Or, when running `deploy-azure.ps1` directly from an extracted infra ZIP:

```powershell
./deploy-azure.ps1 -EnableTelemetry $false
```

If you are still deploying the solution, the [Setup Guide](/en/setup/)
describes where this fits into the overall rollout.

## Contact

For telemetry and privacy questions about this solution, contact
[privacy@workoho.com](mailto:privacy@workoho.com).

For responsible disclosure of security vulnerabilities, contact
[security@workoho.com](mailto:security@workoho.com).

For deployment help or managed operations, see [Support](/en/support/).
