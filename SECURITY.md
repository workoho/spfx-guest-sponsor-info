# Security Policy

## Supported Versions

Only the **latest published release** receives security fixes.
Older versions are not patched. Please upgrade before filing a report.

| Version | Supported |
| ------- | --------- |
| Latest  | ✅        |
| Older   | ❌        |

## Reporting a Vulnerability

**Do not open a public GitHub issue for security vulnerabilities.**

Use one of the following channels for responsible disclosure:

1. **GitHub private advisory (preferred)**
   Open a [private security advisory][advisory] directly in this repository.
   GitHub keeps the report confidential until a fix is published.

2. **E-mail**
   Send details to [security@workoho.com](mailto:security@workoho.com).
   Encrypt sensitive reports with the PGP key published on
   [keys.openpgp.org](https://keys.openpgp.org/search?q=security%40workoho.com)
   (if a key is available there).

[advisory]: https://github.com/workoho/spfx-guest-sponsor-info/security/advisories/new

### What to include

Provide enough detail to reproduce and assess the issue:

- A description of the vulnerability and its potential impact
- The affected component (SPFx web part, Azure Function, Bicep template, or
  scripts)
- Steps to reproduce, or a minimal proof-of-concept
- The version or git commit you tested against
- Any suggested mitigations you have already identified

## What to Expect

| Milestone | Target |
| --- | --- |
| Acknowledgement | Within **5 business days** |
| Initial assessment | Within **10 business days** |
| Fix or mitigation published | Severity-dependent (see below) |

Severity is assessed using [CVSS v3.1][cvss]. Typical timelines:

| CVSS rating | Target fix timeline |
| --- | --- |
| Critical (9.0–10.0) | ≤ 7 days |
| High (7.0–8.9) | ≤ 14 days |
| Medium (4.0–6.9) | Next scheduled release |
| Low (0.1–3.9) | Best effort |

You will be notified when a fix is available and may be credited in the
release notes unless you prefer to remain anonymous.

[cvss]: https://www.first.org/cvss/calculator/3.1

## Scope

The following areas are **in scope**:

- **Guest Sponsor API** (`azure-function/src/`) — authentication bypass,
  privilege escalation, data leakage beyond the calling guest's own sponsors
- **SPFx web part** (`src/`) — XSS via unsanitised data, credential exposure
  in the bundle, or insecure communication patterns
- **Azure infrastructure** (`azure-function/infra/`) — overly permissive role
  assignments or Bicep templates that deploy insecure defaults
- **Setup scripts** (`scripts/`, `azure-function/infra/hooks/`) — code
  execution or secrets exposure during deployment

The following are **out of scope**:

- Vulnerabilities in Microsoft 365, SharePoint Online, or Microsoft Entra ID
  itself — report those to [Microsoft MSRC][msrc]
- Missing HTTP security headers served by SharePoint Online or the Azure
  platform — these are outside this project's control
- Rate-limiting in Azure Functions — enforced at the platform level via Azure
  API Management or Function triggers, not in application code
- Social engineering attacks against tenant administrators
- `npm audit` findings in transitive SPFx build-rig dependencies that cannot
  be upgraded without breaking the build (see `README.md` for context)

[msrc]: https://msrc.microsoft.com/create-report

## Security Architecture

For a description of the security model, trust boundaries, and residual risks,
see [docs/security-assessment.md](docs/security-assessment.md).

The high-level data-flow and authentication chain are documented in
[docs/architecture.md](docs/architecture.md#data-paths).

## Privacy

For privacy-related inquiries that are not security vulnerabilities, contact
[privacy@workoho.com](mailto:privacy@workoho.com).
