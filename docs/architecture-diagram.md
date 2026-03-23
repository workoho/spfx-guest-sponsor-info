# Architecture Diagram

Visual system-level overview of the *Guest Sponsor Info* solution.
For the written design decisions behind each component, see [architecture.md](architecture.md).

---

## System Overview — Recommended Path (Azure Function Proxy)

The diagram covers two aspects: **how the web part reaches the guest's browser**
(delivery) and **how it retrieves and keeps data current at runtime** (steps ①–⑦).
Color-coding marks system boundaries at a glance:
**blue** = SharePoint Online · **amber** = Microsoft Entra ID ·
**green** = Azure Sponsor API · **purple** = Microsoft Graph.
Steps ②–③ make the authentication handshake explicit — the web part cannot call
the Sponsor API without first obtaining a signed token from Entra ID.
Presence status (step ⑦) is kept up-to-date through a separate polling loop that
reuses the same token and the same EasyAuth gate, but only fetches presence —
not the full sponsor list.

```mermaid
flowchart TB
    classDef admin    fill:#f1f5f9,stroke:#64748b,color:#1e293b,font-weight:bold
    classDef delivery fill:#dbeafe,stroke:#3b82f6,color:#1e3a8a
    classDef webpart  fill:#1d4ed8,stroke:#1e3a8a,color:#ffffff,font-weight:bold
    classDef token    fill:#fef3c7,stroke:#d97706,color:#78350f,font-weight:bold
    classDef gate     fill:#fde68a,stroke:#b45309,color:#78350f,font-weight:bold
    classDef func     fill:#d1fae5,stroke:#059669,color:#064e3b,font-weight:bold
    classDef infra    fill:#a7f3d0,stroke:#059669,color:#064e3b
    classDef logs     fill:#f8fafc,stroke:#94a3b8,color:#64748b
    classDef msgraph  fill:#ede9fe,stroke:#7c3aed,color:#4c1d95,font-weight:bold

    Admin(["🧑‍💼 SharePoint Admin"]):::admin

    subgraph spo["☁️ SharePoint Online"]
        Catalog["📦 App Catalog"]:::delivery
        CDN["🌐 Public CDN"]:::delivery
        Page["📄 Guest Landing Page"]:::delivery
        WP["🖥️ Guest Sponsor Info\nWeb Part"]:::webpart
    end

    subgraph entra["🔐 Microsoft Entra ID"]
        TokenSvc["🔑 Token Service\n(App Registration\nfor Sponsor API)"]:::token
    end

    subgraph azure["⚡ Azure · Sponsor API"]
        EasyAuth{"🛡️ EasyAuth\n(Azure App Service)\ntoken gate"}:::gate
        Func["⚡ Azure Function\n(sponsor lookup &\nbusiness logic)"]:::func
        MI["🔒 Managed Identity\n(no stored credentials)"]:::infra
        AI[("📊 Application\nInsights")]:::logs
    end

    Graph[("🕸️ Microsoft Graph\n(sponsors · profiles\nphotos · presence)")]:::msgraph

    Admin      -- "deploys"                              --> Catalog
    Catalog    -- "via"                                  --> CDN
    CDN        -- "① web part bundle"                    --> WP
    Page       -- "hosts"                                --> WP

    WP         -- "② request token\nscoped to Sponsor API"   --> TokenSvc
    TokenSvc   -- "signed Bearer token\n(identifies the guest)"  --> WP
    WP         -- "③ call with\nBearer token"            --> EasyAuth
    EasyAuth   -- "④ token valid —\nguest OID confirmed"  --> Func
    EasyAuth   -. "token invalid or missing —\nrequest rejected (HTTP 401)"  .-> WP
    Func       --> MI
    MI         -- "sponsors · profiles · presence\n(app permissions)"  --> Graph
    Func       -- "⑤ full sponsor list\n(one-time on load)"  --> WP
    Func       -. "telemetry"                            .-> AI
    WP         -- "⑥ profile photos\n(direct · delegated token)"  --> Graph

    WP         -. "⑦ presence poll\n(same Bearer token —\nsilently refreshed)"  .-> EasyAuth
    Func       -. "⑦ presence status only\n(no sponsor re-fetch)"  .-> WP

    style spo   fill:#eff6ff,stroke:#3b82f6
    style entra fill:#fffbeb,stroke:#d97706
    style azure fill:#f0fdf4,stroke:#059669

    %% link indices (declaration order, 0-based)
    %% 0–3   delivery: Admin→Catalog, Catalog→CDN, CDN→WP, Page→WP
    linkStyle 0,1,2,3   stroke:#94a3b8,stroke-width:1.5px
    %% 4–5   token roundtrip: WP→TokenSvc, TokenSvc→WP
    linkStyle 4,5       stroke:#d97706,stroke-width:2px
    %% 6     initial API call: WP→EasyAuth
    linkStyle 6         stroke:#1d4ed8,stroke-width:2.5px
    %% 7     valid path: EasyAuth→Func
    linkStyle 7         stroke:#059669,stroke-width:2.5px
    %% 8     rejection path: EasyAuth→WP
    linkStyle 8         stroke:#dc2626,stroke-width:1.5px
    %% 9–10  function→Graph via MI
    linkStyle 9,10      stroke:#7c3aed,stroke-width:2px
    %% 11    sponsor list response: Func→WP
    linkStyle 11        stroke:#059669,stroke-width:2px
    %% 12    telemetry: Func→AI
    linkStyle 12        stroke:#94a3b8,stroke-width:1px
    %% 13    photos: WP→Graph
    linkStyle 13        stroke:#3b82f6,stroke-width:2px
    %% 14–15 presence polling: WP→EasyAuth, Func→WP
    linkStyle 14,15     stroke:#0891b2,stroke-width:1.5px
```

### What each step means

| Step | What happens |
|---|---|
| ① | The guest opens the SharePoint landing page. The browser loads the web part bundle from the Public CDN — no App Catalog access needed at runtime. |
| ② | The web part silently requests a token from Entra ID, scoped specifically to the Sponsor API's App Registration. No extra guest consent is required — the scope is pre-authorized for SharePoint. |
| ③ | Only after a valid token is in hand does the web part call the Sponsor API, with the Bearer token attached. There is no direct path to the function without this token. |
| ④ | [EasyAuth](https://learn.microsoft.com/azure/app-service/overview-authentication-authorization) (Microsoft Azure App Service Authentication) intercepts the request at the Azure Function boundary and validates the token before any function code runs. An invalid or missing token is rejected immediately (HTTP 401); the function never sees the request. |
| ⑤ | The function identifies the guest from the EasyAuth-confirmed OID and calls Microsoft Graph using its own Managed Identity. It returns the full sponsor list — sponsors, profiles, and manager — in one response. This happens **once on page load**. |
| ⑥ | Profile photos are loaded **directly** from Graph using the guest's own delegated token. They bypass the function entirely. |
| ⑦ | After the initial load, the web part polls the Sponsor API for **presence status only** at adaptive intervals — **30 seconds** while a sponsor card is hovered, **2 minutes** while the browser tab is visible, **5 minutes** while the tab is in the background. The token is silently refreshed by the browser before it expires; the EasyAuth gate applies on every poll just as on the initial call. The full sponsor list is never re-fetched during polling. |

---

## Fallback Path — Direct Graph (legacy, no Azure Function)

When no Azure Function URL is configured, the web part calls Microsoft Graph
directly with the guest's delegated token. This requires the guest account to
hold an Entra directory role (*Directory Readers*) — impractical at scale.
The Azure Function proxy removes that requirement.

```mermaid
flowchart LR
    classDef webpart  fill:#1d4ed8,stroke:#1e3a8a,color:#ffffff,font-weight:bold
    classDef token    fill:#fef3c7,stroke:#d97706,color:#78350f,font-weight:bold
    classDef msgraph  fill:#ede9fe,stroke:#7c3aed,color:#4c1d95,font-weight:bold

    subgraph browser["💻 Guest's Browser"]
        WP2["🖥️ Guest Sponsor Info\nWeb Part"]:::webpart
    end

    subgraph entra2["🔐 Microsoft Entra ID"]
        TokenSvc2["🔑 Token Service"]:::token
    end

    Graph2[("🕸️ Microsoft Graph\n(delegated permissions)")]:::msgraph

    WP2 -- "acquire token" --> TokenSvc2
    WP2 -- "sponsors, profiles, photos\n(guest must hold Directory Readers role)" --> Graph2
    WP2 -. "presence (optional)" .-> Graph2

    style browser fill:#eff6ff,stroke:#3b82f6
    style entra2  fill:#fffbeb,stroke:#d97706

    linkStyle 0   stroke:#d97706,stroke-width:2px
    linkStyle 1   stroke:#3b82f6,stroke-width:2px
    linkStyle 2   stroke:#3b82f6,stroke-width:1.5px
```

---

## Component Summary

| Component | Role |
|---|---|
| SharePoint App Catalog | Stores the packaged solution; publishes assets to the CDN |
| Public CDN | Delivers the web part JavaScript bundle to the guest's browser |
| Web Part | Guest-facing UI rendered inside the SharePoint page |
| Token Service (Entra ID) | Issues tokens that identify the guest — no directory role needed |
| Sponsor API (Azure Function) | Secure proxy between the web part and Graph; enforces caller identity |
| [EasyAuth](https://learn.microsoft.com/azure/app-service/overview-authentication-authorization) | Microsoft Azure App Service Authentication — validates tokens at the function boundary before any code runs |
| Managed Identity | Allows the function to call Graph without any stored credentials |
| Microsoft Graph | Source of sponsor relationships, profiles, photos, and presence |
| Application Insights | Telemetry and structured error logs for the function |

---

## Related Documents

- [architecture.md](architecture.md) — design decisions, known limitations, SPFx lifecycle
- [deployment.md](deployment.md) — step-by-step deployment, Azure Function setup, hosting plans
- [development.md](development.md) — local dev setup, build & test commands
- [features.md](features.md) — feature descriptions and the problems they solve
- [README](../README.md) — quick-start and overview
- [Azure Function README](../azure-function/README.md) — function-specific permissions and security design
