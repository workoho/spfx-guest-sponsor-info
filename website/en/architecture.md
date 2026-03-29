---
layout: doc
lang: en
title: Architecture Diagram
permalink: /en/architecture/
description: >-
  Visual system-level overview of the Guest Sponsor Info solution —
  setup flow and runtime flow.
lead: >-
  Interactive diagrams showing who configures what during deployment
  and what happens each time a guest opens the landing page.
mermaid: true
github_doc: architecture-diagram.md
---

## Setup — Two Admin Roles

Two separate admin personas are involved in setting up the solution.
The **SharePoint Admin** only needs the standard SharePoint Administrator role.
The **Azure Admin** covers three distinct responsibilities — Azure resource
deployment, Entra ID app configuration, and Graph permission grants — each
requiring different elevated permissions.

```mermaid
flowchart LR
    classDef admin    fill:#f1f5f9,stroke:#64748b,color:#1e293b,font-weight:bold
    classDef delivery fill:#dbeafe,stroke:#3b82f6,color:#1e3a8a
    classDef token    fill:#fef3c7,stroke:#d97706,color:#78350f,font-weight:bold
    classDef gate     fill:#fde68a,stroke:#b45309,color:#78350f,font-weight:bold
    classDef func     fill:#d1fae5,stroke:#059669,color:#064e3b,font-weight:bold
    classDef infra    fill:#a7f3d0,stroke:#059669,color:#064e3b
    classDef logs     fill:#f8fafc,stroke:#94a3b8,color:#64748b
    classDef msgraph  fill:#ede9fe,stroke:#7c3aed,color:#4c1d95,font-weight:bold

    SpAdmin(["🧑‍💼 SharePoint Admin"]):::admin
    AzAdmin(["🧑‍💼 Azure Admin"]):::admin

    subgraph spo["☁️ SharePoint Online"]
        SCAC["📦 Site Collection App Catalog (landing page site)"]:::delivery
        Visitors["👥 Visitors (landing page site)"]:::delivery
    end

    subgraph entra["🔐 Microsoft Entra ID"]
        AppReg["🔑 App Registration (EasyAuth)"]:::token
        SP["🪪 Service Principal (Enterprise App)"]:::token
    end

    subgraph azure["⚡ Guest Sponsor API"]
        Func["⚡ Azure Function"]:::func
        MI["🔒 Managed Identity"]:::infra
        AI[("📊 App Insights")]:::logs
    end

    Graph[("🕸️ Microsoft Graph")]:::msgraph

    SpAdmin -- "① uploads .sppkg"              --> SCAC
    SpAdmin -- "② grants access (Everyone)"    --> Visitors

    AzAdmin -- "③ creates App Registration"    --> AppReg
    AppReg  -. "auto-creates"                .-> SP
    AzAdmin -- "④ deploys function"           --> Func
    Func    -. "EasyAuth bound to"            .-> SP
    AzAdmin -- "⑤ grants permissions"         --> MI
    AzAdmin -. "⑤ configures"               .-> SP
    Func    -. "uses"                        .-> MI
    MI      -- "Graph app permissions"        --> Graph
    AzAdmin -. "connects"                    .-> AI

    style spo   fill:#eff6ff,stroke:#3b82f6
    style entra fill:#fffbeb,stroke:#d97706
    style azure fill:#f0fdf4,stroke:#059669
```

### Required permissions

| Step | Who | What happens | Required role |
|---|---|---|---|
| ① | SharePoint Admin | Enables Site Collection App Catalog and uploads `.sppkg` | **SharePoint Administrator** + Site Collection Admin |
| ② | SharePoint Admin | Verifies or sets up guest Visitor access | **SharePoint Administrator** |
| ③ | Azure Admin | Creates the App Registration (`setup-app-registration.ps1`) | **Application Administrator** |
| ④ | Azure Admin | Deploys ARM template — Azure resources + Storage role assignments | **Owner** on target resource group |
| ⑤ | Azure Admin | Grants Graph permissions to Managed Identity (`setup-graph-permissions.ps1`) | **Privileged Role Administrator** |

---

## Runtime — Guest Experience

Color-coding marks system boundaries at a glance:
**blue** = SharePoint Online · **amber** = Microsoft Entra ID ·
**green** = Guest Sponsor API · **purple** = Microsoft Graph.

```mermaid
flowchart TB
    classDef delivery fill:#dbeafe,stroke:#3b82f6,color:#1e3a8a
    classDef webpart  fill:#1d4ed8,stroke:#1e3a8a,color:#ffffff,font-weight:bold
    classDef token    fill:#fef3c7,stroke:#d97706,color:#78350f,font-weight:bold
    classDef gate     fill:#fde68a,stroke:#b45309,color:#78350f,font-weight:bold
    classDef func     fill:#d1fae5,stroke:#059669,color:#064e3b,font-weight:bold
    classDef infra    fill:#a7f3d0,stroke:#059669,color:#064e3b
    classDef logs     fill:#f8fafc,stroke:#94a3b8,color:#64748b
    classDef msgraph  fill:#ede9fe,stroke:#7c3aed,color:#4c1d95,font-weight:bold

    subgraph spo["☁️ SharePoint Online"]
        CDN["🌐 Public CDN"]:::delivery
        Page["📄 Guest Landing Page"]:::delivery
        WP["🖥️ Guest Sponsor Info Web Part"]:::webpart
    end

    subgraph entra["🔐 Microsoft Entra ID"]
        TokenSvc["🔑 Token Service (App Registration)"]:::token
    end

    subgraph azure["⚡ Guest Sponsor API"]
        EasyAuth{"🛡️ EasyAuth (Azure App Service)"}:::gate
        Func["⚡ Azure Function (sponsor lookup)"]:::func
        MI["🔒 Managed Identity"]:::infra
        AI[("📊 App Insights")]:::logs
    end

    Graph[("🕸️ Microsoft Graph")]:::msgraph

    CDN        -- "① web part bundle"                    --> WP
    Page       -- "hosts"                                --> WP

    WP         -- "② request token (Guest Sponsor API scope)"  --> TokenSvc
    TokenSvc   -- "signed Bearer token"                  --> WP
    WP         -- "③ call with Bearer token"             --> EasyAuth
    EasyAuth   -- "④ token valid — OID confirmed"        --> Func
    EasyAuth   -. "token invalid → HTTP 401"             .-> WP
    Func       --> MI
    MI         -- "sponsors · profiles · presence (app perms)" --> Graph
    Func       -- "⑤ full sponsor list (one-time)"       --> WP
    Func       -. "telemetry"                            .-> AI
    WP         -- "⑥ profile photos (delegated · direct)" --> Graph

    WP         -. "⑦ presence poll (token auto-refreshed)" .-> EasyAuth
    Func       -. "⑦ presence status only"               .-> WP

    style spo   fill:#eff6ff,stroke:#3b82f6
    style entra fill:#fffbeb,stroke:#d97706
    style azure fill:#f0fdf4,stroke:#059669
```

### What each step means

| Step | What happens |
|---|---|
| ① | The guest opens the SharePoint landing page. The browser loads the web part bundle from the Public CDN. |
| ② | The web part silently requests a token from Entra ID, scoped to the Guest Sponsor API. No extra guest consent required. |
| ③ | The web part calls the Guest Sponsor API with the Bearer token attached. |
| ④ | [EasyAuth](https://learn.microsoft.com/azure/app-service/overview-authentication-authorization) validates the token before any function code runs. Invalid tokens are rejected immediately (HTTP 401). |
| ⑤ | The function identifies the guest from the EasyAuth-confirmed OID and calls Microsoft Graph using its Managed Identity. Returns the full sponsor list in one response. |
| ⑥ | Profile photos are loaded **directly** from Graph using the guest's own delegated token — they bypass the function entirely. |
| ⑦ | After the initial load, presence is polled at adaptive intervals: **30 s** (card hovered) · **2 min** (tab visible) · **5 min** (tab hidden). |

---

## Component Summary

| Component | Role |
|---|---|
| SharePoint App Catalog | Stores the packaged solution; publishes assets to the CDN |
| Public CDN | Delivers the web part JavaScript bundle to the guest's browser |
| Web Part | Guest-facing UI rendered inside the SharePoint page |
| Token Service (Entra ID) | Issues tokens that identify the guest |
| Guest Sponsor API | Secure proxy; validates caller identity via EasyAuth, calls Graph using Managed Identity |
| EasyAuth | Azure App Service Authentication — validates tokens at the function boundary |
| Managed Identity | Allows the function to call Graph without stored credentials |
| Microsoft Graph | Source of sponsor relationships, profiles, photos, and presence |
| Application Insights | Telemetry and structured error logs for the function |
