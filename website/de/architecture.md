---
layout: doc
lang: de
title: Architekturdiagramm
permalink: /de/architecture/
description: >-
  Visuelle Systemübersicht der Guest Sponsor Info Lösung —
  Setup-Ablauf und Laufzeit-Ablauf.
lead: >-
  Interaktive Diagramme, die zeigen, wer was bei der Bereitstellung
  konfiguriert und was passiert, wenn ein Gast die Landingpage öffnet.
mermaid: true
github_doc: architecture-diagram.md
---

## Setup — Zwei Admin-Rollen

Zwei separate Admin-Personas sind an der Einrichtung beteiligt.
Der **SharePoint-Admin** benötigt nur die Standard-SharePoint-Administrator-Rolle.
Der **Azure-Admin** übernimmt drei verschiedene Aufgaben — Azure-Ressourcen-
Bereitstellung, Entra-ID-App-Konfiguration und Graph-Berechtigungsvergabe —
die jeweils unterschiedliche erweiterte Berechtigungen erfordern.

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
        SCAC["📦 Site Collection App Catalog (Landingpage-Site)"]:::delivery
        Visitors["👥 Besucher (Landingpage-Site)"]:::delivery
    end

    subgraph entra["🔐 Microsoft Entra ID"]
        AppReg["🔑 App-Registrierung (EasyAuth)"]:::token
        SP["🪪 Dienstprinzipal (Enterprise App)"]:::token
    end

    subgraph azure["⚡ Guest Sponsor API"]
        Func["⚡ Azure Function"]:::func
        MI["🔒 Managed Identity"]:::infra
        AI[("📊 App Insights")]:::logs
    end

    Graph[("🕸️ Microsoft Graph")]:::msgraph

    SpAdmin -- "① lädt .sppkg hoch"              --> SCAC
    SpAdmin -- "② gewährt Zugriff (Everyone)"     --> Visitors

    AzAdmin -- "③ erstellt App-Registrierung"     --> AppReg
    AppReg  -. "erstellt automatisch"            .-> SP
    AzAdmin -- "④ stellt Function bereit"         --> Func
    Func    -. "EasyAuth gebunden an"            .-> SP
    AzAdmin -- "⑤ erteilt Berechtigungen"         --> MI
    AzAdmin -. "⑤ konfiguriert"                 .-> SP
    Func    -. "verwendet"                       .-> MI
    MI      -- "Graph-App-Berechtigungen"         --> Graph
    AzAdmin -. "verbindet"                       .-> AI

    style spo   fill:#eff6ff,stroke:#3b82f6
    style entra fill:#fffbeb,stroke:#d97706
    style azure fill:#f0fdf4,stroke:#059669
```

### Erforderliche Berechtigungen

| Schritt | Wer | Was passiert | Erforderliche Rolle |
|---|---|---|---|
| ① | SharePoint-Admin | Aktiviert Site Collection App Catalog und lädt `.sppkg` hoch | **SharePoint-Administrator** + Websitesammlungsadministrator |
| ② | SharePoint-Admin | Prüft oder richtet Gast-Besucherzugriff ein | **SharePoint-Administrator** |
| ③ | Azure-Admin | Erstellt die App-Registrierung (`setup-app-registration.ps1`) | **Anwendungsadministrator** |
| ④ | Azure-Admin | Stellt ARM-Vorlage bereit — Azure-Ressourcen + Speicher-Rollenzuweisungen | **Besitzer** der Ziel-Ressourcengruppe |
| ⑤ | Azure-Admin | Erteilt Graph-Berechtigungen an Managed Identity (`setup-graph-permissions.ps1`) | **Administrator für privilegierte Rollen** |

---

## Laufzeit — Gast-Erlebnis

Farbcodierung kennzeichnet Systemgrenzen:
**Blau** = SharePoint Online · **Bernstein** = Microsoft Entra ID ·
**Grün** = Guest Sponsor API · **Violett** = Microsoft Graph.

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
        Page["📄 Gast-Landingpage"]:::delivery
        WP["🖥️ Guest Sponsor Info Web Part"]:::webpart
    end

    subgraph entra["🔐 Microsoft Entra ID"]
        TokenSvc["🔑 Token-Dienst (App-Registrierung)"]:::token
    end

    subgraph azure["⚡ Guest Sponsor API"]
        EasyAuth{"🛡️ EasyAuth (Azure App Service)"}:::gate
        Func["⚡ Azure Function (Sponsor-Abfrage)"]:::func
        MI["🔒 Managed Identity"]:::infra
        AI[("📊 App Insights")]:::logs
    end

    Graph[("🕸️ Microsoft Graph")]:::msgraph

    CDN        -- "① Web-Part-Bundle"                     --> WP
    Page       -- "hostet"                                --> WP

    WP         -- "② fordert Token an (Guest Sponsor API Scope)"  --> TokenSvc
    TokenSvc   -- "signiertes Bearer-Token"               --> WP
    WP         -- "③ Aufruf mit Bearer-Token"             --> EasyAuth
    EasyAuth   -- "④ Token gültig — OID bestätigt"        --> Func
    EasyAuth   -. "Token ungültig → HTTP 401"            .-> WP
    Func       --> MI
    MI         -- "Sponsoren · Profile · Präsenz (App-Berechtigungen)" --> Graph
    Func       -- "⑤ vollständige Sponsorenliste (einmalig)" --> WP
    Func       -. "Telemetrie"                           .-> AI
    WP         -- "⑥ Profilfotos (delegiert · direkt)"    --> Graph

    WP         -. "⑦ Präsenz-Polling (Token wird automatisch erneuert)" .-> EasyAuth
    Func       -. "⑦ nur Präsenzstatus"                  .-> WP

    style spo   fill:#eff6ff,stroke:#3b82f6
    style entra fill:#fffbeb,stroke:#d97706
    style azure fill:#f0fdf4,stroke:#059669
```

### Was jeder Schritt bedeutet

| Schritt | Was passiert |
|---|---|
| ① | Der Gast öffnet die SharePoint-Landingpage. Der Browser lädt das Web-Part-Bundle vom Public CDN. |
| ② | Das Web Part fordert lautlos ein Token von Entra ID an, das auf die Guest Sponsor API beschränkt ist. Keine zusätzliche Einwilligung erforderlich. |
| ③ | Das Web Part ruft die Guest Sponsor API mit dem angehängten Bearer-Token auf. |
| ④ | [EasyAuth](https://learn.microsoft.com/azure/app-service/overview-authentication-authorization) validiert das Token, bevor Function-Code ausgeführt wird. Ungültige Tokens werden sofort abgelehnt (HTTP 401). |
| ⑤ | Die Function identifiziert den Gast anhand der EasyAuth-bestätigten OID und ruft Microsoft Graph mit ihrer Managed Identity auf. Gibt die vollständige Sponsorenliste in einer Antwort zurück. |
| ⑥ | Profilfotos werden **direkt** aus Graph mit dem eigenen delegierten Token des Gastes geladen — sie umgehen die Function vollständig. |
| ⑦ | Nach dem initialen Laden wird Präsenz in adaptiven Intervallen abgefragt: **30 s** (Karte überfahren) · **2 Min.** (Tab sichtbar) · **5 Min.** (Tab verborgen). |

---

## Komponentenübersicht

| Komponente | Aufgabe |
|---|---|
| SharePoint App Catalog | Speichert die paketierte Lösung; veröffentlicht Assets im CDN |
| Public CDN | Liefert das Web-Part-JavaScript-Bundle an den Browser des Gastes |
| Web Part | Gastbezogene Benutzeroberfläche in der SharePoint-Seite |
| Token-Dienst (Entra ID) | Stellt Tokens aus, die den Gast identifizieren |
| Guest Sponsor API | Sicherer Proxy; validiert Identität via EasyAuth, ruft Graph mit Managed Identity auf |
| EasyAuth | Azure App Service Authentication — validiert Tokens an der Function-Grenze |
| Managed Identity | Ermöglicht der Function Graph-Aufrufe ohne gespeicherte Anmeldeinformationen |
| Microsoft Graph | Quelle für Sponsor-Beziehungen, Profile, Fotos und Präsenz |
| Application Insights | Telemetrie und strukturierte Fehlerprotokolle der Function |
