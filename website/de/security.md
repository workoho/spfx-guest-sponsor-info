---
layout: page
lang: de
title: Vertrauen & Sicherheit
meta_title: Vertrauen & Sicherheit für Guest Sponsor Info
permalink: /de/security/
description: >-
  Überblick zu Vertrauen und Sicherheit für Guest Sponsor Info für Microsoft
  Entra B2B — Deployment-Grenzen, Security Controls, Disclosure-Prozess und
  welche Nachweise Enterprise-Kunden heute schon prüfen können.
lead: >-
  Praktischer Überblick für IT-Admins, Security-Reviewer und Procurement-Teams,
  die verstehen möchten, wo die Trust Boundary liegt und welche Assurance-
  Signale heute verfügbar sind.
intro_badges:
  - Azure-Deployment im Kundentenant
  - Keine Graph-Berechtigungen im SPFx-Paket
  - Öffentliche Security- und Privacy-Dokumentation
intro_actions:
  - label: Vollständige Sicherheitsbewertung lesen
    href: https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/security-assessment.md
    style: btn-primary
    external: true
  - label: Support-Optionen ansehen
    href: /de/support/
    style: btn-secondary
---

Diese Seite fasst die aktuelle Trust- und Security-Position von **Guest
Sponsor Info für Microsoft Entra B2B** und der begleitenden **Guest Sponsor
API** zusammen.

Sie dient als praktischer Einstieg für Enterprise-Reviews. Die vollständigen
technischen Details finden Sie in den verlinkten GitHub-Dokumenten.

## Warum das Deployment-Modell entscheidend ist

Die wichtigste Vertrauenseigenschaft dieser Lösung ist der Ort, an dem die
privilegierte Verarbeitung stattfindet.

- **Microsoft Graph Application Permissions bleiben serverseitig** in der Azure
  Function, nicht im SharePoint-Framework-Paket
- **Die Azure Function läuft in Ihrem eigenen Azure-Abonnement** mit Ihrem
  eigenen RBAC, Monitoring, Aufbewahrungs- und Compliance-Rahmen
- **Das Web Part selbst fordert in SharePoint keine Microsoft Graph-
  Berechtigungen an**
- **Workoho betreibt kein geteiltes Multi-Tenant-Backend** für Sponsor-Lookups

Für viele Enterprise-Kunden ist das wichtiger als ein Marketing-Badge: Der
Kunde kontrolliert die Azure-Ressourcen, die Identitätskonfiguration, das
Logging und das Runtime-Zugriffsmodell.

## Vorhandene Sicherheitskontrollen

Die aktuell empfohlene Bereitstellung nutzt ein mehrschichtiges Modell:

- Azure App Service Authentication blockiert nicht authentifizierte Anfragen,
  bevor der Function-Code ausgeführt wird
- Produktionsanfragen werden auf Tenant, Audience und erwartete Claims der
  aufrufenden Anwendung validiert
- Der Zugriff auf Microsoft Graph erfolgt über eine system-assigned Managed
  Identity, nicht über gespeicherte Secrets
- Folgeanfragen für Presence und Fotos sind auf die für den Aufrufer
  autorisierte Sponsor-Menge beschränkt
- CORS ist auf den SharePoint-Origin des Tenants eingegrenzt
- Logs werden redigiert und verbleiben im Azure-Abonnement des Kunden

Für die detaillierten Trust Boundaries, Restrisiken und Härtungsempfehlungen
siehe die
[vollständige Sicherheitsbewertung auf GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/security-assessment.md).

## Transparenz statt Black Box

Enterprise-Sicherheitsprüfungen gehen schneller, wenn die Architektur einfach
zu prüfen ist. Die folgenden Unterlagen sind öffentlich und versioniert:

- [Sicherheitsbewertung auf GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/security-assessment.md)
- [Datenschutzrichtlinie auf GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/privacy-policy.md)
- [Architekturdokumentation auf GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/architecture.md)
- [Repository-Sicherheitsrichtlinie auf GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/SECURITY.md)

Diese Unterlagen geben IT-Admins und Security-Teams konkrete Grundlagen für
eine fundierte Bewertung statt generischer Hersteller-Aussagen.

## Aktuelle Assurance-Signale

Die stärksten aktuellen Vertrauenssignale sind:

- transparente Architektur- und Sicherheitsdokumentation
- kundenkontrolliertes Deployment in die Microsoft-365- und Azure-Umgebung des
  Kunden
- Responsible-Disclosure-Kanäle für Schwachstellen
- eine klare Trennung zwischen clientseitigem Web Part, Azure Function und der
  Azure-Kontrollfläche des Kunden
- öffentliche Dokumentation, die interne Security-, Compliance- und
  Procurement-Prüfungen schon vor dem Deployment unterstützt

Zusammen machen diese Signale die Lösung leichter bewertbar, leichter
freigabefähig und leichter in Enterprise-Governance-Prozesse integrierbar.

Wo passend, kann die Distribution über den Microsoft Commercial Marketplace
zusätzlich die Auffindbarkeit und die interne Beschaffungsabstimmung erleichtern.

## Was Enterprise-Kunden prüfen können

Wenn Sie die Lösung intern bewerten, sind meist folgende Fragen relevant:

- Welche Microsoft-Graph-Berechtigungen werden benötigt, und wo liegen sie?
- Welche Azure-Ressourcen existieren, und wer administriert sie?
- Wie werden Anfragen authentifiziert und auf den angemeldeten Gast begrenzt?
- Wohin gehen Logs und Telemetrie?
- Welche Teile sind optional, etwa Teams-Presence oder Azure Maps?

Diese Antworten sind öffentlich dokumentiert und können vor dem Deployment
geprüft werden.

## Meldung von Sicherheitsproblemen

Potenzielle Schwachstellen sollten nicht als öffentliche GitHub-Issues gemeldet
werden.

Nutzen Sie die Disclosure-Hinweise in der
[Repository-Sicherheitsrichtlinie auf GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/SECURITY.md),
einschließlich des Private-Advisory-Wegs und des dedizierten Security-
Kontakts.

## Aktuelle Support-Optionen

Die freie Distribution enthält kein SLA und keine garantierte Reaktionszeit.
Aktuelle Support-Optionen für Rollout, Deployment und Betrieb finden Sie auf
der [Support-Seite](/de/support/).
