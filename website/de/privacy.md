---
layout: page
lang: de
title: Datenschutz
meta_title: Datenschutz für Guest Sponsor Info
permalink: /de/privacy/
description: >-
  Datenschutzrichtlinie für das Guest Sponsor Info SharePoint Web Part und die
  Guest Sponsor API für Microsoft Entra B2B — wie Daten verarbeitet werden,
  welche Microsoft-Graph-Berechtigungen verwendet werden und wo Daten bleiben.
lead: >-
  Verständlicher Überblick darüber, was in Ihrem Tenant bleibt, welche
  Microsoft-Graph-Berechtigungen genutzt werden und warum keine
  Laufzeitdaten an Workoho fließen.
intro_badges:
  - Daten bleiben im Tenant
  - Graph nur über Azure Function
  - Application Insights optional
intro_actions:
  - label: Telemetrie ansehen
    href: /de/telemetry/
    style: btn-secondary
  - label: Richtlinie auf GitHub lesen
    href: https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/privacy-policy.md
    style: btn-primary
    external: true
---

Diese Richtlinie gilt für **Guest Sponsor Info für Microsoft Entra B2B** (das SharePoint
Web Part) und die **Guest Sponsor API für Microsoft Entra B2B** (die begleitende Azure
Function). Beide sind mit einer **Privacy-First-Architektur** konzipiert.
Die gesamte Datenverarbeitung findet innerhalb Ihrer eigenen Microsoft 365
und Azure Tenant-Grenzen statt.

Wenn Sie speziell nach Azure-Nutzungszuordnung und Opt-out suchen, lesen Sie
die Seite [Telemetrie](/de/telemetry/). Für die operativen Schritte der
Bereitstellung lesen Sie die [Setup-Anleitung](/de/setup/).

## Grundprinzipien

- **Keine Daten an Workoho oder Dritte** — Web Part und Azure Function
  arbeiten vollständig innerhalb Ihres Tenants.
- **Nur Browser-Speicher** — das Web Part hält Sponsor-Daten (Name, Titel,
  E-Mail, Telefon, Teams-Präsenz) nur während der Seitensitzung im
  Browser-Speicher. Nichts wird auf der Festplatte gespeichert oder
  anderweitig versendet.
- **Azure Function ist zustandslos** — jede Anfrage wird verarbeitet und
  verworfen. Keine Sponsor- oder Gastdaten werden gespeichert.
- **Ihre Application Insights** — falls aktiviert, geht die Telemetrie
  in Ihr eigenes Azure-Abonnement. Workoho hat keinen Zugriff.

## Verwendete Berechtigungen

Alle Microsoft Graph-Berechtigungen liegen ausschließlich bei der
**Guest Sponsor API für Microsoft Entra B2B** — das Web Part selbst hat keine.

| Berechtigung | Erforderlich? | Zweck |
|---|---|---|
| [`User.Read.All`](https://learn.microsoft.com/en-us/graph/permissions-reference#userreadall) | Erforderlich | Sponsor-Profile lesen und deaktivierte Konten filtern |
| [`Presence.Read.All`](https://learn.microsoft.com/en-us/graph/permissions-reference#presencereadall) | Optional | Teams-Präsenz-Indikatoren |
| [`MailboxSettings.Read`](https://learn.microsoft.com/en-us/graph/permissions-reference#mailboxsettingsread) | Optional | Freigegebene Postfächer/Raum-/Gerätekonten filtern |
| [`TeamMember.Read.All`](https://learn.microsoft.com/en-us/graph/permissions-reference#teammemberreadall) | Optional | Teams-Konto-Provisionierung von Gästen erkennen |

## Vollständige Richtlinie

Die vollständige Datenschutzrichtlinie einschließlich Betroffenenrechte,
GitHub-Release-Prüfungen und Customer Usage Attribution finden Sie in der
[vollständigen Datenschutzrichtlinie auf GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/privacy-policy.md).

Für Sicherheitsbewertung und Vertrauensannahmen siehe die
[Sicherheitsbewertung auf GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/security-assessment.md).

Für Unterstützungsoptionen rund um Deployment, Anpassung oder Betrieb siehe
[Support](/de/support/).
