---
layout: page
lang: de
title: Datenschutz
permalink: /de/privacy/
description: >-
  Datenschutzrichtlinie für Guest Sponsor Info für Microsoft Entra B2B und
  die Guest Sponsor API für Microsoft Entra B2B — wie Daten verarbeitet
  werden, welche Berechtigungen verwendet werden und wo Daten gespeichert werden.
---

Diese Richtlinie gilt für **Guest Sponsor Info für Microsoft Entra B2B** (das SharePoint
Web Part) und die **Guest Sponsor API für Microsoft Entra B2B** (die begleitende Azure
Function). Beide sind mit einer **Privacy-First-Architektur** konzipiert.
Die gesamte Datenverarbeitung findet innerhalb Ihrer eigenen Microsoft 365
und Azure Tenant-Grenzen statt.

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
