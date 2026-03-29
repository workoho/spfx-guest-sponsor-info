---
layout: doc
lang: de
title: Warum dieses Web Part existiert
permalink: /de/why/
description: >-
  Warum dieses Web Part entwickelt wurde — was passiert, wenn ein Gast zwar in
  Entra existiert, aber noch keine Teams-Präsenz im Mandanten hat.
lead: >-
  Wenn ein Gast eine Einladung annimmt, kann er zwar in Entra existieren, aber
  noch nicht in Teams. Kontaktschaltflächen werden angezeigt — funktionieren
  jedoch lautlos nicht. Dieses Web Part macht das sichtbar.
mermaid: false
---

## Nach der Einladung {#nach-der-einladung}

Wenn jemand außerhalb Ihrer Organisation eine Microsoft-365-Gasteinladung erhält
und auf „Annehmen" klickt, passiert etwas Konkretes in Entra ID: Ein Gastkonto
wird erstellt, die Einladung wird erfasst, und die Person ist technisch gesehen
in Ihrem Mandanten.

Was nicht zwingend sofort passiert: eine Präsenz in Microsoft Teams und Microsoft
Exchange für den Gast in Ihrem Mandanten.

---

## Zwei Wege, einen Gast einzuladen {#zwei-wege}

Wie ein Gast in Ihren Mandanten gelangt, bestimmt, was er direkt tun kann — und
wie schnell.

### Implizite Einladung

Eine interne Person fügt einen externen Kontakt direkt einem Teams-Team oder
einer SharePoint-Website hinzu. Microsoft sendet automatisch eine
Einladungsmail. Sobald der Gast annimmt, ist er dem Team hinzugefügt — und
Teams beginnt sofort, dessen Präsenz aufzubauen. **Der Gast ist innerhalb
weniger Minuten in Teams einsatzbereit.**

### Gesteuertes Lifecycle-Onboarding

Ein Bereitstellungsprozess — eine Governance-Plattform, ein Skript oder ein
Workflow im Entra Admin Center — legt das Gastkonto formal an. Das Konto
existiert in Entra. Aber niemand hat diese Person bisher einem Teams-Team
hinzugefügt. Teams hat für den Gast in Ihrem Mandanten noch keine Präsenz
aufgebaut.

**Der Gast existiert in Entra. Er existiert noch nicht in Teams.**

Das ist die Lücke. Sie ist von außen nicht sichtbar — es sei denn, irgendetwas
zeigt sie explizit an.

---

## Das Teams-Timing-Problem {#teams-timing-problem}

Wenn ein Gastkonto in Entra existiert, aber noch keinem Teams-Team hinzugefügt
wurde, ist dessen Teams-Präsenz in Ihrem Mandanten noch nicht aufgebaut.

Auf einer SharePoint-Begrüßungsseite könnte eine Sponsor-Visitenkarte folgendes
zeigen:

- Name und Profilfoto des Sponsors — ✓ verfügbar über Entra
- E-Mail-Adresse — ✓ verfügbar
- Chat-Schaltfläche — sichtbar angezeigt, aber **lautlos defekt**
- Anruf-Schaltfläche — sichtbar angezeigt, aber **lautlos defekt**

Der Gast sieht eine Visitenkarte, die funktionsbereit wirkt. Die Schaltflächen
tun nichts oder führen an einen unerwarteten Ort. Es gibt keine Fehlermeldung.
Keine Erklärung. Der Gast hat keine Möglichkeit zu erkennen, ob die Schaltfläche
defekt ist, ob er etwas falsch gemacht hat, oder ob die Funktion für ihn einfach
nicht vorgesehen ist.

Das ist kein Teams-Fehler. Es ist das erwartete Verhalten, wenn noch keine
Teams-Präsenz für den Gast aufgebaut wurde. Ohne etwas, das diesen Zustand
erkennt und kommuniziert, ist die Erfahrung für den Gast lautlos kaputt.

---

## Was dieses Web Part macht {#was-das-web-part-macht}

Das **Guest Sponsor Info** Web Part wurde für das Szenario einer Gast-Landingpage
entwickelt. Es wird im SharePoint Entrance Area platziert — der Seite, auf der
Gäste nach der Einladungsannahme landen.

Es macht zwei Dinge:

1. **Es zeigt die Sponsors des Gastes** — die internen Mitarbeiter, die in Entra
   als verantwortlich für den Gastzugang eingetragen sind. Name, Foto, Titel und
   Kontaktmöglichkeiten. Keine Konfiguration pro Gast. Keine manuelle
   Aktualisierung bei Personenwechsel.

2. **Es erkennt den Teams-Status** — wenn der Gast noch keine Teams-Präsenz in
   Ihrem Mandanten hat, erkennt das Web Part dies und passt die Visitenkarte
   entsprechend an: Chat- und Anruf-Schaltflächen werden ausgegraut, und eine
   klare Statusmeldung erklärt die Situation. Der Gast tappt nicht im Dunkeln.
   Er sieht ein Gesicht, einen Namen und eine ehrliche Statusanzeige.

Das Ergebnis: Ein Gast, dessen Zugang noch bereitgestellt wird, kann seinen
Sponsor trotzdem per E-Mail erreichen — und weiß, dass der Teams-Zugang in Kürze
folgt. Keine Verwirrung. Kein lautloses Scheitern.

---

## Was Sie benötigen {#voraussetzungen}

1. Eine SharePoint-Website als Gast-Landingpage (der Entrance Area).
2. Eine Azure-Function-Bereitstellung (die Guest Sponsor API) — darüber ruft das
   Web Part die Sponsor-Daten ab.
3. Gastkonten mit zugewiesenen Sponsors in Microsoft Entra ID.

Eine Schritt-für-Schritt-Anleitung finden Sie in der
[Installationsanleitung](/de/setup/).
