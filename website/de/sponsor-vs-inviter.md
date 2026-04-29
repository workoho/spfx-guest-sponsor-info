---
layout: doc
lang: de
title: Sponsor vs. Einladender
permalink: /de/sponsor-vs-inviter/
description: >-
  Verstehen Sie den Unterschied zwischen Sponsor und Einladendem im Microsoft
  Entra B2B-Gäste-Onboarding und warum eine SharePoint-Gast-Landingpage die
  Sponsor-Beziehung klar zeigen sollte.
lead: >-
  Im Microsoft Entra B2B-Gäste-Onboarding können Sponsor und Einladender
  dieselbe Person sein, sind es aber oft nicht. Diese Unterscheidung wird
  wichtig, sobald ein Gast auf der SharePoint-Landingpage einen verlässlichen
  Ansprechpartner braucht.
---

## Kurzantwort

Der **Einladende** ist die Person oder der Prozess, der die Einladung
ausgelöst hat.

Der **Sponsor** ist die Person oder Gruppe, die im **Sponsors**-Feld von
Microsoft Entra für diese Gastbeziehung eingetragen ist.

Manchmal sind beide Rollen identisch. In strukturierten Gäste-Onboarding-,
Lifecycle-Governance- oder Admin-getriebenen Prozessen sind es aber oft
unterschiedliche Rollen.

## Sponsor oder „Owner“?

Manche Produkte und interne Prozesse verwenden für denselben zuständigen
Kontakt oder dieselbe Gruppe das Wort **Owner**. Auf dieser Website
bevorzugen wir **Sponsor**, weil das der Begriff im **Sponsors**-Feld von
Microsoft Entra ist.

Sponsor bedeutet hier **nicht** Budget-Verantwortung, disziplinarische Rolle
oder besondere Weisungsbefugnis gegenüber dem Gast. Gemeint ist schlicht der
interne Kontakt oder die Gruppe, die für diese Gastbeziehung hinterlegt ist.

## Warum Gäste beides verwechseln

Gäste sehen zuerst meist den Einladenden, weil dessen Name in der Einladung,
im Freigabeprozess oder in der Antrags-Historie auftaucht. Die Sponsor-
Beziehung aus Microsoft Entra bleibt für sie dagegen meist unsichtbar.

Genau daraus entsteht ein praktisches Problem. Ein Gast erinnert sich vielleicht
noch daran, wer die Einladung ausgelöst hat, weiß aber trotzdem nicht,

- wer heute für seinen Zugriff zuständig ist
- wer einspringt, wenn der primäre Sponsor nicht erreichbar ist
- wen er kontaktieren soll, wenn das Onboarding hängt

Darum ist eine SharePoint-Gast-Landingpage so wertvoll: Sie kann der Ort sein,
an dem Sponsor-Sichtbarkeit überhaupt erst entsteht.

## Sponsor vs. Einladender im Überblick

| Frage | Sponsor | Einladender |
|---|---|---|
| Was ist die Rolle? | Person oder Gruppe im Sponsors-Feld von Microsoft Entra für diese Gastbeziehung | Person oder Prozess, der die Einladung ausgelöst hat |
| Was sieht der Gast typischerweise zuerst? | Meist standardmäßig unsichtbar | Meist im Einladungsprozess sichtbar |
| Ist das der richtige Kontakt für laufende Rückfragen? | Meist ja | Nicht unbedingt |
| Kann es mehrere zuständige Kontakte geben? | Ja, Microsoft Entra unterstützt bis zu fünf Sponsoren | Eher nein; es hängt am Einladungsereignis |
| Sollte eine Gast-Landingpage diesen Kontakt betonen? | Ja, weil das Sponsor-Sichtbarkeit schafft | Höchstens als Kontext, nicht als Hauptkontaktweg |

## Warum dieser Unterschied auf der Landingpage zählt

Wenn Ihre SharePoint-Gast-Landingpage nur den Namen aus der Einladungs-Mail
wiederholt, bleibt das eigentliche Onboarding-Problem ungelöst.

Der Gast weiß dann immer noch nicht, wer heute für seinen Zugriff zuständig
ist, wer bei Ausfall des primären Sponsors einspringt oder ob Teams für den
Gastzugriff bereits bereitsteht.

Sponsor-Sichtbarkeit löst damit ein dauerhaftes Problem, nicht nur die Frage,
wer irgendwann einmal auf "Einladen" geklickt hat.

## Was Guest Sponsor Info stattdessen zeigt

Guest Sponsor Info schließt genau diese Lücke. Auf der SharePoint-Landingpage
kann das Web Part zeigen:

- die dem Gast in Microsoft Entra zugewiesenen Sponsor-Kontakte
- Ersatz-Sponsoren, wenn mehrere zuständige Kontakte existieren
- Manager-Kontext und Kontaktdetails
- einen ehrlichen Teams-Onboarding-Status, wenn Chat oder Anruf noch nicht
  bereit sind

So entsteht eine verlässliche Kontaktfläche während des Gäste-Onboardings und
auch danach.

## Wenn Sponsor und Einladender identisch sind

Es gibt einfache Einladungsprozesse, in denen dieselbe Person sowohl die
Einladung auslöst als auch dauerhaft Sponsor bleibt. In standardmäßigen Entra-
Einladungsflüssen ist das sogar der Standard, sofern kein anderer Sponsor
angegeben wird. SharePoint-Freigabeeinladungen an neue externe Benutzer sind
eine dokumentierte Ausnahme: Dort müssen Sponsoren gegebenenfalls manuell
ergänzt werden.

Selbst dann ist es sinnvoll, die Sponsor-Beziehung ausdrücklich anzuzeigen.
Damit wird sichtbar, dass diese Person nicht nur der historische Absender der
Einladungs-Mail ist, sondern der aktuelle zuständige Kontakt im Tenant.

<div class="doc-cta-box">
  <div>
    <p class="doc-cta-title">Die Sponsor-Beziehung bewusst sichtbar machen</p>
    <p class="doc-cta-sub">Starten Sie mit dem Warum oder gehen Sie direkt ins Setup.</p>
  </div>
  <div class="doc-cta-actions">
    <a href="{{ '/de/why/' | relative_url }}" class="btn btn-outline">Warum dieses Web Part?</a>
    <a href="{{ '/de/setup/' | relative_url }}" class="btn btn-teal">Setup-Anleitung</a>
  </div>
</div>
