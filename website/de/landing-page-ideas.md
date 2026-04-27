---
layout: doc
lang: de
title: Landingpage-Ideen
permalink: /de/landing-page-ideas/
description: >-
  Zusätzliche Ideen für eine SharePoint-Gast-Landingpage in Microsoft Entra
  B2B — besonders Quick-Links-Bereiche, tenant-fixierte
  Microsoft-365-Deeplinks und unterstützende Links rund um
  Sponsor-Sichtbarkeit.
lead: >-
  Ein praktischer Begleiter für Administratoren, die ihre
  SharePoint-Gast-Landingpage für Microsoft-Entra-B2B-Gäste-Onboarding
  hilfreicher machen wollen. Diese Ideen ergänzen das Guest Sponsor Info Web
  Part; hilfreich sind sie, Voraussetzung für das Produkt aber nicht.
---

## So Nutzen Sie Diese Ideen

Diese Seite ist bewusst praktisch gehalten. Sie soll keinen vollständigen
Blueprint für eine Landingpage vorgeben, sondern bewährte Ideen dafür sammeln,
was eine gemeinsame Landingpage für Gäste zusätzlich zum Sponsor-Web-Part noch
enthalten kann.

Wichtig ist die Rollenverteilung: **Guest Sponsor Info löst die Sponsor-
Sichtbarkeit**. Die umgebenden Landingpage-Elemente lösen Orientierung,
SharePoint-Gastzugriff, Microsoft-365-Einstiegspunkte und Self-Service. Erst
zusammen wird aus einer generischen Ankunftsseite ein sinnvoller Einstieg für
B2B-Gäste.

Wenn Sie das SharePoint-**Quick-Links**-Web-Part verwenden, lässt sich vieles
davon schon ohne eigene Entwicklung umsetzen. Mit aktiviertem Audience
Targeting kann dieselbe Landingpage unterschiedliche Links für Mitarbeitende
und Gäste anzeigen.

Nicht jedes SharePoint-Web-Part unterstützt Audience Targeting gleichermaßen.
Quick Links ist hier meist das verlässlichste Arbeitspferd.

Die Beispiele unten verwenden Platzhalter wie `<tenant-id>`, `<tenant-name>`
und `<tenant-domain>`. Ersetzen Sie diese durch Ihre eigenen Werte.

## Ein Gutes Standardmuster

Für viele Microsoft-Entra-B2B-Tenants reichen bereits zwei
Quick-Links-Web-Parts für eine praxistaugliche SharePoint-Gast-Landingpage:

- Ein Bereich für **Microsoft 365**-Einstiegspunkte im richtigen Tenant.
- Ein Bereich für **Mein Gastkonto**-Self-Service-Aktionen.
- Optional ein kleines separates Quick-Links-Web-Part für **Weitere Apps**,
  damit dieser Link wie eine Nebenfunktion wirkt und nicht wie das Hauptziel.

Das funktioniert besonders gut, wenn auch die Landingpage selbst Audience
Targeting verwendet. Mitarbeitende sehen interne Ressourcen, während Gäste nur
die Links sehen, die ihnen im Ressourcen-Tenant tatsächlich helfen.

Unabhängig von diesen ausgehenden Links kann die Landingpage selbst
ebenfalls die Hub Site für den Gastbereich sein. Damit gewinnen Sie eine
gemeinsame Navigationsschicht, Hub-Branding-Optionen und eine klarere
Identität für den gesamten Bereich, auch wenn noch keine weiteren Sites
zugeordnet sind.

Das hilft auch bei der Benennung. So kann die zugrunde liegende Site zum
Beispiel weiterhin einen freundlichen Titel wie `Welcome @ Contoso` tragen,
während die Hub-Identität und die Hub-Navigation den größeren Bereich als
`Entrance Area` sichtbar machen.

Wenn später weitere Sites hinzukommen, wird diese Hub Site noch nützlicher.
Sie können Links zu zugeordneten oder nicht zugeordneten Sites in der Hub-
Navigation ergänzen und Audience Targeting nutzen, damit Mitarbeitende und
Gäste nicht dieselben Hub-Links sehen müssen.

So entsteht auf der Seite eine klare Aufgabenverteilung:

- das Sponsor-Web-Part beantwortet **wer hilft mir**
- die Quick-Links-Bereiche beantworten **wo gehe ich als Nächstes hin**
- die Konto-Links beantworten **was kann ich selbst lösen**

Der folgende Screenshot zeigt eine mögliche Komposition: eine gemeinsame
Gast-Landingpage mit tenant-fixierten Quick Links weit oben und dem Sponsor-
Bereich weiter unten auf der Seite.

<img src="{{ '/assets/images/entrance-landingpage-example.jpg' | relative_url }}" alt="Beispiel-Screenshot einer Landingpage.">

## Die Seite Leicht Wiederauffindbar Machen

Gäste profitieren von einer Seite, die sie ohne Reibungsverlust wiederfinden.

- Wenn die Landingpage später auf der Root Site (`/`) liegt, ist die URL oft
  schon für sich genommen leicht merkbar.
- Wenn das nicht möglich ist, kann sich eine gut merkbare Kurz-URL oder ein
  Shortlink auf diese Landingpage lohnen.
- Platzieren Sie weit oben auf der Seite eine sichtbare Call-to-Action, die den
  Gast nach dem ersten erfolgreichen Anmelden dazu auffordert, sich diese Seite
  zu bookmarken.
- Nennen Sie die Seite weiterhin in Einladungs- und Onboarding-Mails, gehen Sie
  aber nicht davon aus, dass Gäste diese E-Mails dauerhaft aufbewahren oder
  später noch einmal danach suchen wollen.

Da Browser keinen einheitlichen "Lesezeichen hinzufügen"-Link anbieten, der
überall sauber funktioniert, ist das meist besser als einfacher Hinweis oder
Callout gelöst und nicht als spezieller Skript-Button.

## Weitere Inhaltsbausteine, Die Sich Lohnen

Neben den Quick-Links-Bereichen helfen ein paar kleine Inhaltsblöcke dabei,
aus einer funktionalen Seite eine wirklich orientierende Seite zu machen.

- Eine kurze Begrüßung mit Kontext: zwei oder drei Sätze reichen oft schon,
  damit der Gast sofort versteht, in welcher Organisation und in welchem
  Kollaborationsrahmen er gelandet ist.
- Ein klarer erster Schritt statt einer Komplettübersicht: nennen Sie die eine
  Teamseite, den einen Kanal oder die eine Projektfläche, die beim ersten
  Besuch wirklich zählt.
- Kuratierte Ressourcen statt Sitemap: zeigen Sie die wenigen Links, die in der
  ersten Woche relevant sind, und nicht jede App, die es theoretisch gibt.
- Eine kleine News- oder Hinweisfläche: wenn auf der Seite später auch
  Wartungen, Policy-Änderungen oder projektbezogene Updates auftauchen, lohnt
  sich das Wiederkommen und Bookmarken deutlich mehr.
- Eine echte Kontaktmöglichkeit: Name und Kanal helfen mehr als eine anonyme
  Funktionsmailbox. Genau hier ergänzt das Sponsor-Web-Part den Rest der Seite.

Wenn Sie einen Begrüßungstext nur für Gäste zeigen möchten, kann Quick Links
auch dafür als pragmatischer Workaround dienen: Link zurück auf dieselbe Seite,
unauffälliges Layout, Audience Targeting auf dem Web-Part.

## Sprache, Branding Und Seitenidentität

Orientierung entsteht nicht nur durch Links, sondern schon in den ersten
Sekunden über Sprache und Erscheinungsbild.

- Wenn Ihre Landingpage international genutzt wird, ist Englisch als
  Standardsprache der Site Collection meist die sicherste Basis. Dieser Wert
  lässt sich nach dem Erstellen nicht mehr ändern.
- Veröffentlichen Sie zusätzliche Sprachversionen für wichtige Zielgruppen.
  Gerade bei Gästen zahlt sich das schneller aus, als man zunächst vermutet.
- Stellen Sie sicher, dass Organisationsname, Logo, globale Navigation und
  SharePoint-Theme sauber konfiguriert sind. Branding beantwortet sofort die
  Frage, in wessen Umgebung der Gast gerade gelandet ist.
- Wenn Sie die Landingpage als Root Site oder als Hub Site nutzen, wird diese
  Identität noch stärker. Das hilft nicht nur bei der Orientierung, sondern
  auch beim späteren Wiederfinden.

## Bereich 1 — Microsoft 365

Dieser Bereich gibt Gästen stabile Einstiegspunkte in den Ressourcen-Tenant.
Wo eine Microsoft-URL `tenantId` unterstützt, sollten Sie sie verwenden. Bei
SharePoint legt der Tenant-Hostname den Tenant-Kontext bereits selbst fest.

Gerade im Microsoft-Entra-B2B-Gäste-Onboarding ist das wichtiger, als es erst
einmal klingt. Gäste wissen oft, dass sie eingeladen wurden, aber nicht,
welcher tenant-spezifische Zielort ihr verlässlicher Startpunkt sein soll.

### Microsoft Teams

Verwenden Sie einen tenant-fixierten Teams-Einstiegslink, wenn Teams im
richtigen Ressourcen-Tenant geöffnet werden soll und nicht in dem Tenant, der
zufällig gerade zuvor aktiv war.

```text
https://teams.cloud.microsoft/?tenantId=<tenant-id>
```

Das ist hilfreich, weil es nicht voraussetzt, dass der Gast den Tenant-Wechsel
in Teams bereits manuell beherrscht. Außerdem vermeiden Sie damit Team-
spezifische Deeplinks, solange Sie nicht sicher wissen, dass die Team-
Mitgliedschaft bereits existiert. Microsoft dokumentiert ausdrücklich, dass
Gastfunktionen in Teams erst verfügbar werden, wenn der Gast Mitglied in
mindestens einem Team ist.

### Microsoft SharePoint

Verlinken Sie auf eine tenant-eigene Übersichtsseite, eine andere Hub Site
oder eine Site-Navigation, über die Gäste gemeinsame Arbeitsbereiche und
Speicherorte finden können, ohne zuerst durch Teams navigieren zu müssen.

```text
https://<tenant-name>.sharepoint.com/teams/overview
```

In manchen Tenants ist das eine Hub Site. In anderen ist es einfach eine
kuratierte Übersichtsseite. Beides ist in Ordnung. Entscheidend ist, dass die
URL durch den SharePoint-Hostnamen bereits tenant-fixiert ist. Außerdem hilft
sie dabei, Teams-nahe Speicherorte oder normale Team Sites zu finden, die gar
nicht "teamified" sind.

Das ist auch einer der Gründe, warum eine SharePoint-Landingpage als erster
Zielort so stark ist: sie funktioniert zuverlässig, bevor für einen Gast jede
Teams-Funktion im Ressourcen-Tenant wirklich bereitsteht.

### Viva Engage

Wenn Ihr Tenant Viva Engage als breitere Community-Ebene nutzt, kann das ein
sinnvoller paralleler Einstiegspunkt neben Teams und SharePoint sein.

```text
https://engage.cloud.microsoft/main/org/<tenant-domain>/
```

Das funktioniert am besten, wenn der Gast dort tatsächlich auf relevante
Communities zugreifen kann. Andernfalls sollten Sie den Link per Audience
Targeting aussteuern oder ganz weglassen.

### Weitere Apps

Ein tenant-fixierter Link zu "My Applications" ist als Fallback nützlich. Auf
der Landingpage funktioniert er aber meist besser als sekundäre Aktion und
nicht als primärer Einstiegspunkt.

```text
https://myapplications.microsoft.com/?tenantId=<tenant-id>
```

Ein gutes Muster ist, diesen Link in einem kleinen separaten Quick-Links-Web-
Part ohne sichtbare Abschnittsüberschrift darzustellen. Dann wirkt er eher wie
ein zusätzlicher Utility-Link als wie der Hauptpfad.

Wenn Sie My Applications aktiv pflegen, kann dort zusätzlich ein gut sichtbarer
Link zurück zur Entrance Area sinnvoll sein. My Applications ist ein nützlicher
Fallback, aber selten die beste Startseite.

## Bereich 2 — Mein Gastkonto

Dieser Bereich konzentriert sich auf Self-Service. Er hilft Gästen dabei, ihr
Konto direkt im richtigen Ressourcen-Tenant zu verwalten, ohne erst selbst den
Tenant-Wechsel verstehen zu müssen.

Auf einer gut gestalteten SharePoint-Gast-Landingpage ergänzen diese Links den
Sponsor-Bereich, statt mit ihm zu konkurrieren. Die Sponsor-Beziehung zeigt,
wer für den Zugriff zuständig ist. Manche Tools nennen dieselbe Rolle
„Owner“; hier verwenden wir den Microsoft-Begriff Sponsor. Die Self-Service-
Links helfen bei den Themen, die keinen menschlichen Rückruf brauchen.

### Gastkonto

Verlinken Sie direkt auf die Kontenansicht des Gasts im richtigen Tenant.

```text
https://myaccount.microsoft.com/?tenantId=<tenant-id>
```

Das ist hilfreich, wenn der Gast Kontokontext, Organisationsinformationen oder
kontobezogene Hinweise im Ressourcen-Tenant prüfen muss.

### Sicherheitsinformationen

Dieser Link kann hilfreich sein, wenn der Gast Authentifizierungsmethoden im
Ressourcen-Tenant prüfen oder registrieren muss.

```text
https://mysignins.microsoft.com/security-info?tenantId=<tenant-id>
```

Behandeln Sie dies als praktisches Deeplink-Muster und testen Sie es
regelmäßig erneut.

### Nutzungsbedingungen

Wenn Ihr Tenant Terms of Use verwendet, kann ein Direktlink frühere
Akzeptanzen leichter wieder auffindbar machen.

```text
https://myaccount.microsoft.com/termsofuse/myacceptances?tenantId=<tenant-id>
```

Behandeln Sie dies als praktisches Deeplink-Muster und testen Sie es
regelmäßig erneut.

### Gastzugang Löschen

Microsoft dokumentiert das Verlassen einer Organisation über den Bereich
**Organizations** im My-Account-Portal. Wenn Ihre Landingpage diesen Ausstieg
direkter anbieten soll, verwenden Sie einen tenant-qualifizierten Leave-
Deeplink statt nur auf die allgemeine Kontoseite zu verweisen.

```text
https://myaccount.microsoft.com/organizations/leave/<tenant-id>?tenant=<tenant-id>
```

Das zielt direkter auf denselben Leave-Ablauf. Behandeln Sie dies als
praktisches Deeplink-Muster und testen Sie es regelmäßig erneut. Falls es in
Ihrem Tenant irgendwann nicht mehr funktioniert, nutzen Sie als Fallback den
tenant-fixierten Einstieg in My Account und führen den Gast manuell zu
**Organizations** -> **Leave**.

Auf der Seite selbst können Sie diesen Link auch etwas expliziter benennen,
zum Beispiel als **Gastzugang löschen**, **Meinen Gastzugang entfernen** oder
**Diese Organisation verlassen**.

Wenn Ihre Organisation zusätzlich interne Externenkonten für dieselben Personen
führt, oft mit Mustern wie `.ext`, `vendor`, `partner` oder ähnlichem, lohnt
sich eine klare Kopplung im Lebenszyklus: Das Gastkonto ist das führende
Objekt, und wenn es deaktiviert oder gelöscht wird, wird auch das verknüpfte
interne Externenkonto bereinigt. Dann ist dieser Link nicht nur ein
Transparenz-Feature, sondern auch ein sinnvoller Einstiegspunkt in einen
geordneten Bereinigungsprozess.

## Faustregeln Für Links Und Deeplinks

Einige Beispiele auf dieser Seite sind echte Deeplinks. Andere sind einfach
tenant-fixierte URLs, die sich als verlässliche Startpunkte eignen. Dieselbe
Prüflogik gilt trotzdem für beide Arten von Links.

- Verwenden Sie `tenantId` überall dort, wo der Zieldienst es unterstützt.
- Nutzen Sie für SharePoint eine tenant-eigene URL statt einer generischen
  Microsoft-365-Startseite.
- Bevorzugen Sie Übersichtsseiten und Navigationshubs gegenüber Links, die nur
  nach bestehender Team-Mitgliedschaft funktionieren.
- Testen Sie jeden wichtigen Link, während Sie gleichzeitig im Home Tenant und
  mindestens in einem weiteren Gast-Tenant angemeldet sind.
- Prüfen Sie diese Links regelmäßig erneut. Die Teams-Muster sind dokumentiert,
  manche Account-Portal-URLs sind jedoch eher praktische Deeplink-Muster, die
  Microsoft im Lauf der Zeit ändern kann.

## Ein Einfaches Audience-Targeting-Modell

Auf derselben Landingpage müssen Gäste und Mitarbeitende nicht dieselben
Umfeldinhalte sehen.

- Zeigen Sie Gästen Sponsor-Hilfe, tenant-fixierte App-Links,
  Sicherheitsinfo-Self-Service und bei Bedarf Gast-Richtlinien.
- Zeigen Sie Mitarbeitenden interne Navigation, internen IT-Support, HR-
  Ressourcen und interne Kollaborationsziele.
- Denken Sie bei Bedarf noch feiner: Partner mit eigenem internen Externenkonto
  wie `.ext`, `vendor` oder ähnlichen Mustern haben oft andere Bedürfnisse als
  klassische Gäste, weil sie mehrere Konten parallel nutzen und häufig weiter
  auf den Geräten ihres eigenen Unternehmens arbeiten.
- Lassen Sie das Guest Sponsor Info Web Part dort stehen, wo es Mehrwert
  schafft, und nutzen Sie Quick Links darum herum, damit die gesamte Seite wie
  ein bewusst gestalteter Einstieg wirkt.

## Den Gasttyp Sichtbar Machen

Wenn Sie Audience Targeting ohnehin nutzen, kann die Landingpage dem Gast auch
direkt sagen, in welchem Rahmen er eingebunden ist: **Partner**,
**Kunde**, **Collaboration** oder bei Bedarf **Lieferant**. Das ist nicht nur
ein nettes Extra. Es gibt Kontext. Viele der restlichen Links erklären sich
dadurch fast von selbst.

Von Haus aus unterstützen in modernem SharePoint nur wenige Bausteine Audience
Targeting wirklich gut. Offiziell gehören dazu vor allem Navigation, Seiten,
News, Highlighted Content, Quick Links und Events. Für genau diesen Use Case
ist **Quick Links** oft trotzdem am geeignetsten, weil dort einzelne Links
direkt einer Zielgruppe zugeordnet werden können.

Ein nützliches Muster ist ein kleines eigenes Quick-Links-Web-Part in einem
zurückhaltenden Layout, oft **List**. Legen Sie dort pro Gasttyp einen Link auf
dieselbe Seite an, blenden Sie unnötige visuelle Elemente aus und verwenden Sie
Formulierungen wie:

- **Sie sind als Partner-Gast eingebunden.**
- **Sie sind als Kunden-Gast eingebunden.**
- **Ihr Zugriff ist als kurzfristige Collaboration angelegt.**

Solange man nicht mit der Maus darüber fährt, wird so ein Link oft eher als
normaler Hinweistext wahrgenommen und nicht als große Call-to-Action. Technisch
ist es trotzdem ein Link und damit audience-targeting-fähig. Das ist kein
offizielles Text-Web-Part-Feature, sondern ein pragmatischer Workaround.

## Gastsegmente Praktisch Nutzen

Wenn Ihr Tenant Gäste bereits grob in **Partner**, **Kunden** und
**Collaboration** einordnet, sollte diese Logik nicht im Verzeichnis stecken
bleiben. Nutzen Sie sie auch auf der Landingpage. Starten Sie dabei ruhig
einfach: eine breite Gruppe für alle Gäste und nur die Segmente, bei denen
Risiko, Standardfreigaben oder Onboarding wirklich unterschiedlich sind.
Lieferanten können dafür eine eigene Zielgruppe bekommen oder zunächst unter
**Partner** laufen.

Ein pragmatisches Muster auf der Seite kann so aussehen:

- **Alle Gäste** sehen Sponsor-Sichtbarkeit, My Account, Sicherheitsinfos,
  Leave-Self-Service und die allgemeinen tenant-fixierten Einstiegslinks.
- **Partner** sehen zusätzlich dauerhafte Projektflächen, Betriebs- oder
  Lieferdokumentation, Service- oder Ticket-Einstiege und wiederkehrende
  Kollaborations-Apps.
- **Kunden** sehen eher kuratierte Projekt- oder Support-Ziele, bewusst enger
  ausgewählte Ressourcen und sehr klare Kontaktpfade.
- **Collaboration-Gäste** sehen vor allem den einen Workshop-, Review- oder
  Dateiaustausch-Kontext plus deutliche Hinweise zu Laufzeit und nächstem
  Schritt.
- **Lieferanten** können, falls Sie sie separat führen, zusätzlich procurement-
  oder delivery-nahe Links und abgestimmte Freigabeprozesse sehen.

Der Mehrwert endet nicht auf der Landingpage. Derselbe Zuschnitt hilft oft
auch außerhalb der Seite:

- Access Packages können pro Segment passende Standardbündel aus Gruppen,
  Apps und SharePoint-Ressourcen bereitstellen.
- Anwendungen außerhalb von Microsoft Teams lassen sich für bestimmte
  Gastgruppen gruppenbasiert freischalten, statt jede Freigabe einzeln zu
  beantragen.
- Access Reviews werden aussagekräftiger, wenn nicht alle externen Konten in
  derselben Prüfung landen.
- Conditional Access und Nutzungsbedingungen können gezielter auf die
  richtigen Gasttypen wirken.
- Im Security Incident liefert die Segmentzuordnung oft einen schnellen ersten
  Kontext, auch wenn sie natürlich nur ein Indikator unter mehreren ist.

Für diesen Anwendungsfall ist in der Praxis meist eine **statische Gruppe pro
Gasttyp** die robusteste Grundlage. SharePoint-Audience-Targeting unterstützt
zwar auch Microsoft-Entra-Gruppen mit dynamischer Mitgliedschaft, aber das
hilft nur dann, wenn der Kontotyp in einem Attribut liegt, das dynamische
Regeln überhaupt auswerten können.

Je nach Architektur kann das ein Standardattribut, ein Extension-Attribut oder
eine unterstützte **Directory Extension** sein. Viele Governance-Lösungen
halten ihre fachlichen Zusatzdaten jedoch bewusst getrennt vom übrigen Schema.
Dann ist es oft einfacher, die Zielgruppen-Gruppen direkt bei der Anlage des
Gastkontos mit zu pflegen.

In **EasyLife 365 Collaboration** ist genau das ein sinnvolles Muster: Die für
Audience Targeting benötigte Gruppenzuweisung kann aus dem gewählten Template
automatisch mit gesetzt werden, während weitergehende Gast-Metadaten separat in
einer app-eigenen **Directory Extension** am Gastobjekt liegen. Das vermeidet
Konflikte im Kundentenant und lässt sich bei Bedarf sauber wieder entfernen.

<div class="doc-cta-box">
  <div>
    <p class="doc-cta-title">Die Landingpage als Gesamtsystem denken</p>
    <p class="doc-cta-sub">Sponsor-Sichtbarkeit, Gast-Self-Service und
      tenant-fixierte Einstiegspunkte funktionieren zusammen am besten.</p>
  </div>
  <div class="doc-cta-actions">
    <a href="{{ '/de/sponsor-vs-inviter/' | relative_url }}" class="btn btn-outline">Sponsor vs. Einladender</a>
    <a href="{{ '/de/setup/' | relative_url }}" class="btn btn-teal">Setup-Anleitung</a>
  </div>
</div>

## Passende Microsoft-Dokumentation

- [Target content to a specific audience on a SharePoint site](https://support.microsoft.com/office/overview-of-audience-targeting-in-modern-sharepoint-sites-68113d1b-be99-4d4c-a61c-73b087f48a81)
- [Use the Quick Links web part](https://support.microsoft.com/office/use-the-quick-links-web-part-e1df7561-209d-4362-96d4-469f85ab2a82)
- [Deep links in Microsoft Teams](https://learn.microsoft.com/microsoftteams/platform/concepts/build-and-test/deep-link-teams)
- [Planning your SharePoint hub sites](https://learn.microsoft.com/sharepoint/planning-hub-sites)
- [Manage rules for dynamic membership groups in Microsoft Entra ID](https://learn.microsoft.com/entra/identity/users/groups-dynamic-membership)
- [Change resource roles for an access package in entitlement management](https://learn.microsoft.com/entra/id-governance/entitlement-management-access-package-resources)
- [Conditional Access: Users, groups, agents, and workload identities](https://learn.microsoft.com/entra/identity/conditional-access/concept-conditional-access-users-groups)
- [Add custom data to resources by using extensions](https://learn.microsoft.com/graph/extensibility-overview)

## Passende EasyLife-Dokumentation

- [Guest Accounts Learning Guide](https://docs.easylife365.cloud/collab/getting-started/learningguides/guest-account-learning-guide):
  zeigt einen kompletten Gast-Lifecycle von Einladung über Bestätigung bis
  Deaktivierung oder Löschung.
- [Guest Accounts im Admin-Bereich](https://docs.easylife365.cloud/collab/admin/manage/guest-accounts):
  interessant für Bulk-Änderungen, Template- und Policy-Wechsel,
  Compliance-Status sowie Import und Export.
- [Templates Overview](https://docs.easylife365.cloud/collab/admin/templates):
  relevant, wenn Vorlagen nicht nur Felder und Policies, sondern auch
  Zielgruppen und Sichtbarkeit steuern sollen.
- [Confirmation Policy for Guest Accounts](https://docs.easylife365.cloud/collab/policies/confirmation-guest-accounts):
  gut passend zum Thema Sponsor- bzw. Owner-Bestätigung und Gast-Lifecycle.
