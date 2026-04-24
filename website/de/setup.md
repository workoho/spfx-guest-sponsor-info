---
layout: doc
lang: de
title: Setup-Anleitung
permalink: /de/setup/
description: >-
  Schritt-für-Schritt-Anleitung zum Setup des Guest Sponsor
  Info Web Parts und der Guest Sponsor API — SharePoint- und Azure-Setup.
lead: >-
  Erstmaliges Setup und Konfigurationsreferenz für
  SharePoint- und Azure-Administratoren.
github_doc: deployment.md
---

## Übersicht

Das Setup von Guest Sponsor Info umfasst drei Phasen:

| Phase | Wo | Mindestrolle |
|---|---|---|
| 1 — SharePoint | SharePoint Admin Center + Landingpage-Site | SharePoint-Administrator |
| 2 — Guest Sponsor API | PowerShell (`deploy-azure.ps1`) | Azure-Mitwirkender + Besitzer + Entra-Rollen via PIM |
| 3 — Web Part | SharePoint-Landingpage (Bearbeitungsmodus) | Websitebesitzer |

> **Das Web Part enthält einen integrierten Setup-Assistenten**
>
> Wenn Sie das Web Part zum ersten Mal auf einer Seite platzieren, öffnet
> sich automatisch ein **Setup-Assistent**. Er führt Sie durch die Auswahl
> zwischen Produktionsmodus (Guest Sponsor API) und Demo-Modus, zeigt den
> Deploy-Befehl mit Kopier-Button an und lässt Sie am Ende
> die API-Zugangsdaten eingeben. Diese Seite ist die vollständige Referenz,
> auf die der Assistent verweist — arbeiten Sie Phasen 1 und 2 durch, bevor
> oder parallel zum Ausführen des Assistenten, und schließen Sie Phase 3
> dann direkt im Assistenten ab.

---

## Phase 1 — SharePoint

### Installation aus dem Microsoft AppSource

> **AppSource-Eintrag in Prüfung** — Das Web Part wurde beim Microsoft
> Commercial Marketplace eingereicht und wartet derzeit auf die Freigabe.
> Die folgenden Installationsschritte beschreiben den Ablauf, sobald der
> Eintrag live ist.

Das Web Part wird im
[**Microsoft Commercial Marketplace (AppSource)**](https://appsource.microsoft.com/)
verfügbar sein. Die Installation über AppSource stellt das Web Part mandantenweit
über den Tenant App Catalog bereit — kein Datei-Upload und keine manuelle
Bereitstellung erforderlich.

**Installation über das SharePoint Admin Center:**

1. Öffnen Sie **SharePoint Admin Center → Weitere Features → Apps → Öffnen**.
2. Klicken Sie auf **Apps aus Marketplace holen** und suchen Sie nach *Guest Sponsor Info*.
3. Wählen Sie die App aus und klicken Sie auf **Jetzt holen**.

Die Lösung verwendet `skipFeatureDeployment: false` — das Web Part wird **nicht**
automatisch mandantenweit verfügbar. Nach der Installation im Tenant App Catalog
muss ein Site Collection-Administrator die App explizit auf der Landingpage-Site
hinzufügen: **Websiteinhalte → App hinzufügen → Guest Sponsor Info**.
Dies ist beabsichtigt und verhindert eine versehentliche Installation auf
nicht vorgesehenen Sites.

Das Web Part fordert **keine Microsoft Graph-Berechtigungen** an — die
**API-Zugriff**-Warteschlange bleibt leer. Alle Graph-Aufrufe erfolgen
serverseitig durch die zugehörige Azure Function über ihre Managed Identity.

### Web Part für Gastbenutzer zugänglich machen

Bei der Installation über AppSource oder den Tenant App Catalog wird das
JavaScript-Bundle des Web Parts aus der `ClientSideAssets`-Bibliothek des
Tenant App Catalogs bereitgestellt. B2B-Gastbenutzer können auf diese
Bibliothek nicht zugreifen, bevor sie sich beim Host-Mandanten authentifiziert
haben — was vor dem Seitenaufruf nicht garantiert ist. Wenn Gäste das Bundle
nicht laden können, wird das Web Part lautlos nicht gerendert.

Das integrierte **Gastzugänglichkeits-Diagnose**-Panel des Web Parts
(Eigenschaftsbereich) erkennt das aktuelle Szenario und zeigt das Ergebnis
jeder Prüfung mit einer Empfehlung an.

**Option A — Office 365 Public CDN aktivieren (empfohlen)**

Wenn das Office 365 Public CDN aktiviert ist, repliziert SharePoint die
Web-Part-Bundles auf Microsofts Edge-CDN
(`publiccdn.sharepointonline.com`), das anonym zugänglich ist — ohne
SharePoint-Authentifizierung. Das ist der zuverlässigste Ansatz für
Gastbenutzer.

**Erforderliche Rolle:** SharePoint-Administrator.

```powershell
# SharePoint Online Management Shell (Windows):
Connect-SPOService -Url "https://<tenant>-admin.sharepoint.com"
Set-SPOTenantCdnEnabled -CdnType Public -Enable $true

# Prüfen, ob der ClientSideAssets-Ursprung enthalten ist (wird standardmäßig hinzugefügt):
Get-SPOTenantCdnOrigins -CdnType Public
# Erwartete Ausgabe enthält: */CLIENTSIDEASSETS
```

Falls `*/CLIENTSIDEASSETS` fehlt, manuell hinzufügen:

```powershell
Add-SPOTenantCdnOrigin -CdnType Public -OriginUrl "*/CLIENTSIDEASSETS"
```

```powershell
# PnP PowerShell (plattformübergreifend):
Connect-PnPOnline -Url "https://<tenant>-admin.sharepoint.com" `
    -ClientId "<your-pnp-app-client-id>" -Interactive
Set-PnPTenantCdnEnabled -CdnType Public -Enable $true
```

> Die CDN-Propagierung dauert **bis zu 15 Minuten**. Danach ändert sich die
> Bundle-URL automatisch auf `publiccdn.sharepointonline.com` — keine
> Neukonfiguration erforderlich.

Falls das Public CDN in Ihrer Umgebung nicht aktiviert werden kann oder Sie
außerhalb des AppSource deployen, finden Sie im
[vollständigen Deployment-Dokument auf GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/deployment.md)
alternative Optionen, darunter direkter Upload in den Tenant App Catalog
und Site Collection App Catalog-Bereitstellungen.

### Gastzugriff auf die Landingpage-Site prüfen

Gäste benötigen mindestens **Lesen** (Besucher)-Berechtigung auf der
Landingpage-Site. Verwenden Sie anstelle einer dynamischen Entra-Gruppe —
die bis zu 24 Stunden zur Aktualisierung brauchen kann — die integrierte
**Everyone**-Gruppe. Sie deckt jeden authentifizierten Benutzer ab,
einschließlich B2B-Gäste, und wirkt sofort.

Die *Everyone*-Gruppe wird durch die `ShowEveryoneClaim`-Mandanteneinstellung
gesteuert. Seit März 2018 erhalten Gastbenutzer den Everyone-Anspruch
standardmäßig nicht mehr — die Einstellung muss explizit gesetzt werden.
Falls *Everyone* nicht im Personen-Auswähler erscheint, führen Sie aus:

```powershell
# SharePoint Online Management Shell (Windows):
Set-SPOTenant -ShowEveryoneClaim $true

# PnP PowerShell (plattformübergreifend):
Set-PnPTenant -ShowEveryoneClaim $true
```

Fügen Sie dann *Everyone* zur Besuchergruppe der Site hinzu:
**Websiteeinstellungen → Personen und Gruppen → [Site] Besucher →
Neu → Benutzer hinzufügen** → suchen Sie nach *Everyone* → **Freigeben**.

> **Fallstrick — ähnlich klingende Gruppen:**
>
> - *Everyone* — schließt B2B-Gäste ein ✓
> - *Everyone except external users* — **schließt** Gäste aus ✗

### Externe Freigabe

Die mandantenweite Freigabeeinstellung ist eine **Obergrenze**: einzelne
Sites können nicht freizügiger sein als der Mandant erlaubt.

- **Aktive Websites → [Landingpage-Site] → Richtlinien → Externe
  Freigabe** — mindestens *Nur vorhandene Gäste* einstellen.

Falls diese Option ausgegraut ist, erhöhen Sie zuerst den Wert unter
**SharePoint Admin Center → Richtlinien → Freigabe** auf mindestens
*Nur vorhandene Gäste*, und konfigurieren Sie danach die Site.

---

## Phase 2 — Guest Sponsor API

Die Guest Sponsor API ist eine begleitende Azure Function, die alle Microsoft
Graph-Aufrufe im Auftrag des Web Parts weiterleitet. Gäste authentifizieren
sich über
[EasyAuth](https://learn.microsoft.com/azure/app-service/overview-authentication-authorization),
und die Function fragt Graph über ihre eigene Managed Identity ab — Gäste
benötigen keinerlei Verzeichnisberechtigungen in Ihrem Mandanten.

Das Skript `deploy-azure.ps1` übernimmt die vollständige Bereitstellung in
einem einzigen Schritt: Es erstellt die Entra-App-Registrierung, stellt die
gesamte Azure-Infrastruktur bereit und weist die erforderlichen Microsoft
Graph-Berechtigungen zu — gestützt auf die
[Microsoft Graph Bicep-Erweiterung](https://learn.microsoft.com/azure/templates/microsoft.graph/applications).

Für eingeschränkte Umgebungen (Privileged Access Workstations), in denen die
für die Bicep Graph-Erweiterung erforderlichen Entra-Verzeichnisrollen nicht
aktiviert werden können, lesen Sie
[Bereitstellen von einer PAW](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/deployment.md#deploying-from-a-privileged-access-workstation-paw)
in der Deployment-Dokumentation.

### Bereitstellen mit deploy-azure.ps1

Führen Sie aus einem lokalen Klon des Repositorys aus:

```powershell
./deploy-azure.ps1
```

[Azure Developer CLI (azd)](https://aka.ms/azd) wird automatisch installiert,
falls noch nicht vorhanden. Das Skript führt durch die Auswahl von Abonnement
und Ressourcengruppe, führt eine Pre-Provision-Prüfung und die Bicep-Bereitstellung
aus und gibt am Ende die Web-Part-Konfigurationswerte aus.

#### Was das Skript ausführt

- **Erstellt die Entra-App-Registrierung** —
  `Guest Sponsor Info - SharePoint Web Part Auth`
  (via [Microsoft Graph Bicep-Erweiterung](https://learn.microsoft.com/azure/templates/microsoft.graph/applications))
- **Stellt Azure-Infrastruktur bereit** — Function App, Storage Account, App Service Plan
- **Weist Microsoft Graph-Berechtigungen** der Managed Identity zu:
  `User.Read.All`, `Presence.Read.All` (optional), `MailboxSettings.Read`
  (optional), `TeamMember.Read.All` (optional)
- **Konfiguriert EasyAuth** auf der Function App mit der App-Registrierung
- **Gibt die Web-Part-Konfigurationswerte** am Ende aus

#### Erforderliche Azure- und Entra-Rollen

| Bereich | Erforderliche Rolle |
|---|---|
| Ressourcengruppe | **Mitwirkender** |
| Ressourcengruppe | **Besitzer** (oder Benutzerzugriffsadministrator) — für Managed Identity-Rollenzuweisungen |
| Entra ID | **Cloud-Anwendungsadministrator** — zum Erstellen und Konfigurieren der App-Registrierung |
| Entra ID | **Administrator für privilegierte Rollen** — zum Zuweisen von Graph-App-Rollen an die Managed Identity |

> **PIM-Hinweis:** Wenn Ihre Organisation
> [Privileged Identity Management (PIM)](https://learn.microsoft.com/entra/id-governance/privileged-identity-management/pim-configure)
> verwendet, aktivieren Sie die erforderlichen Entra-Rollen vor dem Ausführen
> des Skripts. Der Pre-Provision-Hook prüft Ihre aktiven Verzeichnisrollen und
> warnt, falls eine fehlt.
>
> *Alternativ* ersetzt der **Globale Administrator** beide Entra-Rollen durch
> eine einzige Rolle — die Azure-Rollen **Mitwirkender** und **Besitzer** auf
> der Ressourcengruppe werden jedoch weiterhin separat benötigt.

#### Hosting-Plan-Optionen

| | **Consumption** (Standard) | **Flex Consumption** |
|---|---|---|
| Kostenloser Tarif | 1 Mio. Ausführungen + 400K GB-s/Monat | 250K Ausführungen + 100K GB-s/Monat (On-Demand) |
| Kaltstarts | ~2–5 s nach ~20 Min. Inaktivität | Eliminiert mit `alwaysReadyInstances=1` |
| Betriebssystem | Windows | Nur Linux |
| Kostenschutz | `dailyMemoryTimeQuota` | `maximumFlexInstances` |
| Geschätzte Kosten | Kostenlos (innerhalb des Free Tiers) | ~2–5 €/Monat mit 1 warmer Instanz |

Prüfen Sie die [Liste der unterstützten Regionen](https://learn.microsoft.com/en-us/azure/azure-functions/flex-consumption-how-to#view-currently-supported-regions)
für die Verfügbarkeit von Flex Consumption.

#### Bereitstellungsausgaben

Am Ende des Laufs gibt `deploy-azure.ps1` aus:

| Wert | Verwendung |
|---|---|
| **Guest Sponsor API Base URL** | Web-Part-Eigenschaftenbereich → **Guest Sponsor API Base URL** |
| **Web Part Client ID** | Web-Part-Eigenschaftenbereich → **Guest Sponsor API Client ID** |

Sie können diese Werte auch später abrufen mit `azd env get-values`.

---

## Phase 3 — Web Part konfigurieren

Wenn Phase 1 und Phase 2 abgeschlossen sind, öffnen Sie die SharePoint-Landingpage
im Bearbeitungsmodus und fügen Sie das Web Part **Guest Sponsor Info** der Seite hinzu.

Der **Setup-Assistent** öffnet sich automatisch (er erscheint, sobald die API-URL
noch nicht konfiguriert ist). Wählen Sie **Guest Sponsor API** und gehen Sie
durch die Schritte des Assistenten bis zum **Verbinden**-Bildschirm. Geben Sie ein:

- **Guest Sponsor API Base URL** — die Base URL, die am Ende von
  `deploy-azure.ps1` ausgegeben wird (oder aus `azd env get-values`),
  z.B. `https://guest-sponsor-info-xyz.azurewebsites.net`
- **Guest Sponsor API Client ID** — die Web Part Client ID, die am Ende von
  `deploy-azure.ps1` ausgegeben wird (oder aus `azd env get-values`)

Der Assistent prüft das Format beider Werte vor dem Speichern. Sie können den
Assistenten auch überspringen und das Web Part manuell konfigurieren: Öffnen Sie
den **Eigenschaftenbereich** (Zahnrad-Symbol im Bearbeitungsmodus) und füllen
Sie die **Guest Sponsor API**-Gruppe direkt aus.

> **Gastzugänglichkeits-Prüfung**
>
> Öffnen Sie nach dem Speichern den Eigenschaftenbereich und navigieren Sie
> zum **Gastzugänglichkeit**-Panel. Es führt eine Reihe von Prüfungen durch
> (CDN-Status, Website-Berechtigungen, externe Freigabe) und zeigt das
> Ergebnis jeder Prüfung mit einer Empfehlung an. Nutzen Sie dies, um zu
> bestätigen, dass die Voraussetzungen aus Phase 1 korrekt erfüllt sind.

---

Für Sicherheitsbewertung und Vertrauensannahmen siehe die
[Sicherheitsbewertung auf GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/security-assessment.md).

Für Telemetrie und Zuordnungsdetails siehe
[Telemetrie]({{ '/de/telemetry/' | relative_url }}).

Bei Problemen oder Fragen zum Betrieb siehe die [Support]({{ '/de/support/' | relative_url }})-Seite.
