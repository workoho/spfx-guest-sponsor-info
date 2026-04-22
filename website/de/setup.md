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
| 2 — Guest Sponsor API | Azure-Portal / Cloud Shell / PowerShell | Azure-Mitwirkender + Entra-Admin |
| 3 — Web Part | SharePoint-Landingpage (Bearbeitungsmodus) | Websitebesitzer |

> **Das Web Part enthält einen integrierten Setup-Assistenten**
>
> Wenn Sie das Web Part zum ersten Mal auf einer Seite platzieren, öffnet
> sich automatisch ein **Setup-Assistent**. Er führt Sie durch die Auswahl
> zwischen Produktionsmodus (Guest Sponsor API) und Demo-Modus, zeigt die
> Azure-Setup-Befehle inline mit Kopier-Buttons an und lässt Sie am Ende
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

Der Setup-Assistent zeigt diese drei Schritte inline mit kopierbaren Befehlen.
Die folgenden Abschnitte sind die vollständige Referenz für jeden Schritt.

### Schritt 1: App-Registrierung erstellen

EasyAuth benötigt eine Entra App-Registrierung als Identitätsanbieter für
die Azure Function.

**Option A — direkt aus dem Web ausführen**
([PowerShell 7+](https://learn.microsoft.com/powershell/scripting/install/installing-powershell)):

```powershell
& ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/setup-app-registration.ps1').Content))
```

**Option B — aus einem lokalen Klon:**

```powershell
./azure-function/infra/setup-app-registration.ps1
```

Kopieren Sie die am Ende angezeigte **Client ID** — Sie benötigen sie in
Schritt 3 und bei der Konfiguration des Web Parts.

### Schritt 2: In Azure bereitstellen

Klicken Sie auf die Schaltfläche:

[![In Azure bereitstellen](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2Fworkoho%2Fspfx-guest-sponsor-info%2Fmain%2Fazure-function%2Finfra%2Fazuredeploy.json)

Oder aus der [Azure Cloud Shell](https://shell.azure.com):

```bash
az deployment group create \
  --resource-group <ihre-ressourcengruppe> \
  --template-uri https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/azuredeploy.json \
  --parameters \
      tenantId=<ihre-tenant-id> \
      tenantName=<ihr-tenant-name> \
      functionAppName=<global-eindeutiger-name> \
      webPartClientId=<client-id-aus-schritt-1>
```

<details>
<summary>Optional: als Deployment Stack bereitstellen</summary>

[Deployment Stacks](https://learn.microsoft.com/azure/azure-resource-manager/bicep/deployment-stacks)
verwalten alle Ressourcen als verwaltetes Set. Das vollständige Entfernen
erfordert nur einen einzigen Befehl.

```bash
az stack group create \
  --name guest-sponsor-info \
  --resource-group <ihre-ressourcengruppe> \
  --template-uri https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/azuredeploy.json \
  --parameters \
      tenantId=<ihre-tenant-id> \
      tenantName=<ihr-tenant-name> \
      functionAppName=<global-eindeutiger-name> \
      webPartClientId=<client-id-aus-schritt-1> \
  --action-on-unmanage deleteResources \
  --deny-settings-mode none
```

</details>

#### Erforderliche Parameter

| Parameter | Beschreibung |
|---|---|
| `tenantId` | Ihre Entra-Mandanten-ID (GUID) |
| `tenantName` | Mandantenname ohne Domain-Suffix, z.B. `contoso` |
| `functionAppName` | Global eindeutiger Name für die Function App |
| `webPartClientId` | Client ID aus Schritt 1 |
| `appVersion` | `"latest"` (Standard) oder feste SemVer ohne `v` |
| `location` | Azure-Region |

#### Hosting-Plan-Optionen

| | **Consumption** (Standard) | **Flex Consumption** |
|---|---|---|
| Free Tier | 1M Ausführungen + 400K GB-s/Monat | 250K Ausführungen + 100K GB-s/Monat (On-Demand) |
| Kaltstarts | ~2-5 s nach ~20 Min. Inaktivität | Eliminiert mit `alwaysReadyInstances=1` |
| Betriebssystem | Windows | Nur Linux |
| Deploy to Azure Button | Unterstützt | Unterstützt |
| Kostenschutz | `dailyMemoryTimeQuota` | `maximumFlexInstances` |
| Geschätzte Kosten | Kostenlos (innerhalb des Free Tiers) | ~€2-5/Monat mit 1 warmer Instanz |

Prüfen Sie die [Liste der unterstützten Regionen](https://learn.microsoft.com/en-us/azure/azure-functions/flex-consumption-how-to#view-currently-supported-regions)
für die Verfügbarkeit von Flex Consumption. Zusätzliche Parameter für Flex:

```bash
az deployment group create \
  --resource-group <ihre-ressourcengruppe> \
  --template-uri https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/azuredeploy.json \
  --parameters \
      tenantId=<ihre-tenant-id> \
      tenantName=<ihr-tenant-name> \
      functionAppName=<global-eindeutiger-name> \
      webPartClientId=<client-id-aus-schritt-1> \
      hostingPlan=FlexConsumption \
      maximumFlexInstances=10
```

#### Bereitstellungsausgaben

Nach der Bereitstellung: **Ressourcengruppe → Bereitstellungen → Ausgaben**:

| Ausgabe | Verwendung |
|---|---|
| `managedIdentityObjectId` | Benötigt für Schritt 3 |
| `functionAppUrl` | Web-Part-Eigenschaftenbereich → **Guest Sponsor API Base URL** |
| `sponsorApiUrl` | Vollständige Endpunkt-URL (für Integritätsprüfungen) |

### Schritt 3: Graph-Berechtigungen erteilen

**Option A — direkt aus dem Web ausführen:**

```powershell
& ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/setup-graph-permissions.ps1').Content))
```

**Option B — aus einem lokalen Klon:**

```powershell
./azure-function/infra/setup-graph-permissions.ps1
```

Dieses Skript:

1. **Managed Identity Graph-Berechtigungen** — weist
  [`User.Read.All`](https://learn.microsoft.com/en-us/graph/permissions-reference#userreadall),
  [`Presence.Read.All`](https://learn.microsoft.com/en-us/graph/permissions-reference#presencereadall)
  (optional),
  [`MailboxSettings.Read`](https://learn.microsoft.com/en-us/graph/permissions-reference#mailboxsettingsread)
  (optional) und
  [`TeamMember.Read.All`](https://learn.microsoft.com/en-us/graph/permissions-reference#teammemberreadall)
  (optional) zu.
2. **App-Registrierung einrichten** — stellt einen `user_impersonation`-Scope
   bereit und autorisiert *SharePoint Online Web Client Extensibility* vorab,
   damit das Web Part Token lautlos abrufen kann.

---

## Phase 3 — Web Part konfigurieren

Wenn Phase 1 und Phase 2 abgeschlossen sind, öffnen Sie die SharePoint-Landingpage
im Bearbeitungsmodus und fügen Sie das Web Part **Guest Sponsor Info** der Seite hinzu.

Der **Setup-Assistent** öffnet sich automatisch (er erscheint, sobald die API-URL
noch nicht konfiguriert ist). Wählen Sie **Guest Sponsor API** und gehen Sie
durch die Schritte des Assistenten bis zum **Verbinden**-Bildschirm. Geben Sie ein:

- **Guest Sponsor API Base URL** — die `functionAppUrl` aus den
  Bereitstellungsausgaben (Phase 2, Schritt 2),
  z.B. `https://guest-sponsor-info-xyz.azurewebsites.net`
- **Guest Sponsor API Client ID** — die Client ID aus Phase 2, Schritt 1

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

## Administration und Operations

Für den laufenden Betrieb siehe den separaten
[Operations Guide]({{ '/de/operations/' | relative_url }}):

- [Web Part aktualisieren]({{ '/de/operations/' | relative_url }}#web-part-aktualisieren)
- [Inline-Adresskarte konfigurieren]({{ '/de/operations/' | relative_url }}#inline-adresskarte-azure-maps)
- [Function aktualisieren]({{ '/de/operations/' | relative_url }}#function-aktualisieren)

Für Sicherheitsbewertung und Vertrauensannahmen siehe die
[Sicherheitsbewertung auf GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/security-assessment.md).

Für Telemetrie und Zuordnungsdetails siehe
[Telemetrie]({{ '/de/telemetry/' | relative_url }}).

Bei Problemen oder Fragen zum Betrieb siehe die [Support]({{ '/de/support/' | relative_url }})-Seite.
