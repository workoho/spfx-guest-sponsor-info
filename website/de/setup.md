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

## SharePoint-Setup

### Installation aus dem Microsoft AppSource

Das Web Part ist im
[**Microsoft Commercial Marketplace (AppSource)**](https://appsource.microsoft.com/)
verfügbar. Die Installation über AppSource stellt das Web Part mandantenweit
über den Tenant App Catalog bereit — kein Site Collection App Catalog und kein
manueller Datei-Upload erforderlich.

**Installation über das SharePoint Admin Center:**

1. Öffne **SharePoint Admin Center → Weitere Features → Apps → Öffnen**.
2. Klicke auf **Apps aus Marketplace holen** und suche nach *Guest Sponsor Info*.
3. Wähle die App aus und klicke auf **Jetzt holen**.

Die Lösung verwendet `skipFeatureDeployment: false` — das Web Part wird **nicht**
automatisch mandantenweit verfügbar. Nach der Installation im Tenant App Catalog
muss ein Site Collection-Administrator die App explizit auf der gewünschten Site
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

> Die CDN-Propagierung dauert **15-30 Minuten**. Danach ändert sich die
> Bundle-URL automatisch auf `publiccdn.sharepointonline.com` — keine
> Neukonfiguration erforderlich.

**Option B — Jeder-Gruppe Lesezugriff auf den Tenant App Catalog gewähren**

Falls das Public CDN nicht aktiviert werden kann, der integrierten
**Jeder**-Gruppe Lesezugriff auf die Tenant App Catalog-Site gewähren.

**Erforderliche Rollen:** SharePoint-Administrator und Site Collection
Administrator der Tenant App Catalog-Site
(`https://<tenant>.sharepoint.com/sites/appcatalog`).

```powershell
# SharePoint Online Management Shell (Windows):
Connect-SPOService -Url "https://<tenant>-admin.sharepoint.com"
Add-SPOUser -Site "https://<tenant>.sharepoint.com/sites/appcatalog" `
    -LoginName "c:0(.s|true" -Group "App Catalog Visitors"
```

```powershell
# PnP PowerShell (direkte Verbindung zur App Catalog-Site):
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com/sites/appcatalog" `
    -ClientId "<your-pnp-app-client-id>" -Interactive
Add-PnPGroupMember -LoginName "c:0(.s|true" -Group "App Catalog Visitors"
```

> **Einschränkung:** Betrifft nur Gäste, die sich bereits beim Host-Mandanten
> authentifiziert haben. Das Public CDN (Option A) hat diese Einschränkung
> nicht.

Für eine erweiterte Alternative (Site Collection App Catalog, ohne
Marketplace) siehe das vollständige
[Setup-Dokumentation auf GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/deployment.md#option-c--use-a-site-collection-app-catalog).

### Gastzugriff auf die Landingpage-Site prüfen

Gäste benötigen mindestens **Lesen** (Besucher)-Berechtigung. Verwenden
Sie anstelle einer dynamischen Entra-Gruppe — die bis zu 24 Stunden zur
Aktualisierung brauchen kann — die integrierte **Everyone**-Gruppe. Sie
deckt jeden authentifizierten Benutzer ab, einschließlich B2B-Gäste, und
wirkt sofort.

Die *Everyone*-Gruppe wird durch die `ShowEveryoneClaim`-Mandanteneinstellung
gesteuert, die bei nach März 2018 erstellten Mandanten standardmäßig
`$false` ist:

```powershell
# SharePoint Online Management Shell (Windows):
(Get-SPOTenant).ShowEveryoneClaim   # aktuellen Wert prüfen
Set-SPOTenant -ShowEveryoneClaim $true

# PnP PowerShell (plattformübergreifend):
(Get-PnPTenant).ShowEveryoneClaim
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

---

## Guest Sponsor API

### Vorbereitung: App-Registrierung erstellen

Die Azure Function verwendet
[EasyAuth](https://learn.microsoft.com/azure/app-service/overview-authentication-authorization).
EasyAuth benötigt eine Entra App-Registrierung als Identitätsanbieter.

**Option A — direkt aus dem Web ausführen**
([PowerShell 7+](https://learn.microsoft.com/powershell/scripting/install/installing-powershell)):

```powershell
& ([scriptblock]::Create((iwr 'https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/setup-app-registration.ps1')))
```

**Option B — aus einem lokalen Klon:**

```powershell
./azure-function/infra/setup-app-registration.ps1
```

Kopieren Sie die am Ende angezeigte **Client ID**.

### In Azure einrichten

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
      functionClientId=<client-id-aus-vorbereitung>
```

### Erforderliche Parameter

| Parameter | Beschreibung |
|---|---|
| `tenantId` | Ihre Entra-Mandanten-ID (GUID) |
| `tenantName` | Mandantenname ohne Domain-Suffix, z.B. `contoso` |
| `functionAppName` | Global eindeutiger Name für die Function App |
| `functionClientId` | Client ID aus der Vorbereitung |
| `appVersion` | `"latest"` (Standard) oder feste SemVer ohne `v` |
| `location` | Azure-Region |

### Hosting-Plan-Optionen

| | **Consumption** (Standard) | **Flex Consumption** |
|---|---|---|
| Free Tier | 1M Ausführungen + 400K GB-s/Monat | Keins |
| Kaltstarts | ~2-5 s nach ~20 Min. Inaktivität | Eliminiert mit `alwaysReadyInstances=1` |
| Betriebssystem | Windows | Nur Linux |
| Deploy to Azure Button | Unterstützt | Unterstützt |
| Kostenschutz | `dailyMemoryTimeQuota` | `maximumFlexInstances` |
| Geschätzte Kosten | Kostenlos (innerhalb des Free Tiers) | ~€2-5/Monat mit 1 warmer Instanz |

Prüfen Sie [aka.ms/flex-region](https://aka.ms/flex-region) für die
regionale Verfügbarkeit von Flex Consumption.

### Setup-Ausgaben

Nach dem Setup: **Ressourcengruppe → Bereitstellungen → Ausgaben**:

| Ausgabe | Verwendung |
|---|---|
| `managedIdentityObjectId` | Benötigt für `setup-graph-permissions.ps1` |
| `functionAppUrl` | Web-Part-Eigenschaftenbereich → **Guest Sponsor API Base URL** |
| `sponsorApiUrl` | Vollständige Endpunkt-URL (für Integritätsprüfungen) |

### Graph-Berechtigungen erteilen

**Option A — direkt aus dem Web ausführen:**

```powershell
& ([scriptblock]::Create((iwr 'https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/setup-graph-permissions.ps1')))
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
   damit das Web Part Tokens lautlos abrufen kann.

### Web Part konfigurieren

Im Eigenschaftenbereich (**Guest Sponsor API**-Gruppe):

- **Guest Sponsor API Base URL** — z.B.
  `https://guest-sponsor-info-xyz.azurewebsites.net`
- **Guest Sponsor API Client ID (App Registration)** — die Client ID
  der App-Registrierung aus der Vorbereitung

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
