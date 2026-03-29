---
layout: doc
lang: de
title: Bereitstellungsanleitung
permalink: /de/setup/
description: >-
  Schritt-für-Schritt-Anleitung zur Bereitstellung des Guest Sponsor
  Info Web Parts und der Guest Sponsor API — SharePoint- und Azure-Setup.
lead: >-
  Erstmalige Bereitstellung und Konfigurationsreferenz für
  SharePoint- und Azure-Administratoren.
github_doc: deployment.md
---

## SharePoint-Bereitstellung

### Site Collection App Catalog aktivieren

Das Web-Part-Bundle wird in einem
[**Site Collection App Catalog**](https://learn.microsoft.com/sharepoint/dev/general-development/site-collection-app-catalog)
direkt auf der Gast-Landingpage-Site gehostet. Da Gastbenutzer bereits
Lesezugriff auf diese Site benötigen, sind keine CDN-Konfiguration oder
zusätzliche Berechtigungen für den globalen App-Catalog erforderlich.

Aktivieren Sie den Site Collection App Catalog einmalig. Für diesen
Schritt gibt es keine GUI-Option — PowerShell ist erforderlich. Das
ausführende Konto muss **alle drei** unten aufgeführten Bedingungen
erfüllen:

**Erforderliche Bedingungen:**

1. [**SharePoint-Administrator**](https://learn.microsoft.com/sharepoint/sharepoint-admin-role)-Rolle
   in Microsoft 365. Ein globaler Administrator erfüllt dies implizit.
2. **Websitesammlungsadministrator auf dem mandantenweiten App Catalog**
   (typischerweise `https://<tenant>.sharepoint.com/sites/appcatalog`).
   Falls nötig, fügen Sie Ihr Konto zuerst hinzu:

   ```powershell
   Set-SPOUser -Site "https://<tenant>.sharepoint.com/sites/appcatalog" `
       -LoginName "<admin@tenant.onmicrosoft.com>" `
       -IsSiteCollectionAdmin $true
   ```

3. **Websitesammlungsadministrator auf der Landingpage-Site** selbst.

> **Voraussetzung — der mandantenweite App Catalog muss zuerst existieren:**
> Ein App Catalog wird auf einem neuen Microsoft-365-Mandanten *nicht*
> automatisch erstellt. Falls er noch nicht vorhanden ist, öffnen Sie
> **SharePoint Admin Center → Weitere Features → Apps → Öffnen** — dies
> löst die automatische Erstellung aus.

Unter Windows ist die
[**SharePoint Online Management Shell**](https://learn.microsoft.com/powershell/sharepoint/sharepoint-online/connect-sharepoint-online)
die einfachste Option:

```powershell
Connect-SPOService -Url "https://<tenant>-admin.sharepoint.com"
Add-SPOSiteCollectionAppCatalog -Site "https://<tenant>.sharepoint.com/sites/<landing-site>"
```

Unter macOS oder Linux verwenden Sie
[PnP PowerShell](https://pnp.github.io/powershell/). Aktuelle Versionen
erfordern eine eigene Entra-App-Registrierung:

```powershell
Connect-PnPOnline -Url "https://<tenant>-admin.sharepoint.com" `
    -ClientId "<ihre-pnp-app-client-id>" -Interactive
Add-PnPSiteCollectionAppCatalog -Site "https://<tenant>.sharepoint.com/sites/<landing-site>"
```

### Hochladen und installieren

1. Laden Sie die neueste `guest-sponsor-info.sppkg` von
   [Releases](https://github.com/workoho/spfx-guest-sponsor-info/releases)
   herunter.
2. Öffnen Sie den Site Collection App Catalog und laden Sie die
   `.sppkg`-Datei hoch.

   > **Navigationstipp — verwenden Sie die direkte URL:**
   > Der Site Collection App Catalog ist eine Dokumentbibliothek namens
   > **Apps für SharePoint** in der Landingpage-Site. Am zuverlässigsten
   > finden Sie ihn unter:
   > `https://<tenant>.sharepoint.com/sites/<landing-site>/AppCatalog/`

3. Das Web Part ist sofort auf allen Seiten dieser Websitesammlung
   verfügbar — kein zusätzlicher „App hinzufügen"-Schritt erforderlich.

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

Das [Architekturdiagramm]({{ '/de/architecture/' | relative_url }})
gibt einen visuellen Überblick über alle Administratorrollen und
Bereitstellungsschritte.

### Vorbereitung: App-Registrierung erstellen

Die Azure Function verwendet
[EasyAuth](https://learn.microsoft.com/azure/app-service/overview-authentication-authorization).
EasyAuth benötigt eine Entra App-Registrierung als Identitätsanbieter.

**Option A — direkt aus dem Web ausführen**
([PowerShell 7+](https://learn.microsoft.com/powershell/scripting/install/installing-powershell)):

```powershell
& ([scriptblock]::Create((iwr 'https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/setup-app-registration.ps1'))) -TenantId "<ihre-tenant-id>"
```

**Option B — aus einem lokalen Klon:**

```powershell
./azure-function/infra/setup-app-registration.ps1 -TenantId "<ihre-tenant-id>"
```

Kopieren Sie die am Ende angezeigte **Client ID**.

### In Azure bereitstellen

Klicken Sie auf die Schaltfläche:

[![In Azure bereitstellen](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fgithub.com%2Fworkoho%2Fspfx-guest-sponsor-info%2Freleases%2Flatest%2Fdownload%2Fazuredeploy.json)

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
| Kaltstarts | ~2–5 s nach ~20 Min. Inaktivität | Eliminiert mit `alwaysReadyInstances=1` |
| Betriebssystem | Windows | Nur Linux |
| Deploy to Azure Button | Unterstützt | Unterstützt |
| Kostenschutz | `dailyMemoryTimeQuota` | `maximumFlexInstances` |
| Geschätzte Kosten | Kostenlos (innerhalb des Free Tiers) | ~€2–5/Monat mit 1 warmer Instanz |

Prüfen Sie [aka.ms/flex-region](https://aka.ms/flex-region) für die
regionale Verfügbarkeit von Flex Consumption.

### Bereitstellungsausgaben

Nach der Bereitstellung: **Ressourcengruppe → Bereitstellungen → Ausgaben**:

| Ausgabe | Verwendung |
|---|---|
| `managedIdentityObjectId` | Benötigt für `setup-graph-permissions.ps1` |
| `functionAppUrl` | Web-Part-Eigenschaftenbereich → **Guest Sponsor API Base URL** |
| `sponsorApiUrl` | Vollständige Endpunkt-URL (für Integritätsprüfungen) |

### Graph-Berechtigungen erteilen

**Option A — direkt aus dem Web ausführen:**

```powershell
& ([scriptblock]::Create((iwr 'https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/setup-graph-permissions.ps1'))) `
  -ManagedIdentityObjectId "<oid-aus-bereitstellungsausgabe>" `
  -TenantId "<ihre-tenant-id>" `
  -FunctionAppClientId "<client-id-aus-vorbereitung>"
```

**Option B — aus einem lokalen Klon:**

```powershell
./azure-function/infra/setup-graph-permissions.ps1 `
  -ManagedIdentityObjectId "<oid-aus-bereitstellungsausgabe>" `
  -TenantId "<ihre-tenant-id>" `
  -FunctionAppClientId "<client-id-aus-vorbereitung>"
```

Dieses Skript:

1. **Managed Identity Graph-Berechtigungen** — weist `User.Read.All`,
   `Presence.Read.All` (optional) und `MailboxSettings.Read` (optional) zu.
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

## Administration und Betrieb

Für den laufenden Betrieb siehe den separaten
[Betriebsleitfaden]({{ '/de/operations/' | relative_url }}):

- [Web Part aktualisieren]({{ '/de/operations/' | relative_url }}#web-part-aktualisieren)
- [Inline-Adresskarte konfigurieren]({{ '/de/operations/' | relative_url }}#inline-adresskarte-azure-maps)
- [Function aktualisieren]({{ '/de/operations/' | relative_url }}#function-aktualisieren)

Für Sicherheitsbewertung und Vertrauensannahmen siehe die
[Sicherheitsbewertung auf GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/security-assessment.md).

Für Telemetrie und Zuordnungsdetails siehe
[Telemetrie]({{ '/de/telemetry/' | relative_url }}).
