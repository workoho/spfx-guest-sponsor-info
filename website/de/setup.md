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
| 2 — Guest Sponsor API | PowerShell (`install.ps1` via `iwr`) | Azure-Mitwirkender + Besitzer + Entra-Rollen via PIM |
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

Diese Anleitung geht davon aus, dass Sie eine dedizierte
**SharePoint-Landingpage** als erstes verlässliches Ziel für Gastbenutzer
vorbereiten. Wenn Ihr Einladungsprozess oder Ihr Governance-Werkzeug eine
eigene Redirect-URL unterstützt, sollte diese auf diese Seite zeigen statt auf
ein generisches MyApps-Ziel. MyApps erklärt dem Gast nicht, wer seine
Sponsoren sind, und ein mandantenbezogener Teams-Deeplink hilft erst dann,
wenn der Gast bereits mindestens einem Team in Ihrem Mandanten hinzugefügt
wurde.

## Phase 1 — SharePoint

### Festlegen, was der Gast zuerst öffnen soll

Bevor Sie etwas installieren, legen Sie fest, welche SharePoint-Seite als
Entrance-Seite für Gäste dienen soll. Genau diese Seite sollten Sie in
Onboarding-Mails, Governance-Workflows oder Einladungs-Redirects
kommunizieren.

- Verwenden Sie eine dedizierte Landingpage und nicht die generische Startseite
  einer Kollaborationssite.
- Platzieren Sie das Web Part weit oben auf der Seite, damit Sponsor,
  Ersatz-Sponsor und Kontaktkontext sofort sichtbar sind.
- Behandeln Sie Teams-Links als nachgelagerten Schritt von dieser Seite aus,
  nicht als einziges erstes Ziel.

Wenn Sie ohnehin eine neue Landingpage aufbauen, sollten Sie außerdem prüfen,
ob sie mittel- bis langfristig auf der **Root Site** des Mandanten (`/`)
liegen sollte. Microsoft beschreibt die SharePoint Home Site als zentralen
organisatorischen Einstiegspunkt, und gerade in jüngeren Tenants ist die Root
Site häufig noch flexibel genug, um diese Entscheidung früh zu treffen. Wenn
Sie `/` verwenden, ist die Adresse für Gäste außerdem oft auch ohne zusätzlichen
Shortlink-Dienst leicht merkbar.

Das bedeutet nicht, dass Ihr Mitarbeiterportal zwingend auf derselben Seite
liegen muss. In vielen Organisationen liegen interne Inhalte bereits an
anderer Stelle, und die gemeinsame Entrance-Seite verlinkt nur dorthin. Mit
SharePoint Audience Targeting können Sie auf derselben Landingpage außerdem
unterschiedliche Navigation, News und Web-Part-Inhalte für Mitarbeitende und
Gäste einblenden.

Auch wenn die Root Site heute bereits belegt ist, kann das eine sinnvolle
langfristige Zielarchitektur sein. Sie können zunächst mit einer
Kommunikationssite wie `/sites/entrance` starten, diese als gemeinsame
Landingpage etablieren und das Erlebnis später per von Microsoft unterstütztem
Root-Site-Swap nach `/` verschieben, wenn der Zeitpunkt passt. Wenn Sie das in
Betracht ziehen, halten Sie die Landingpage als moderne Kommunikationssite und
prüfen Sie die Voraussetzungen, Berechtigungen und Freigabeeinstellungen für
die Root Site frühzeitig.

Siehe auch:

- [Landingpage-Ideen]({{ '/de/landing-page-ideas/' | relative_url }})
- [Modernize your root site](https://learn.microsoft.com/sharepoint/modern-root-site)
- [Plan, build, and launch a SharePoint home site](https://learn.microsoft.com/viva/connections/home-site-plan)

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

Aktivieren Sie das Office 365 Public CDN.

Wenn das Office 365 Public CDN aktiviert ist, repliziert SharePoint die
Web-Part-Bundles auf Microsofts Edge-CDN
(`publiccdn.sharepointonline.com`), das anonym zugänglich ist — ohne
SharePoint-Authentifizierung. Das ist der zuverlässigste Ansatz für
Gastbenutzer.

**Erforderliche Rolle:** SharePoint-Administrator.

Wählen Sie eine der folgenden gleichwertigen Admin-Shells:

<details markdown="1">
<summary>Windows: SharePoint Online Management Shell</summary>

```powershell
Connect-SPOService -Url "https://<tenant>-admin.sharepoint.com"
Set-SPOTenantCdnEnabled -CdnType Public -Enable $true

# Prüfen, ob der ClientSideAssets-Ursprung enthalten ist (wird standardmäßig hinzugefügt):
Get-SPOTenantCdnOrigins -CdnType Public
# Erwartete Ausgabe enthält: */CLIENTSIDEASSETS

# Falls der Ursprung fehlt, manuell hinzufügen:
Add-SPOTenantCdnOrigin -CdnType Public -OriginUrl "*/CLIENTSIDEASSETS"
```

</details>

<details markdown="1">
<summary>Plattformübergreifend: PowerShell 7 mit PnP PowerShell (funktioniert auch unter Windows)</summary>

```powershell
Connect-PnPOnline -Url "https://<tenant>-admin.sharepoint.com" `
  -ClientId "<your-pnp-app-client-id>" -Interactive
Set-PnPTenantCdnEnabled -CdnType Public -Enable $true

# Prüfen, ob der ClientSideAssets-Ursprung enthalten ist (wird standardmäßig hinzugefügt):
Get-PnPTenantCdnOrigin -CdnType Public
# Erwartete Ausgabe enthält: */CLIENTSIDEASSETS

# Falls der Ursprung fehlt, manuell hinzufügen:
Add-PnPTenantCdnOrigin -CdnType Public -OriginUrl "*/CLIENTSIDEASSETS"
```

</details>

> Die CDN-Propagierung dauert **bis zu 15 Minuten**. Danach ändert sich die
> Bundle-URL automatisch auf `publiccdn.sharepointonline.com` — keine
> Neukonfiguration erforderlich.

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

## Phase 2 — Guest Sponsor API

Die Guest Sponsor API ist eine begleitende Azure Function, die alle Microsoft
Graph-Aufrufe im Auftrag des Web Parts weiterleitet. Gäste authentifizieren
sich über
[EasyAuth](https://learn.microsoft.com/azure/app-service/overview-authentication-authorization),
und die Function fragt Graph über ihre eigene Managed Identity ab — Gäste
benötigen keinerlei Verzeichnisberechtigungen in Ihrem Mandanten.

Das Skript `install.ps1` ist der empfohlene Einstiegspunkt. Es lädt das
Infra-Paket herunter, startet den Bereitstellungs-Assistenten, erstellt die
Entra-App-Registrierung, stellt die gesamte Azure-Infrastruktur bereit und
weist die erforderlichen Microsoft Graph-Berechtigungen zu — gestützt auf die
[Microsoft Graph Bicep-Erweiterung](https://learn.microsoft.com/azure/templates/microsoft.graph/applications).

### Installer ausführen

<details markdown="1">
<summary>Optional: Skripte vor der Ausführung prüfen</summary>

Wenn Sie die Skripte vor der Ausführung prüfen möchten, sehen Sie sich zuerst
den
[Quelltext von install.ps1](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/azure-function/infra/install.ps1)
und den
[Quelltext von deploy-azure.ps1](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/azure-function/infra/deploy-azure.ps1)
auf GitHub an.

`install.ps1` ist ein schlanker Bootstrap-Wrapper: Das Skript lädt das
aktuelle Infra-Paket in ein temporäres Verzeichnis herunter, entpackt es,
übergibt Ihre Parameter und startet dann `deploy-azure.ps1`.

`deploy-azure.ps1` ist der eigentliche Bereitstellungs-Assistent: Er sammelt
oder übernimmt die Azure-Einstellungen, stellt die benötigten CLIs sicher,
führt die `azd`-/Bicep-Bereitstellung aus, richtet den App-Registration-Ablauf
ein und gibt danach die Werte aus, die das Web Part benötigt.

Kurz gesagt: `install.ps1` ist der empfohlene Einstiegspunkt für einen sauberen
Start, und `deploy-azure.ps1` erledigt die eigentliche Bereitstellungsarbeit,
sobald das Infra-Paket lokal verfügbar ist.

</details>

Führen Sie diesen Befehl in PowerShell 7+ aus:

```powershell
& ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.ps1').Content))
```

[Azure Developer CLI (azd)](https://aka.ms/azd) wird automatisch installiert,
falls noch nicht vorhanden. Der Installer lädt das Infra-Paket herunter,
führt durch die Auswahl von Abonnement und Ressourcengruppe, führt eine
Pre-Provision-Prüfung und die Bicep-Bereitstellung aus und gibt am Ende die
Web-Part-Konfigurationswerte aus.

### Welche Fragen der Bereitstellungs-Assistent stellt

<details markdown="1">
<summary>Optional: Prompt-für-Prompt-Referenz anzeigen</summary>

Vor den eigentlichen Parameterabfragen kann der Assistent Sie zunächst bitten,
das Azure-Abonnement auszuwählen, in dem die Bereitstellung stattfinden soll.

| Assistent fragt nach | Bedeutung | Typische Antwort |
|---|---|---|
| **azd Environment Name** | Lokaler `azd`-Umgebungsname, unter dem Einstellungen und Outputs gespeichert werden. Das ist primär eine Operator-Kennung, kein öffentlicher Azure-Ressourcenname. | Standardwert beibehalten oder einen kurzen Namen wie `contoso-gsi` verwenden. |
| **Resource Group** | Ziel-Ressourcengruppe für alle bereitgestellten Azure-Ressourcen. Der Assistent erstellt sie bei Bedarf. | `rg-<environment>` beibehalten, außer Ihre Organisation hat ein anderes Namensschema. |
| **Azure Location** | Azure-Region für Function App, Storage Account, Monitoring und die übrige Infrastruktur. | Die passende Region für Residency und Nähe wählen, z.B. `westeurope`. |
| **Environment Tag** | Optionales Governance-Tag auf Ressourcengruppe und Ressourcen. Hilft bei Inventarisierung, Policy und Kostenansicht. | `prod` ist ein sinnvoller Standard; `-` nur verwenden, wenn bewusst keine Tags gesetzt werden sollen. |
| **Criticality Tag** | Optionales Tag für die fachliche Kritikalität des Workloads. Relevant, wenn Azure-Richtlinien oder FinOps-Reports diese Tags nutzen. | `low` ist für die meisten Gast-Landingpages der sinnvolle Standard. |
| **SharePoint Tenant Name** | Kurzer Tenantname vor `.sharepoint.com`. Die Bereitstellung nutzt ihn für die korrekte SharePoint-bezogene Konfiguration. | Für `contoso.sharepoint.com` also `contoso` eingeben. |
| **Function App Name** | Öffentlicher Hostname der Azure Function App. Wenn Sie selbst einen Namen setzen, muss er global eindeutig sein. | Leer lassen, sofern Sie keinen festen Namen aus Governance- oder DNS-Gründen brauchen. |
| **Hosting Plan** | Laufzeit-/Skalierungsmodell für Azure Functions. Das beeinflusst Skalierung, Kaltstarts, Linux/Windows-Support und Kostenprofil. | Mit `Consumption` starten, außer Sie benötigen gezielt Flex-Consumption-Funktionen. |
| **Azure Maps** | Ob zusätzlich eine Azure-Maps-Ressource für eingebettete Kartenanzeige im Web Part bereitgestellt wird. | `true`, wenn Adresskarten direkt im Web Part erscheinen sollen; sonst `false`. |
| **Function Package Version** | Release-Tag des bereitzustellenden Azure-Function-Pakets. | `latest` beibehalten, außer Sie möchten bewusst eine getestete Version fixieren. |
| **Enable Monitoring** | Stellt Application Insights, Log Analytics und Alert-Ressourcen bereit. | Für Produktion `true`; nur für sehr minimale Testumgebungen deaktivieren. |
| **Graph permissions** | Legt fest, ob Graph-App-Rollen direkt im Deployment oder später per `setup-graph-permissions.ps1` zugewiesen werden. | Standardmäßig sofort zuweisen, wenn die nötige Entra-Rolle verfügbar ist; nur bei Rollentrennung oder PAW-Prozess verschieben. |

Einige Folgefragen erscheinen nur in bestimmten Fällen:

| Sichtbar wenn | Assistent fragt nach | Bedeutung | Typische Antwort |
|---|---|---|---|
| Monitoring aktiviert | **Enable Failure Anomalies alert** | Aktiviert den Application-Insights-Smart-Detector für auffällige Fehlerspitzen per E-Mail. | `false` für ruhigere Standard-Setups, `true` für proaktive Benachrichtigung ab Tag 1. |
| Hosting-Plan ist `FlexConsumption` | **Always-Ready Instances** | Anzahl vorgewärmter Instanzen. `0` erlaubt vollständiges Scale-to-zero, `1` reduziert Kaltstarts deutlich. | `1` für die meisten produktionsnahen Setups. |
| Hosting-Plan ist `FlexConsumption` | **Maximum Flex Instances** | Harte Obergrenze für parallele Flex-Instanzen. Das ist primär eine Kosten- und Burst-Control-Einstellung. | Standardwert `10` beibehalten, außer Sie brauchen bewusst ein engeres oder höheres Limit. |
| Hosting-Plan ist `FlexConsumption` | **Instance memory in MB** | Arbeitsspeicher pro Flex-Instanz. Mehr Speicher erhöht Reserve und Kosten. | `2048` beibehalten, außer Sie optimieren bewusst auf Minimalkosten. |

</details>

### Was der Installer ausführt

<details markdown="1">
<summary>Optional: Installer-Ablauf anzeigen</summary>

- **Lädt das Infra-Paket herunter** und startet den Bereitstellungs-Assistenten
- **Erstellt die Entra-App-Registrierung** —
  `Guest Sponsor Info - SharePoint Web Part Auth`
  (via [Microsoft Graph Bicep-Erweiterung](https://learn.microsoft.com/azure/templates/microsoft.graph/applications))
- **Stellt Azure-Infrastruktur bereit** — Function App, Storage Account, App Service Plan
- **Weist Microsoft Graph-Berechtigungen** der Managed Identity zu:
  `User.Read.All`, `Presence.Read.All` (optional), `MailboxSettings.Read`
  (optional), `TeamMember.Read.All` (optional)
- **Konfiguriert EasyAuth** auf der Function App mit der App-Registrierung
- **Gibt die Web-Part-Konfigurationswerte** am Ende aus

</details>

### Erforderliche Azure- und Entra-Rollen

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
> **Globaler Administrator** deckt die Entra-Anforderungen ebenfalls mit einer
> einzigen Rolle ab — die Azure-Rollen **Mitwirkender** und **Besitzer** auf
> der Ressourcengruppe werden jedoch weiterhin separat benötigt.

### Hosting-Plan-Optionen

<details markdown="1">
<summary>Optional: Consumption und Flex Consumption vergleichen</summary>

| | **Consumption** (Standard) | **Flex Consumption** |
|---|---|---|
| Kostenloser Tarif | 1 Mio. Ausführungen + 400K GB-s/Monat | 250K Ausführungen + 100K GB-s/Monat (On-Demand) |
| Kaltstarts | ~2–5 s nach ~20 Min. Inaktivität | Eliminiert mit `alwaysReadyInstances=1` |
| Betriebssystem | Windows | Nur Linux |
| Kostenschutz | `dailyMemoryTimeQuota` | `maximumFlexInstances` |
| Geschätzte Kosten | Kostenlos (innerhalb des Free Tiers) | ~2–5 €/Monat mit 1 warmer Instanz |

Prüfen Sie die [Liste der unterstützten Regionen](https://learn.microsoft.com/en-us/azure/azure-functions/flex-consumption-how-to#view-currently-supported-regions)
für die Verfügbarkeit von Flex Consumption.

</details>

### Bereitstellungsausgaben

Am Ende des Laufs gibt der Installer aus:

| Wert | Verwendung |
|---|---|
| **Guest Sponsor API Base URL** | Web-Part-Eigenschaftenbereich → **Guest Sponsor API Base URL** |
| **Web Part Client ID** | Web-Part-Eigenschaftenbereich → **Guest Sponsor API Client ID** |

Sie können diese Werte auch später abrufen mit `azd env get-values`.

## Phase 3 — Web Part konfigurieren

### Web Part auf der Landingpage einfügen

Wenn Phase 1 und Phase 2 abgeschlossen sind, öffnen Sie die SharePoint-Landingpage
im Bearbeitungsmodus und fügen Sie das Web Part **Guest Sponsor Info** der Seite hinzu.

Platzieren Sie es weit oben auf der Seite, noch vor längeren Textblöcken oder
nachgelagerten Links. Die Landingpage funktioniert am besten, wenn sie zuerst
die zwei Fragen beantwortet, die MyApps und Teams allein meist nicht klären:
wer die Sponsoren des Gasts sind und wie er sie sofort erreichen kann.

### Web Part mit der API verbinden

Solange der **Setup-Assistent** noch nicht abgeschlossen wurde, öffnet er sich
im Bearbeitungsmodus automatisch. Wenn Sie ihn bereits abgeschlossen haben,
öffnen Sie den **Eigenschaftenbereich** manuell (Zahnrad-Symbol im Bearbeitungsmodus).
Wählen Sie dann im Assistenten **Guest Sponsor API** oder tragen Sie die Werte
direkt in der Eigenschaftsgruppe **Guest Sponsor API** ein:

- **Guest Sponsor API Base URL** — die Base URL, die am Ende des
  `install.ps1`-Laufs ausgegeben wird (oder aus `azd env get-values`),
  z.B. `https://guest-sponsor-info-xyz.azurewebsites.net`
- **Guest Sponsor API Client ID** — die Web Part Client ID, die am Ende des
  `install.ps1`-Laufs ausgegeben wird (oder aus `azd env get-values`)

Der Assistent prüft das Format beider Werte vor dem Speichern. Falls er sich
nicht mehr automatisch öffnet, tragen Sie dieselben Werte direkt in der
Eigenschaftsgruppe **Guest Sponsor API** im Eigenschaftenbereich ein.

### Gastzugänglichkeit prüfen

> **Gastzugänglichkeits-Prüfung**
>
> Öffnen Sie nach dem Speichern den Eigenschaftenbereich und navigieren Sie
> zum **Gastzugänglichkeit**-Panel. Es führt eine Reihe von Prüfungen durch
> (CDN-Status, Website-Berechtigungen, externe Freigabe) und zeigt das
> Ergebnis jeder Prüfung mit einer Empfehlung an. Nutzen Sie dies, um zu
> bestätigen, dass die Voraussetzungen aus Phase 1 korrekt erfüllt sind.

## Weiterführende Informationen

Für Sicherheitsbewertung und Vertrauensannahmen siehe die
[Sicherheitsbewertung auf GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/security-assessment.md).

Für Telemetrie und Zuordnungsdetails siehe
[Telemetrie]({{ '/de/telemetry/' | relative_url }}).

Bei Problemen oder Fragen zum Betrieb siehe die [Support]({{ '/de/support/' | relative_url }})-Seite.
