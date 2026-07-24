---
layout: doc
lang: de
title: Setup-Anleitung
permalink: /de/setup/
description: >-
  Schritt-für-Schritt-Anleitung für eine SharePoint-Gast-Landingpage mit
  Guest Sponsor Info und der Guest Sponsor API — Azure-Setup,
  SharePoint-Gastzugriff und Sponsor-Sichtbarkeit.
lead: >-
  Umsetzungsleitfaden für SharePoint- und Azure-Administratoren, die ein
  saubereres Gäste-Onboarding, verlässlichen SharePoint-Gastzugriff und
  sichtbare Sponsoren auf der Landingpage wollen.
github_doc: deployment.md
---

## Übersicht

Das Setup von Guest Sponsor Info hat drei Phasen:

| Phase | Wo | Mindestrolle |
|---|---|---|
| 1 — SharePoint | SharePoint Admin Center + Landingpage-Site | SharePoint-Administrator |
| 2 — Guest Sponsor API | [Azure Cloud Shell](https://shell.azure.com/) (empfohlen) oder lokale PowerShell/Shell | Azure-Besitzer oder Azure-Mitwirkender + eine Zugriffsverwaltungsrolle (Benutzerzugriffsadministrator oder Role Based Access Control Administrator), plus Microsoft Entra-Rollen via PIM. Ressourcengruppen-Scope ist der Normalfall; Subscription-Scope wird nur benötigt, wenn dieser Lauf noch Provider registrieren oder die Ressourcengruppe anlegen muss. |
| 3 — Web Part | SharePoint-Landingpage (Bearbeitungsmodus) | Websitebesitzer |

> [!NOTE]
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

<figure class="doc-screenshot">
  <img src="{{ '/assets/images/my-sponsors-setup-wizard.jpg' | relative_url }}"
    alt="Schritt im Setup-Assistenten mit der Auswahl zwischen Guest Sponsor API und Demo-Modus."
    class="doc-screenshot__img" data-zoomable>
  <figcaption class="doc-screenshot__caption">
    Der integrierte Setup-Assistent lässt Seitenbearbeiter direkt mit der
    Guest Sponsor API starten oder das Web Part zunächst im Demo-Modus testen.
  </figcaption>
</figure>

### Bevor Sie beginnen

Diese Anleitung geht von einer dedizierten **SharePoint-Landingpage** als
erstem verlässlichen Ziel für Gastbenutzer aus. Wenn Ihr Einladungsprozess
oder Ihr Governance-Werkzeug eine
eigene Redirect-URL unterstützt, sollte diese auf diese Seite zeigen statt auf
ein generisches My-Apps-Ziel. My Apps ist für App-Start gedacht, nicht für
Sponsor-Sichtbarkeit, und ein mandantenbezogener Microsoft Teams-Deeplink hilft erst
dann, wenn der Gast bereits mindestens einem Team in Ihrem Mandanten
hinzugefügt wurde.

Außerdem sollten Sie die Begriffe früh sauber trennen: Sponsor und
Einladender sind in Gäste-Onboarding-Prozessen nicht immer dieselbe Person.
Manche Tools bezeichnen den Sponsor zusätzlich als „Owner“ der
Gastbeziehung. Wenn Landingpage, E-Mails oder Admin-Anweisungen diese Rollen
vermischen, landet der Gast leicht beim falschen Ansprechpartner.

[Vollständige Erklärung zu Sponsor vs. Einladender]({{ '/de/sponsor-vs-inviter/' | relative_url }}).

Wenn Sie für Gastkonten noch gar keine Governance- und Lifecycle-Schicht
haben, sollten Sie diese zuerst lösen. Das Web Part entfaltet seinen Nutzen
erst dann richtig, wenn Gäste zuverlässig auf Ihrer Landingpage ankommen und
die Sponsor-Zuordnung dauerhaft aktuell bleibt.

Im standardmäßigen indirekten Einladungsfluss haben Sie in der Regel keinen
Einfluss darauf, wohin Gäste nach dem Annehmen der Einladung weitergeleitet
werden. Dadurch landen sie oft in My Apps statt auf Ihrer Landingpage. Damit
die Weiterleitung zuverlässig auf die Landingpage zeigt, brauchen Sie einen
Prozess auf Basis von Microsoft Graph-Einladungen oder ein Governance-Werkzeug,
das genau darauf aufsetzt. EasyLife 365 Collaboration tut genau das. Man kann
so etwas grundsätzlich auch selbst bauen, müsste dann aber zusätzlich die
eigentlichen Governance-, Sponsor-Zuordnungs- und Lifecycle-Funktionen
nachbauen. In der Praxis ist das meist teurer als eine EasyLife-Lizenz.

<div class="doc-easylife-box">
  <div class="easylife-card easylife-card--doc">
    <a href="{{ site.author_url }}?utm_source=guest-sponsor-info&amp;utm_medium=website&amp;utm_content=easylife-doc-workoho-setup-de"
      target="_blank" rel="noopener" class="easylife-doc-workoho-link no-external-icon">
      <img src="{{ '/assets/images/workoho-logo.svg' | relative_url }}" alt="Workoho"
        class="easylife-doc-workoho-logo" width="110" height="27">
    </a>
    <p class="easylife-doc-partner-note">
      Workoho, das Team hinter Guest Sponsor Info, ist EasyLife 365
      <span class="easylife-tier easylife-tier--inline">Platinum Partner</span>.
    </p>
    <a href="https://easylife365.cloud/products/collaboration/?utm_source=guest-sponsor-info&amp;utm_medium=website&amp;utm_content=easylife-setup-logo-de"
      target="_blank" rel="noopener" class="easylife-doc-easylife-link no-external-icon">
      <img src="{{ '/assets/images/easylife365-logo.svg' | relative_url }}" alt="EasyLife 365 Collaboration"
        class="easylife-logo" width="240" height="33">
    </a>
    <p class="easylife-doc-copy">
      Buchen Sie eine Demo mit Workoho und sehen Sie, wie EasyLife Gäste auf
      die Landingpage bringt und Governance stabil hält.
    </p>
    <a href="https://wkho.io/easylife365-demo?utm_source=guest-sponsor-info&amp;utm_medium=website&amp;utm_content=easylife-setup-de"
      target="_blank" rel="noopener" class="easylife-cta">Demo buchen</a>
  </div>
</div>

Für Microsoft Graph-Berechtigungen und Laufzeit-Datenverarbeitung siehe
[Datenschutz](/de/privacy/). Für Azure-Nutzungszuordnung und Opt-out siehe
[Telemetrie](/de/telemetry/). Wenn Sie statt eines Self-Service-Rollouts
direkte Unterstützung brauchen, siehe [Support](/de/support/).

## Phase 1 — SharePoint

### Festlegen, was der Gast zuerst öffnen soll

Bevor Sie etwas installieren, legen Sie fest, welche SharePoint-Seite als
Landingpage für Gäste dienen soll. Genau diese Seite sollte in
Onboarding-Mails, Governance-Workflows und Einladungs-Redirects auftauchen.
Sie sollte nach der Annahme der Einladung der erste verlässliche
SharePoint-Gastzugriffspunkt sein.

- Verwenden Sie eine dedizierte Landingpage und nicht die generische Startseite
  einer Kollaborationssite.
- Platzieren Sie das Web Part weit oben auf der Seite, damit Sponsor,
  Ersatz-Sponsor und Kontaktkontext sofort sichtbar sind.
- Behandeln Sie Links zu Microsoft Teams als nachgelagerten Schritt von dieser Seite aus,
  nicht als einziges erstes Ziel.

### Festlegen, wo die Landingpage liegen soll

Wenn Sie ohnehin eine neue Landingpage aufbauen, sollten Sie außerdem prüfen,
ob sie mittel- bis langfristig auf der **Root Site** des Mandanten (`/`)
liegen sollte. Microsoft beschreibt die SharePoint Home Site als zentralen
organisatorischen Einstiegspunkt, und gerade in jüngeren Tenants ist die Root
Site oft noch flexibel genug, um diese Entscheidung früh zu treffen. Wenn Sie
`/` verwenden, ist die Adresse für Gäste außerdem oft auch ohne zusätzlichen
Shortlink-Dienst leicht merkbar.

Das bedeutet nicht, dass Ihr Mitarbeiterportal zwingend auf derselben Seite
liegen muss. In vielen Organisationen liegen interne Inhalte bereits an
anderer Stelle, und die gemeinsame Landingpage verlinkt nur dorthin. Mit
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

> [!IMPORTANT]
> **AppSource-Eintrag in Prüfung** — Das Web Part wurde beim Microsoft
> Commercial Marketplace eingereicht und wartet derzeit auf die Freigabe.
> Die folgenden Installationsschritte beschreiben den Ablauf, sobald der
> Eintrag live ist. Wenn Sie vorher bereitstellen müssen, nutzen Sie den
> [Deployment Guide auf GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/deployment.md)
> als Alternative ohne AppSource.

Das Web Part ist im
[**Microsoft Commercial Marketplace (AppSource)**](https://appsource.microsoft.com/)
verfügbar. Die Installation über AppSource stellt das Web Part mandantenweit
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

PowerShell-Voraussetzungen für die folgenden Befehle:

- **Windows / SharePoint Online Management Shell:**
  `Microsoft.Online.SharePoint.PowerShell` einmal installieren.
- Für die Modulinstallation bevorzugt `Install-PSResource` verwenden. Unter
  Windows PowerShell 5.1 zuerst
  [PowerShellGet / PSResourceGet aktualisieren](https://learn.microsoft.com/powershell/gallery/powershellget/install-powershellget?view=powershellget-3.x),
  weil `Install-PSResource` dort nicht ab Werk verfügbar ist.
- **PnP-Pfad:** auch unter Windows **PowerShell 7+** verwenden,
  [PnP PowerShell](https://pnp.github.io/powershell/) einmal installieren und
  einmalig eine eigene
  [Entra-App registrieren](https://pnp.github.io/powershell/articles/registerapplication.html),
  weil `Connect-PnPOnline -Interactive` eine Client-ID benötigt.

Wählen Sie eine der folgenden gleichwertigen Admin-Shells:

<details markdown="1">
<summary>Windows: SharePoint Online Management Shell</summary>

```powershell
# Einmal installieren: Install-PSResource Microsoft.Online.SharePoint.PowerShell -Repository PSGallery -Scope CurrentUser
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
# Einmal installieren (PowerShell 7+): Install-PSResource PnP.PowerShell -Repository PSGallery -Scope CurrentUser
# Einmal registrieren: https://pnp.github.io/powershell/articles/registerapplication.html
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

> [!NOTE]
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

PowerShell-Voraussetzungen für die folgenden Befehle:

- **Windows / SharePoint Online Management Shell:**
  `Microsoft.Online.SharePoint.PowerShell` einmal installieren.
- Für die Modulinstallation bevorzugt `Install-PSResource` verwenden. Unter
  Windows PowerShell 5.1 zuerst
  [PowerShellGet / PSResourceGet aktualisieren](https://learn.microsoft.com/powershell/gallery/powershellget/install-powershellget?view=powershellget-3.x),
  weil `Install-PSResource` dort nicht ab Werk verfügbar ist.
- **PnP-Pfad:** auch unter Windows **PowerShell 7+** verwenden,
  [PnP PowerShell](https://pnp.github.io/powershell/) einmal installieren und
  einmalig eine eigene
  [Entra-App registrieren](https://pnp.github.io/powershell/articles/registerapplication.html),
  weil `Connect-PnPOnline -Interactive` eine Client-ID benötigt.

Wählen Sie eine der folgenden gleichwertigen Admin-Shells:

<details markdown="1">
<summary>Windows: SharePoint Online Management Shell</summary>

```powershell
# Einmal installieren: Install-PSResource Microsoft.Online.SharePoint.PowerShell -Repository PSGallery -Scope CurrentUser
Set-SPOTenant -ShowEveryoneClaim $true
```

</details>

<details markdown="1">
<summary>Plattformübergreifend: PowerShell 7 mit PnP PowerShell (funktioniert auch unter Windows)</summary>

```powershell
# Einmal installieren (PowerShell 7+): Install-PSResource PnP.PowerShell -Repository PSGallery -Scope CurrentUser
# Einmal registrieren: https://pnp.github.io/powershell/articles/registerapplication.html
Connect-PnPOnline -Url "https://<tenant>-admin.sharepoint.com" `
  -ClientId "<your-pnp-app-client-id>" -Interactive
Set-PnPTenant -ShowEveryoneClaim $true
```

</details>

Fügen Sie dann *Everyone* zur Besuchergruppe der Site hinzu:
**Websiteeinstellungen → Personen und Gruppen → [Site] Besucher →
Neu → Benutzer hinzufügen** → suchen Sie nach *Everyone* → **Freigeben**.

> [!WARNING]
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

Die Guest Sponsor API ist eine begleitende Azure Function, die Microsoft-
Graph-Aufrufe für das Web Part weiterleitet. Gäste authentifizieren
sich über
[EasyAuth](https://learn.microsoft.com/azure/app-service/overview-authentication-authorization),
und die Function fragt Graph über ihre eigene Managed Identity ab — Gäste
benötigen keinerlei Verzeichnisberechtigungen in Ihrem Mandanten.

Verwenden Sie `install.ps1` als Standard-Einstiegspunkt. Es lädt das
Infra-Paket herunter, startet den Bereitstellungs-Assistenten, erstellt die
Entra-App-Registrierung, stellt die Azure-Infrastruktur bereit und weist die
erforderlichen Microsoft Graph-Berechtigungen zu — gestützt auf die
[Microsoft Graph Bicep-Erweiterung](https://learn.microsoft.com/graph/templates/bicep/overview-bicep-templates-for-graph).

### Installer ausführen

> [!TIP]
> **[Azure Cloud Shell](https://shell.azure.com/) ist für die meisten Administratoren der einfachste Weg.**
>
> Öffnen Sie [Azure Cloud Shell](https://shell.azure.com/), wählen Sie
> **PowerShell**, und führen Sie den Installer direkt dort aus. Auf Ihrer
> lokalen Maschine müssen Sie dafür nichts vorbereiten: keine lokale
> PowerShell-Installation, kein Azure CLI, kein `azd`, kein `curl` und kein
> Homebrew.
>
> Der Bereitstellungs-Assistent erkennt
> [Azure Cloud Shell](https://shell.azure.com/) automatisch und kann `azd` bei
> Bedarf im Cloud-Shell-Home-Verzeichnis selbst installieren.
>
> Einen kurzen Überblick über die aktuelle Ersteinrichtung zeigt
> [Get started with Azure Cloud Shell](https://learn.microsoft.com/en-us/azure/cloud-shell/get-started/ephemeral).

**Empfohlener Standardpfad: [Azure Cloud Shell](https://shell.azure.com/) (PowerShell)**

1. Öffnen Sie [Azure Cloud Shell](https://shell.azure.com/).
2. Wählen Sie **PowerShell**.
3. Führen Sie diesen Befehl aus:

```powershell
& ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.ps1').Content))
```

Damit läuft der gesamte Azure-Teil im Browserfenster. Ihre lokale Maschine
bleibt unangetastet.

> [!IMPORTANT]
> **In Cloud Shell ist eine Anmeldung per Gerätecode erforderlich.**
>
> Der Assistent meldet sich mit dem Azure CLI an, statt die
> Cloud-Shell-Identität weiterzuverwenden — diese darf die
> Entra-App-Registrierung nicht anlegen. Sie erhalten dafür einen Code und eine
> URL, die Sie in einem Browser bestätigen. Prüfen Sie vorher zwei Punkte:
>
> - **Die Anmeldung per Gerätecode muss für Ihr Konto erlaubt sein.** Manche
>   Organisationen sperren dieses Verfahren per Conditional Access. Ist das bei
>   Ihnen der Fall, führen Sie den Installer stattdessen in einer lokalen
>   PowerShell aus — dort erfolgt eine normale Browser-Anmeldung — oder lassen
>   Sie das Verfahren von Ihrer Entra-Administration für das
>   Bereitstellungskonto freigeben.
> - **Admin-Konten mit PAW-Bindung bestätigen den Code auf der PAW.** Darf sich
>   Ihr Administratorkonto nur von einer Privileged Access Workstation aus
>   anmelden, öffnen Sie die Gerätecode-URL auf genau diesem Gerät. Eine
>   Bestätigung von einem anderen Gerät wird abgewiesen — unabhängig davon, wo
>   Cloud Shell geöffnet ist.

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

Wenn Sie stattdessen lokal in PowerShell 7+ arbeiten möchten, führen Sie
diesen Befehl aus:

```powershell
& ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.ps1').Content))
```

Unter macOS oder Linux können Sie alternativ aus einer normalen Shell starten.
Dieser Bootstrapper installiert PowerShell bei Bedarf und führt dann
denselben Installer aus:

```bash
curl -fsSL https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.sh | bash
```

[Azure Developer CLI (azd)](https://aka.ms/azd) wird automatisch installiert,
falls noch nicht vorhanden. Der Installer lädt das Infra-Paket herunter,
führt durch die Auswahl von Abonnement und Ressourcengruppe, führt eine
Pre-Provision-Prüfung und die Bicep-Bereitstellung aus und gibt am Ende die
Web-Part-Konfigurationswerte aus.

Im Assistenten wählen Sie Azure-Abonnement, Ressourcengruppe, Region und den
SharePoint-Tenant-Namen. Verwenden Sie ein Abonnement im selben Entra-Tenant
wie Ihr SharePoint-Tenant. Lassen Sie den Function-App-Namen leer, sofern Sie
keinen festen Namen benötigen; Azure erzeugt dann einen sicheren eindeutigen
Namen. Die Standardwerte passen für die meisten Bereitstellungen, und die
Graph-Berechtigungszuweisung kann aktiviert bleiben, wenn Ihre Entra-Rollen
aktiv sind.

### Ablauf des Installers

<details markdown="1">
<summary>Optional: Installer-Ablauf anzeigen</summary>

- **Lädt das Infra-Paket herunter** und startet den Bereitstellungs-Assistenten
- **Erstellt die Entra-App-Registrierung** —
  `Guest Sponsor Info - SharePoint Web Part Auth`
  (via [Microsoft Graph Bicep-Erweiterung](https://learn.microsoft.com/graph/templates/bicep/overview-bicep-templates-for-graph))
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
| Subscription (nur wenn dieser Lauf noch Provider registrieren oder die Ressourcengruppe anlegen muss) | **Besitzer** für einen Ein-Rollen-Pfad oder **Mitwirkender** + eine Zugriffsverwaltungsrolle (**Benutzerzugriffsadministrator** oder **Role Based Access Control Administrator**) für einen aufgeteilten Least-Privilege-Pfad |
| Ressourcengruppe (normaler Regelbetrieb, einschließlich erstem Deployment in eine bereits vorhandene Ressourcengruppe, nachdem die Provider schon registriert sind) | **Besitzer** für einen Ein-Rollen-Pfad oder **Mitwirkender** + eine Zugriffsverwaltungsrolle (**Benutzerzugriffsadministrator** oder **Role Based Access Control Administrator**) |
| Entra ID | **Cloud-Anwendungsadministrator** — zum Erstellen und Konfigurieren der App-Registrierung |
| Entra ID | **Administrator für privilegierte Rollen** — zum Zuweisen von Graph-App-Rollen an die Managed Identity |

> [!TIP]
> Der entscheidende Unterschied ist nicht „erstes Deployment“ versus
> „Update“, sondern ob dieser Lauf noch **Bootstrap-Aktionen auf
> Subscription-Ebene** enthält. Wenn noch Provider registriert oder die
> Ressourcengruppe angelegt werden müssen, verwenden Sie Azure-Rechte auf der
> **Subscription**: entweder **Besitzer** allein oder
> **Mitwirkender** + eine Zugriffsverwaltungsrolle
> (**Benutzerzugriffsadministrator** oder **Role Based Access Control
> Administrator**). Mitwirkender deckt die Provider-Registrierung und bei
> Bedarf die Anlage der Ressourcengruppe ab, die Zugriffsverwaltungsrolle
> deckt die Rollenzuweisungen auf dem Storage Account ab.
>
> Sobald die Provider registriert sind und die Ressourcengruppe bereits
> existiert, ist der Normalfall **Besitzer** allein auf
> Ressourcengruppen-Ebene oder **Mitwirkender** + eine
> Zugriffsverwaltungsrolle (**Benutzerzugriffsadministrator** oder
> **Role Based Access Control Administrator**) auf
> Ressourcengruppen-Ebene. Dieser
> engere Scope reicht sowohl für normale Updates als auch für ein erstes
> Deployment in eine bereits vorbereitete Ressourcengruppe.
> **Mitwirkender** allein reicht nicht, weil das Deployment-Template die
> `Microsoft.Authorization/roleAssignments` für die Managed Identity der
> Function App weiterhin im Sollzustand hält und ARM bei inkrementellen
> Deployments Ressourcen, die im Template verbleiben, erneut auswertet. Wenn
> ein späteres Update erstmals einen bisher nicht registrierten Provider
> aktiviert, braucht genau dieser Lauf wieder Subscription-weiten
> **Mitwirkender** (oder **Besitzer**) für die Provider-Registrierung.
>
> **PIM-Hinweis:** Wenn Ihre Organisation
> [Privileged Identity Management (PIM)](https://learn.microsoft.com/entra/id-governance/privileged-identity-management/pim-configure)
> verwendet, aktivieren Sie die erforderlichen Entra-Rollen vor dem Ausführen
> des Skripts. Der Pre-Provision-Hook prüft Ihre aktiven Verzeichnisrollen und
> warnt, falls eine fehlt.
>
> **Globaler Administrator** deckt die Entra-Anforderungen ebenfalls mit einer
> einzigen Rolle ab — die Azure-Rollen müssen weiterhin separat vergeben
> werden.

### Bereitstellungsausgaben

Am Ende gibt der Installer aus:

| Wert | Verwendung |
|---|---|
| **Guest Sponsor API Base URL** | Web-Part-Eigenschaftenbereich → **Guest Sponsor API Base URL** |
| **Web Part Client ID** | Web-Part-Eigenschaftenbereich → **Guest Sponsor API Client ID** |

Sie können diese Werte auch später abrufen mit `azd env get-values`.

## Phase 3 — Web Part konfigurieren

### Web Part auf der Landingpage einfügen

Wenn Phase 1 und Phase 2 abgeschlossen sind, öffnen Sie die
SharePoint-Landingpage im Bearbeitungsmodus und fügen Sie das Web Part
**Guest Sponsor Info** der Seite hinzu.

Platzieren Sie es weit oben auf der Seite, noch vor längeren Textblöcken oder
nachgelagerten Links. Die Landingpage funktioniert am besten, wenn sie zuerst
die zwei Fragen beantwortet, die MyApps und Teams allein meist nicht klären:
wer die Sponsoren des Gasts sind und wie er sie sofort erreichen kann.

Im Bearbeitungsmodus kann das Web Part bereits realistische Demo-Sponsoren
anzeigen. So lassen sich Platzierung, Abstände und geöffnete Karten prüfen,
bevor die Azure-Seite vollständig verbunden ist. Wenn Sie im
Eigenschaftenbereich zusätzlich den öffentlichen Demo-Modus aktivieren,
zeigt die Live-Seite simulierte Sponsoren auch für normale Besucher,
einschließlich interner Benutzer und Gäste. So lässt sich das Web Part ohne
Bearbeitungsmodus und ohne Gastkonto ausprobieren.

<figure class="doc-screenshot">
  <img src="{{ '/assets/images/my-sponsors-editor-preview.jpg' | relative_url }}"
    alt="SharePoint-Bearbeitungsmodus mit dem Web Part, Demo-Sponsoren und geöffneter Sponsorkarte."
    class="doc-screenshot__img" data-zoomable>
  <figcaption class="doc-screenshot__caption">
    Der Vorschaumodus für Editoren hilft dabei, das Web Part auf der Seite zu
    platzieren und die Kartenansicht zu prüfen, bevor sich echte Gäste anmelden.
    Der öffentliche Demo-Modus kann danach simulierte Sponsoren auf der
    Live-Seite sichtbar machen.
  </figcaption>
</figure>

### Web Part mit der API verbinden

Solange der **Setup-Assistent** noch nicht abgeschlossen wurde, öffnet er sich
im Bearbeitungsmodus automatisch. Andernfalls öffnen Sie den
**Eigenschaftenbereich** manuell (Zahnrad-Symbol im Bearbeitungsmodus).
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

### Festlegen, was Gäste sehen sollen

Zusätzlich zur API-Verbindung umfasst der Eigenschaftenbereich deutlich mehr
Optionen, als auf diesem Screenshot sichtbar sind. Die eingeblendeten Bereiche
zeigen exemplarisch, wie Admins Sponsor-Eignung, Benachrichtigungen, Anzeige
und Kontaktdetails an die Datenqualität im Tenant anpassen und den
öffentlichen Demo-Modus auf der Live-Seite aktivieren können. So lässt sich
steuern, wie viel Besucher auf der Landingpage sehen.

<figure class="doc-screenshot doc-screenshot--narrow">
  <img src="{{ '/assets/images/my-sponsor-preferences.jpg' | relative_url }}"
    alt="Eigenschaftenbereich des Web Parts mit den sichtbaren Bereichen
    Settings, Sponsor Eligibility, Guest Notifications, Display und Contact."
    class="doc-screenshot__img" data-zoomable>
  <figcaption class="doc-screenshot__caption">
    Diese Beispielansicht zeigt nur einen Ausschnitt des Eigenschaftenbereichs,
    unter anderem mit Schalter für den öffentlichen Demo-Modus,
    Sponsor-Filtern und gastsichtigen Kontaktdetails.
  </figcaption>
</figure>

### Gastzugänglichkeit prüfen

> [!TIP]
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
