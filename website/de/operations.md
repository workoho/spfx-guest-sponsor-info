---
layout: doc
lang: de
title: Operations Guide
permalink: /de/operations/
description: >-
  Day-2-Betriebsreferenz — Web Part aktualisieren, Azure Maps
  konfigurieren und Azure Function aktualisieren.
lead: >-
  Laufende Administrationsaufgaben für SharePoint- und
  Azure-Administratoren. Für die Ersteinrichtung verwenden Sie
  die Setup-Anleitung.
github_doc: operations.md
---

## Web Part aktualisieren

Für nachfolgende Versionsupdates gelten die drei Bedingungen aus der
Ersteinrichtung nicht mehr. Nur **Websitesammlungsadministrator** auf
der Landingpage-Site ist erforderlich — Sie benötigen weder die
SharePoint-Admin-Rolle noch Zugriff auf den mandantenweiten App Catalog.

Laden Sie die neue `.sppkg` über die bestehende hoch unter:
`https://<tenant>.sharepoint.com/sites/<landing-site>/AppCatalog/`

Alternativer Pfad: **Websiteinhalte → Apps für SharePoint**.

SharePoint ersetzt die vorherige Version sofort. Im Normalfall ist weder
eine Seitenveröffentlichung noch eine Cache-Leerung erforderlich.

> **Tipp:** Jeder Benutzer mit **Vollständige Kontrolle** auf der
> Landingpage-Site (z.B. ein Websitebesitzer) kann den Upload durchführen.

---

## Inline-Adresskarte (Azure Maps)

Die ARM-Vorlage stellt standardmäßig ein Azure Maps-Konto bereit
(`deployAzureMaps=true`).

### Kartenrendering aktivieren

1. Schlüssel abrufen:

   ```bash
   az maps account keys list \
     -g <ressourcengruppe> \
     -n <azure-maps-kontoname> \
     --query primaryKey -o tsv
   ```

2. Im Web-Part-Eigenschaftenbereich:
   - **Adresskarten-Vorschau anzeigen** aktivieren
   - Schlüssel in **Azure Maps Subscription Key** einfügen
   - Fallback-Anbieter wählen (`Bing`, `Google`, `Apple`,
     `OpenStreetMap`, `HERE`)

Ohne Azure Maps-Schlüssel (oder bei fehlgeschlagenem Geocoding) zeigt
die Karte einen Fallback-Link zum externen Kartenanbieter.

### CSP-eingeschränkte Umgebungen

Erlauben Sie mindestens:

- `https://atlas.microsoft.com` (Geocoding und statisches Kartenbild)
- Die Domain des gewählten externen Kartenanbieters für Fallback-Links

### Schnelle Entscheidungshilfe

1. Behalten Sie `deployAzureMaps=true` — die Bereitstellung allein kostet
   nichts.
2. Geben Sie den Schlüssel im Web Part erst ein, wenn Sie Inline-Karten
   möchten.
3. Kein konfigurierter Schlüssel = keine Azure Maps-Anfragen = keine Kosten.

**Abrechnung:** Azure Maps basiert auf Anfragen mit einem kostenlosen
monatlichen Kontingent (S0).

---

## Function aktualisieren

### Consumption-Plan

Die Function App verwendet `WEBSITE_RUN_FROM_PACKAGE` mit Verweis auf das
aktuelle GitHub-Release-ZIP. Ein Neustart lädt das aktuelle ZIP:

```bash
az functionapp restart \
  --resource-group <ihre-ressourcengruppe> \
  --name <ihr-function-app-name>
```

Oder im Azure-Portal: **Function App → Übersicht → Neustart**.

### Flex Consumption-Plan

Stellen Sie die ARM-Vorlage mit einer festen `appVersion` erneut bereit:

```bash
az deployment group create \
  --resource-group <ihre-ressourcengruppe> \
  --template-uri https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/azuredeploy.json \
  --parameters \
      tenantId=<ihre-tenant-id> \
      tenantName=<ihr-tenant-name> \
      functionAppName=<ihr-function-app-name> \
      functionClientId=<ihre-client-id> \
      hostingPlan=FlexConsumption \
      maximumFlexInstances=10 \
      appVersion=1.x.y
```

<details>
<summary>Manueller Upload über Azure-Portal oder CLI</summary>

**Über das Azure-Portal:**

1. Öffnen Sie **Speicherkonto → Container → `app-package`**.
2. Laden Sie die ZIP-Datei von der
   [Releases-Seite](https://github.com/workoho/spfx-guest-sponsor-info/releases) hoch.
3. Unter „Erweitert" setzen Sie den Blob-Namen auf `function.zip`,
   aktivieren Sie „Überschreiben" und laden Sie hoch.

**Über Azure CLI ([Cloud Shell](https://shell.azure.com)):**

```bash
curl -sSfL -o function.zip \
  https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/guest-sponsor-info-function.zip

az storage blob upload \
  --account-name <speicherkontoname> \
  --container-name app-package \
  --name function.zip \
  --file function.zip \
  --auth-mode login \
  --overwrite
```

</details>

<details>
<summary>Infrastruktur geändert? Vollständige Bereitstellung erneut ausführen</summary>

Wenn ein Release angibt, dass Azure-Infrastruktur aktualisiert wurde,
führen Sie die ARM-Bereitstellung erneut aus (idempotent):

```bash
az deployment group create \
  --resource-group <ihre-ressourcengruppe> \
  --template-uri https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/azuredeploy.json \
  --parameters \
      tenantId=<ihre-tenant-id> \
      tenantName=<ihr-tenant-name> \
      functionAppName=<ihr-function-app-name> \
      functionClientId=<ihre-client-id>
```

Für Deployment Stacks verwenden Sie `az stack group create` mit denselben
Parametern.

Zum Entfernen aller bereitgestellten Ressourcen:

```bash
az stack group delete \
  --name guest-sponsor-info \
  --resource-group <ihre-ressourcengruppe> \
  --action-on-unmanage deleteResources \
  --yes
```

</details>

---

## Support

Bei Problemen oder Fragen zum Betrieb siehe die [Support]({{ '/de/support/' | relative_url }})-Seite.
