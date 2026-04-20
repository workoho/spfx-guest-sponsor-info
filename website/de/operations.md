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

Wenn eine neue Version im AppSource veröffentlicht wird, erscheint sie als
ausstehende Aktualisierung unter **SharePoint Admin Center → Apps → Öffnen**.
Bestätigen Sie die Aktualisierung dort, um sie auf allen Sites mit installierter
App zu deployen.

SharePoint ersetzt die vorherige Version sofort. Im Normalfall ist weder
eine Seitenveröffentlichung noch eine Cache-Leerung erforderlich.

> **Erweiterte Deployment-Szenarien** — Wenn Sie außerhalb des AppSource
> bereitgestellt haben (direkter Upload in den Tenant App Catalog oder über
> einen Site Collection App Catalog), finden Sie das entsprechende
> Aktualisierungsverfahren im
> [Operations-Dokument auf GitHub](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/operations.md).

---

## Inline-Adresskarte (Azure Maps) {#inline-adresskarte-azure-maps}

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
   - Fallback-Anbieter wählen (`Bing`, `Google`, `Apple`, `OpenStreetMap`)

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
monatlichen Kontingent (Gen2).

---

## Function aktualisieren

### Consumption-Plan

Die Function App verwendet `WEBSITE_RUN_FROM_PACKAGE` mit Verweis auf das
GitHub-Release-ZIP. Bei Bereitstellung mit `appVersion=latest` (Standard)
lädt ein Neustart immer das aktuelle Release:

```bash
az functionapp restart \
  --resource-group <ihre-ressourcengruppe> \
  --name <ihr-function-app-name>
```

Oder im Azure-Portal: **Function App → Übersicht → Neustart**.

> **Feste Version?** Wenn Sie ursprünglich mit einer bestimmten `appVersion`
> bereitgestellt haben, lädt ein Neustart keine neuere Version. Führen Sie
> zunächst die ARM-Bereitstellung mit der neuen Versionsnummer (oder
> `appVersion=latest`) erneut aus, um die Paket-URL zu aktualisieren,
> und starten Sie dann neu.

### Flex Consumption-Plan

Stellen Sie die ARM-Vorlage mit der neuen Version erneut bereit:

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
