---
layout: doc
lang: de
title: Telemetrie
permalink: /de/telemetry/
description: >-
  Welche Daten die Azure-Bereitstellung von Guest Sponsor Info erfasst,
  was nicht erfasst wird und wie Sie die Erfassung deaktivieren.
lead: >-
  Eine verständliche Zusammenfassung der Telemetrie in der
  Azure-Bereitstellung. Es werden keine personenbezogenen Daten
  erfasst oder weitergegeben.
github_doc: telemetry.md
---

## Was bei der Bereitstellung passiert

Wenn Sie die Azure Function mit der mitgelieferten ARM-Vorlage
bereitstellen, wird ein kleiner Tracking-Marker in Ihrer
Ressourcengruppe angelegt:

```text
pid-18fb4033-c9f3-41fa-a5db-e3a03b012939
```

Das ist eine leere, harmlose verschachtelte Bereitstellung. Microsoft
verwendet diese GUID, um **aggregierte Azure-Verbrauchszahlen**
(Rechenzeit, Speicher-Transaktionen und ähnliche Abrechnungssignale)
für diese Ressourcengruppe über Partner Center an
[Workoho](https://workoho.com) weiterzuleiten.

Dieser Mechanismus heißt
[Customer Usage Attribution (CUA)](https://aka.ms/partnercenter-attribution)
und hilft Workoho zu verstehen, wie die Lösung genutzt wird, um die
Weiterentwicklung zu rechtfertigen.

## Was NICHT erfasst wird

- **Keine personenbezogenen Daten** — keine Benutzernamen, E-Mail-Adressen
  oder Tenant-IDs
- **Keine Ressourcennamen**, Konfigurationen oder Geheimnisse
- **Keine Daten verlassen Ihr Azure-Abonnement** — Microsoft teilt nur
  zusammengefasste Verbrauchszahlen mit Workoho auf Basis vorhandener
  Abrechnungsdaten

Informationen über personenbezogene Daten, die zur Laufzeit in Ihrem Tenant
verarbeitet werden, finden Sie in der
[Datenschutzrichtlinie](https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/privacy-policy.md).

## Was Sie im Azure-Portal sehen

Unter **Ressourcengruppe → Bereitstellungen** sehen Sie eine Bereitstellung
namens `pid-18fb4033-c9f3-41fa-a5db-e3a03b012939`. Es ist eine leere
verschachtelte Bereitstellung. Das Löschen hat keine Auswirkung auf
laufende Ressourcen, beendet aber die zukünftige Zuordnung für diese
Ressourcengruppe.

## So deaktivieren Sie die Telemetrie

Setzen Sie `enableTelemetry=false` bei der Bereitstellung:

```bash
az deployment group create \
  --resource-group <ihre-ressourcengruppe> \
  --template-uri https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/azuredeploy.json \
  --parameters \
      tenantId=<ihre-tenant-id> \
      tenantName=<ihr-tenant-name> \
      functionAppName=<global-eindeutiger-name> \
      functionClientId=<client-id-aus-vorbereitung> \
      enableTelemetry=false
```

Oder über die **Deploy to Azure**-Schaltfläche: Erweitern Sie *Telemetry*
im Parameterformular und deaktivieren Sie *Enable Telemetry*.

## Kontakt

Für Fragen zu Telemetrie und Datenschutz dieser Lösung wenden Sie sich an
[privacy@workoho.com](mailto:privacy@workoho.com).

Für die verantwortungsvolle Offenlegung von Sicherheitslücken wenden
Sie sich an [security@workoho.com](mailto:security@workoho.com).
