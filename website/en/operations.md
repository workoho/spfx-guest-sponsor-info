---
layout: doc
lang: en
title: Operations Guide
permalink: /en/operations/
description: >-
  Day-2 operations reference — updating the web part, configuring
  Azure Maps, and updating the Azure Function.
lead: >-
  Day-2 administration tasks for SharePoint and Azure administrators.
  For initial setup, use the Deployment Guide.
github_doc: operations.md
---

## Updating the Web Part

For subsequent version updates, the three conditions from the initial setup do
not apply. Only **Site Collection Administrator** on the landing-page site is
required — you do not need the SharePoint Admin role and do not need access to
the tenant App Catalog.

Upload the new `.sppkg` over the existing one at:
`https://<tenant>.sharepoint.com/sites/<landing-site>/AppCatalog/`

Alternative path: **Site Contents → Apps for SharePoint**.

SharePoint replaces the previous version immediately. No page republish or
cache flush is needed in normal circumstances.

> **Tip:** any user who has **Full Control** on the landing-page site (for
> example, a Site Owner) can perform the upload. Site Collection Administrator
> is only required when the site's permission model restricts App Catalog
> library access.

---

## Inline Address Map (Azure Maps)

The ARM template deploys an Azure Maps account by default
(`deployAzureMaps=true`).

### Enable map rendering

1. Get the key:

   ```bash
   az maps account keys list \
     -g <resource-group> \
     -n <azure-maps-account-name> \
     --query primaryKey -o tsv
   ```

2. In the web part property pane:
   - Enable **Show address map preview**
   - Paste the key into **Azure Maps subscription key**
   - Choose fallback provider (`Bing`, `Google`, `Apple`, `OpenStreetMap`,
     `HERE`)

Without an Azure Maps key (or when geocoding fails), the card shows an
external map link fallback.

### CSP-restricted environments

Allow at least:

- `https://atlas.microsoft.com` (geocoding and static map image)
- The selected external map provider domain for fallback links

### Quick decision guide

1. Keep `deployAzureMaps=true` — deploying Azure Maps costs nothing initially.
2. Enter the key in the web part only when you want inline maps.
3. No key configured means no Azure Maps requests are issued.

**Billing:** Azure Maps pricing is request-based with a free monthly quota
(S0). No key configured in the web part = no requests = no cost.

---

## Updating the Function

### Consumption plan

The Function App uses `WEBSITE_RUN_FROM_PACKAGE` pointing to the latest GitHub
Release ZIP. A restart pulls the current ZIP:

```bash
az functionapp restart \
  --resource-group <your-resource-group> \
  --name <your-function-app-name>
```

Or from the Azure Portal: **Function App → Overview → Restart**.

### Flex Consumption plan

Re-deploy the ARM template with a pinned `appVersion`:

```bash
az deployment group create \
  --resource-group <your-resource-group> \
  --template-uri https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/azuredeploy.json \
  --parameters \
      tenantId=<your-tenant-id> \
      tenantName=<your-tenant-name> \
      functionAppName=<your-function-app-name> \
      functionClientId=<your-client-id> \
      hostingPlan=FlexConsumption \
      maximumFlexInstances=10 \
      appVersion=1.x.y
```

<details>
<summary>Manual upload via Azure Portal or CLI</summary>

**Via the Azure Portal:**

1. Open **Storage Account → Containers → `app-package`**.
2. Upload the ZIP from the
   [Releases page](https://github.com/workoho/spfx-guest-sponsor-info/releases).
3. Under Advanced, set blob name to `function.zip`, enable overwrite, then
   upload.

**Via Azure CLI ([Cloud Shell](https://shell.azure.com)):**

```bash
curl -sSfL -o function.zip \
  https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/guest-sponsor-info-function.zip

az storage blob upload \
  --account-name <storage-account-name> \
  --container-name app-package \
  --name function.zip \
  --file function.zip \
  --auth-mode login \
  --overwrite
```

</details>

<details>
<summary>Infrastructure changed? Re-run the full deployment</summary>

If a release states that Azure infrastructure was updated, re-run the ARM
deployment (idempotent):

```bash
az deployment group create \
  --resource-group <your-resource-group> \
  --template-uri https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/azuredeploy.json \
  --parameters \
      tenantId=<your-tenant-id> \
      tenantName=<your-tenant-name> \
      functionAppName=<your-function-app-name> \
      functionClientId=<your-client-id>
```

For Deployment Stacks, use `az stack group create` with the same parameters.

To remove all deployed resources:

```bash
az stack group delete \
  --name guest-sponsor-info \
  --resource-group <your-resource-group> \
  --action-on-unmanage deleteResources \
  --yes
```

</details>
