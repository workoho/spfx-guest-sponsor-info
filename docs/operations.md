# Administration and Operations Guide

Day-2 operations reference for SharePoint and Azure administrators.

For initial setup, use [deployment.md](deployment.md).
For security posture and trust assumptions, see
[security-assessment.md](security-assessment.md).
For telemetry and attribution details, see [telemetry.md](telemetry.md).

## Table of Contents

- [Updating the Web Part](#updating-the-web-part)
  - [Step 1 - Identify the original deployment path](#step-1---identify-the-original-deployment-path)
  - [Step 2 - Update the package source](#step-2---update-the-package-source)
  - [Step 3 - Upgrade the installed site app instance](#step-3---upgrade-the-installed-site-app-instance)
- [Inline Address Map (Azure Maps)](#inline-address-map-azure-maps)
  - [Step 1 - Decide whether inline map rendering is needed](#step-1---decide-whether-inline-map-rendering-is-needed)
  - [Step 2 (Optional) - Configure the Azure Maps key](#step-2-optional---configure-the-azure-maps-key)
  - [Step 3 (Optional) - Allow required endpoints in CSP](#step-3-optional---allow-required-endpoints-in-csp)
- [Updating the Function](#updating-the-function)
  - [Step 1 - Identify the hosting plan and package mode](#step-1---identify-the-hosting-plan-and-package-mode)
  - [Step 2A - Update a Consumption deployment](#step-2a---update-a-consumption-deployment)
  - [Step 2B - Update a Flex Consumption deployment](#step-2b---update-a-flex-consumption-deployment)
  - [Alternative - Manual package upload](#alternative---manual-package-upload)
  - [Advanced - Re-run the full deployment when infrastructure changes](#advanced---re-run-the-full-deployment-when-infrastructure-changes)

---

## Updating the Web Part

### Step 1 - Identify the original deployment path

The web part update path depends on how the package was installed during the
initial deployment:

- **Site Collection App Catalog** - the package is stored in the landing site's
  `Apps for SharePoint` library (`https://<tenant>.sharepoint.com/sites/<landing-site>/AppCatalog/`).
- **Tenant App Catalog** - the package is stored in the tenant-level
  `Apps for SharePoint` library.
- **AppSource** - the package entered the tenant through the marketplace-backed
  tenant App Catalog flow.

If you are not sure which path was used originally, check
[deployment.md](deployment.md#step-1---install-the-web-part).

### Step 2 - Update the package source

#### Option A - Site Collection App Catalog deployment

Use this path when the web part was installed from the landing site's Site
Collection App Catalog.

1. Open the landing site's `Apps for SharePoint` library at:
  `https://<tenant>.sharepoint.com/sites/<landing-site>/AppCatalog/`
2. Upload the new `.sppkg` over the existing package.
3. Click **Deploy** if SharePoint prompts for confirmation.

Required access: permission to upload to the site's `Apps for SharePoint`
library.

#### Option B - Tenant App Catalog deployment

Use this path when the package was uploaded directly to the tenant App Catalog.

1. Open **SharePoint Admin Center -> More features -> Apps -> Open**.
2. Upload the new `.sppkg` to **Apps for SharePoint**.
3. Click **Deploy** when SharePoint shows the deployment dialog.

#### Option C - AppSource deployment

Use this path when the package was originally acquired from Microsoft
AppSource.

1. Open the tenant App Catalog.
2. Select the app and use **Upgrade Store App** when a newer marketplace
  version is available.
3. After the tenant package is upgraded, continue with Step 3 below.

### Step 3 - Upgrade the installed site app instance

This solution uses `skipFeatureDeployment: false`. Updating the package source
does not by itself update the installed app instance on the landing-page site.

1. Open **Site Contents** on the landing-page site.
2. If SharePoint shows an **Update** banner on **Guest Sponsor Info**, click
  **Update**.
3. If no update banner appears, remove the app instance and add it again via
  **Site Contents -> Add an app**.

No page republish or manual cache flush is normally required after the updated
app instance is active.

---

## Inline Address Map (Azure Maps)

### Step 1 - Decide whether inline map rendering is needed

The Azure deployment creates an Azure Maps account by default
(`deployAzureMaps=true`), but the web part does not use it until you configure
an Azure Maps subscription key in the property pane.

If you leave the key empty, the sponsor card falls back to the external map
link. No Azure Maps requests are sent in that state.

### Step 2 (Optional) - Configure the Azure Maps key

1. Get the key:

   ```bash
   az maps account keys list \
     -g <resource-group> \
     -n <azure-maps-account-name> \
     --query primaryKey -o tsv
   ```

2. In the web part property pane:
    - Enable **Show address map preview**
    - Paste the value into **Azure Maps subscription key**
    - Choose the fallback provider (`Bing`, `Google`, `Apple`, `OpenStreetMap`)

### Step 3 (Optional) - Allow required endpoints in CSP

If your environment uses a restrictive Content Security Policy, allow at least:

- `https://atlas.microsoft.com` (geocoding and static map image)
- The selected external map provider domain for fallback links

Azure Maps pricing is request-based with a free monthly quota on S0. If no key
is configured in the web part, no Azure Maps requests are issued by this
solution.

---

## Updating the Function

### Step 1 - Identify the hosting plan and package mode

The function update path depends on the hosting plan chosen during initial
deployment:

- **Consumption** - `WEBSITE_RUN_FROM_PACKAGE` points to a GitHub release ZIP
  URL. The default deployment tracks `latest`.
- **Flex Consumption** - the deployment uploads `function.zip` to the
  `app-package` container and the app runs from that storage-backed package URL.

If you are not sure which mode is active, inspect the
`WEBSITE_RUN_FROM_PACKAGE` app setting.

### Step 2A - Update a Consumption deployment

The default Consumption deployment uses `WEBSITE_RUN_FROM_PACKAGE` pointing to
the `latest` GitHub Release ZIP. In that default mode, a restart pulls the
current ZIP:

```bash
az functionapp restart \
  --resource-group <your-resource-group> \
  --name <your-function-app-name>
```

Or from the Azure Portal: Function App -> Overview -> Restart.

If the deployment was pinned to a specific `appVersion`, a full re-deployment is
not strictly required for a code-only update. You can also point
`WEBSITE_RUN_FROM_PACKAGE` to a different release ZIP and let Azure restart the
app:

```bash
az functionapp config appsettings set \
  --resource-group <your-resource-group> \
  --name <your-function-app-name> \
  --settings WEBSITE_RUN_FROM_PACKAGE=https://github.com/workoho/spfx-guest-sponsor-info/releases/download/<release-tag>/guest-sponsor-info-function.zip
```

Or in the Azure Portal:

1. Open **Function App -> Settings -> Environment variables**.
2. Edit `WEBSITE_RUN_FROM_PACKAGE`.
3. Replace the URL with the new GitHub Release ZIP URL.
4. Click **Apply**, then **Confirm**. Saving the app setting restarts the app.

Re-running the deployment is still the safer default when you want the Bicep /
`azd` state to stay aligned with the deployed version, or when the release also
includes infrastructure changes.

### Step 2B - Update a Flex Consumption deployment

Re-run the deployment wizard with a pinned `appVersion`:

```powershell
& ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.ps1').Content)) -AppVersion 1.x.y
```

Or, when running `deploy-azure.ps1` directly:

```powershell
./deploy-azure.ps1 -AppVersion 1.x.y
```

### Alternative - Manual package upload

Use this only when you need to replace the storage-backed package directly
instead of rerunning the deployment wizard.

<details>
<summary>Manual upload via Azure Portal or CLI</summary>

**Via the Azure Portal:**

1. Open **Storage Account -> Containers -> app-package**.
2. Upload and select the ZIP from the
   [Releases page](https://github.com/workoho/spfx-guest-sponsor-info/releases).
3. Under Advanced, set blob name to `function.zip`, enable overwrite, then
   upload.
4. Restart the Function App so the unchanged package URL is reloaded and
  triggers are refreshed.

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

az functionapp restart \
  --resource-group <your-resource-group> \
  --name <your-function-app-name>
```

</details>

### Advanced - Re-run the full deployment when infrastructure changes

Use this when release notes indicate an infrastructure change, not just a code
update.

<details>
<summary>Infrastructure changed? Re-run the full deployment</summary>

If a release states that Azure infrastructure was updated, re-run the
deployment wizard (idempotent):

```powershell
& ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.ps1').Content))
```

Or, when running `deploy-azure.ps1` directly from an extracted infra ZIP:

```powershell
./deploy-azure.ps1
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
