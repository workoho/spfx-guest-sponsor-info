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
  - [Step 1 - Update the function package](#step-1---update-the-function-package)
  - [Alternative - Manual deployment via Azure CLI](#alternative---manual-deployment-via-azure-cli)
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

### Integrity check before package upload (GitHub Releases path)

For update paths that upload `.sppkg` files manually (Tenant App Catalog or
Site Collection App Catalog), verify the package before upload:

> **Minimum:** verify SHA256 against `checksums.txt`.
>
> **Recommended:** verify SHA256 **and** run GitHub attestation verification.

1. Download `guest-sponsor-info.sppkg` and `checksums.txt` from the same
   release page (browser download is fine).
2. Verify SHA256 locally.

Linux/macOS:

```bash
sha256sum -c checksums.txt --ignore-missing
```

PowerShell:

```powershell
$expected = ((Select-String -Path ./checksums.txt -Pattern 'guest-sponsor-info.sppkg').Line -split ' +')[0].ToLower()
$actual = (Get-FileHash ./guest-sponsor-info.sppkg -Algorithm SHA256).Hash.ToLower()
if ($actual -ne $expected) { throw 'SHA256 mismatch for guest-sponsor-info.sppkg' }
```

Optional (requires `gh`):

```bash
gh attestation verify guest-sponsor-info.sppkg \
  --repo workoho/spfx-guest-sponsor-info
```

For AppSource updates, use **Upgrade Store App** and the marketplace delivery
path instead of manual package uploads.

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

The Azure Function uses the **Flex Consumption** plan. During deployment the
native Flex `onedeploy` step copies the selected release ZIP into the app's
configured deployment container. The Function App runs from that frozen copy —
**a restart alone does not pull a newer release**.

### Step 1 - Update the function package

Re-run the deployment wizard. By default, the installer resolves the newest
published infra release and uses that same release for the Azure Function
package. The wizard then republishes that function ZIP through the native Flex
deployment path, replacing the frozen copy.

`-Version` and `-AppVersion` control different artifacts:

- `-Version` selects the installer/infra payload and, by default, also the
  Azure Function package version.
- `-AppVersion` is an expert override when you intentionally want a different
  function package version, or when using `-Version main`.

To pin to a specific published release:

```powershell
& ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.ps1').Content)) -Version v1.x.y
```

In Azure Cloud Shell, prefer this PowerShell installer entry point instead of
`install.sh`. Expect a device code sign-in: the wizard authenticates with the
Azure CLI rather than reusing the Cloud Shell identity, whose Microsoft Graph
scopes cannot create the Entra app registration. Device code sign-in must
therefore be permitted for the deploying account, and a PAW-restricted admin
account has to confirm the code on its Privileged Access Workstation.

Use `-AppVersion` only as an expert override when the function package should
differ from the release selected by `-Version`.

Or, when running `deploy-azure.ps1` directly:

```powershell
./deploy-azure.ps1
```

`deploy-azure.ps1` also supports `-AppVersion`, but keep that for expert
override scenarios.

### Alternative - Manual deployment via Azure CLI

Use this only when you need to publish the ZIP directly instead of rerunning the
deployment wizard.

Before a manual upload, verify the function artifact integrity:

> **Minimum:** verify SHA256 against `checksums.txt`.
>
> **Recommended:** verify SHA256 **and** run GitHub attestation verification.

1. Download `released-package.zip` and `checksums.txt` from the
   same release page.
2. Verify SHA256 locally.

Linux/macOS:

```bash
sha256sum -c checksums.txt --ignore-missing
```

PowerShell:

```powershell
$expected = ((Select-String -Path ./checksums.txt -Pattern 'released-package.zip').Line -split ' +')[0].ToLower()
$actual = (Get-FileHash ./released-package.zip -Algorithm SHA256).Hash.ToLower()
if ($actual -ne $expected) { throw 'SHA256 mismatch for released-package.zip' }
```

Optional (requires `gh`):

```bash
gh attestation verify released-package.zip \
  --repo workoho/spfx-guest-sponsor-info
```

<details>
<summary>Manual deployment via Azure CLI</summary>

Flex Consumption does **not** use `config-zip`. The supported control-plane
path in this project is the native `onedeploy` site extension, which copies a
release ZIP from a reachable HTTPS URL into the Function App's configured
deployment container.

**Via Azure CLI ([Cloud Shell](https://shell.azure.com)):**

```bash
package_url="https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/released-package.zip"

az rest --method PUT \
  --url "https://management.azure.com/subscriptions/<subscription-id>/resourceGroups/<your-resource-group>/providers/Microsoft.Web/sites/<your-function-app-name>/extensions/onedeploy?api-version=2022-09-01" \
  --body "{\"properties\":{\"packageUri\":\"${package_url}\",\"remoteBuild\":false}}"
```

This re-triggers the same OneDeploy mechanism that the Bicep deployment uses.
Azure copies the ZIP into the deployment container configured on the Function
App; a plain restart would not fetch the new package.

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

Or, when you are already working from a local repo or extracted infra package,
re-run the Azure-only phase directly through `azd`:

```bash
azd env select <env-name>
azd provision --no-prompt
```

The direct `azd` path reuses the same pre/post hooks as the wizard. Missing
`AZURE_TENANT_ID`, `AZURE_FUNCTION_APP_NAME`, and `AZURE_WEB_PART_CLIENT_ID`
are derived automatically when possible. If
`AZURE_SKIP_GRAPH_ROLE_ASSIGNMENTS=true` is set, finish with
`setup-graph-permissions.ps1` afterwards.

</details>
