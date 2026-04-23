# Data Collection and Telemetry

> When you deploy this template, Microsoft can identify the installation of
> Workoho software with the deployed Azure resources. Microsoft can correlate
> these resources used to support the software. Microsoft collects this
> information to provide the best experiences with their products and to
> operate their business. The data is collected and governed by Microsoft's
> privacy policies, located at
> [https://www.microsoft.com/trustcenter](https://www.microsoft.com/trustcenter).

The ARM template uses Microsoft's
[Customer Usage Attribution (CUA)](https://aka.ms/partnercenter-attribution)
mechanism. When the template is deployed, Azure creates an empty nested
deployment in your resource group:

```text
pid-18fb4033-c9f3-41fa-a5db-e3a03b012939
```

Microsoft uses this GUID to forward aggregated Azure consumption figures
(compute hours, storage transactions, and similar billing signals) for that
resource group to
[Workoho](https://workoho.com/?utm_source=gsiw&utm_medium=docs&utm_campaign=repo&utm_content=telemetry)
via Partner Center. This helps Workoho understand how the solution is used and
justify continued development.

## What Is Not Collected or Shared

- No personal data (no user names, email addresses, or tenant IDs)
- No resource names, configurations, or secrets
- No data leaves your Azure subscription. Microsoft only shares summary
  consumption figures with Workoho using existing billing data

For information about personal data processed within your tenant at runtime,
see the [Privacy Policy](privacy-policy.md).

## What You Will See in Azure Portal

In Resource Group -> Deployments you will see the deployment named
`pid-18fb4033-c9f3-41fa-a5db-e3a03b012939`. It is an empty, harmless nested
deployment. Deleting it has no effect on running resources but stops future
attribution for that resource group.

## Opt Out

Set `enableTelemetry=false` during deployment:

```bash
az deployment group create \
  --resource-group <your-resource-group> \
  --template-uri https://github.com/workoho/spfx-guest-sponsor-info/releases/latest/download/azuredeploy.json \
  --parameters \
      tenantId=<your-tenant-id> \
      tenantName=<your-tenant-name> \
      functionAppName=<globally-unique-name> \
      webPartClientId=<client-id-from-pre-step> \
      enableTelemetry=false
```

Or via the [Deploy to Azure](../README.md#deploy-to-azure) button: expand
Telemetry in the parameter form and uncheck Enable Telemetry.

## Contact

For telemetry and privacy questions about this Solution, contact
[privacy@workoho.com](mailto:privacy@workoho.com).

For responsible disclosure of security vulnerabilities, contact
[security@workoho.com](mailto:security@workoho.com).
