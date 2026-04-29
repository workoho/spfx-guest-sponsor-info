#!/usr/bin/env -S pwsh -NoLogo -NoProfile

# Verify-DeploymentGuid.ps1
#
# Source: https://gist.github.com/bmoore-msft/ae6b8226311014d6e7177c5127c7eba1
# Author: Brian Moore (Microsoft, Partner Center team)
#
# PURPOSE
# -------
# After deploying via deploy-azure.ps1 / install.ps1, use this script to
# verify that Azure correctly correlated the Customer Usage Attribution (CUA)
# deployment
# (pid-18fb4033-c9f3-41fa-a5db-e3a03b012939) with the actual resources.
#
# The script follows the correlationId of the pid-* deployment and lists every
# resource that was deployed in the same correlation scope. A non-empty list
# confirms attribution is working. An empty list indicates the pid-* deployment
# was created outside a real deployment (e.g. it was added manually after the
# fact), in which case attribution will not be credited.
#
# PREREQUISITES
# -------------
# - Az.Accounts and Az.Resources PowerShell modules — pre-installed in the
#   dev container. Outside the container:
#   Install-Module -Name Az.Accounts, Az.Resources -Scope CurrentUser
# - Logged in:                       Connect-AzAccount
# - Correct subscription selected:   Set-AzContext -SubscriptionId <id>
#
# USAGE
# -----
# .\Verify-DeploymentGuid.ps1 `
#   -deploymentName pid-18fb4033-c9f3-41fa-a5db-e3a03b012939 `
#   -resourceGroupName <your-resource-group>
#
# WHEN TO RUN
# -----------
# Run this after every fresh Bicep/azd deployment (via install.ps1 or
# deploy-azure.ps1) to confirm attribution before a Partner Center
# reporting period. Not needed for day-to-day development.

Param(
  # Full name of the pid-* deployment as it appears in Resource Group → Deployments,
  # e.g. pid-18fb4033-c9f3-41fa-a5db-e3a03b012939
  [string][Parameter(Mandatory = $true)]$deploymentName,
  [string][Parameter(Mandatory = $true)]$resourceGroupName
)

# Retrieve the correlationId that Azure assigned to the pid-* deployment.
# All resources deployed in the same Bicep/azd deployment share this correlationId.
$correlationId = (Get-AzResourceGroupDeployment `
    -ResourceGroupName $resourceGroupName `
    -Name "$deploymentName").correlationId

# Find all deployments in the resource group that share that correlationId.
# The pid-* deployment and the parent deployment will both appear here.
$deployments = Get-AzResourceGroupDeployment -ResourceGroupName $resourceGroupName |
Where-Object { $_.correlationId -eq $correlationId }

# For each deployment in the correlation set, list the real Azure resources
# that were created (skip nested Microsoft.Resources/deployments entries).
# PowerShell doesn't expose outputResources or correlationId on individual
# deployment operations, so we enumerate operations and filter by resource type.
foreach ($deployment in $deployments) {
  ($deployment |
  Get-AzResourceGroupDeploymentOperation |
  Where-Object { $_.targetResource -notlike "*Microsoft.Resources/deployments*" }
  ).TargetResource
}
