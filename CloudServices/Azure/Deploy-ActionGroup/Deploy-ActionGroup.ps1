# =============================================================================
# Script: Deploy-ActionGroup.ps1
# Created: 2025-01-16 17:30:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-24 13:22:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.0
# Additional Info: Initial script creation
# =============================================================================

<#
.SYNOPSIS
Deploys an Azure Action Group using ARM templates.
.DESCRIPTION
Deploys an Azure Action Group using local ARM template and parameter files.
Requires Az PowerShell module and Azure authentication.
.EXAMPLE
.\Deploy-ActionGroup.ps1
Deploys the Action Group using template and parameter files in the AGTemplate-JB-TEST-RG2 subdirectory.
#>

# Install Azure PowerShell if not already installed
# Install-Module -Name Az

# Sign in to Azure (if not already signed in)
Connect-AzAccount

# Set the context to the correct subscription if needed
# Set-AzContext -SubscriptionId "YourSubscriptionId"

# Deploy ARM template using Azure PowerShell
New-AzResourceGroupDeployment `
    -Name AGDeployment `
    -ResourceGroupName JB-TEST-RG2 `
    -TemplateFile ".\AGTemplate-JB-TEST-RG2\template.json" `
    -TemplateParameterFile ".\AGTemplate-JB-TEST-RG2\parameters.json"
