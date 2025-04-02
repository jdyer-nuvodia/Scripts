# =============================================================================
# Script: Deploy-ActionGroup.ps1
# Created: 2025-01-16 17:30:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-02 20:43:00 UTC
# Updated By: GitHub-Copilot
# Version: 1.1.0
# Additional Info: Added parameter validation and enhanced documentation
# =============================================================================

<#
.SYNOPSIS
    Deploys an Azure Action Group using ARM templates.
.DESCRIPTION
    This script deploys an Azure Action Group using local ARM template and parameter files.
    It handles the following tasks:
    - Verifies Azure PowerShell module installation
    - Validates ARM template files exist
    - Connects to Azure (if not connected)
    - Deploys the Action Group using specified templates
    
    Dependencies:
    - Az PowerShell module (Install-Module -Name Az)
    - Valid Azure subscription
    - Appropriate Azure RBAC permissions
    
    Required files:
    - template.json: ARM template file
    - parameters.json: ARM template parameters file
.PARAMETER ResourceGroupName
    Name of the target Azure Resource Group. Default is 'JB-TEST-RG2'.
.PARAMETER TemplateFolder
    Path to the folder containing template.json and parameters.json files.
    Default is '.\AGTemplate-JB-TEST-RG2'.
.PARAMETER DeploymentName
    Name for the deployment operation. Default is 'AGDeployment'.
.PARAMETER SubscriptionId
    Optional. Azure subscription ID to target for deployment.
.EXAMPLE
    .\Deploy-ActionGroup.ps1
    Deploys the Action Group using default template location and resource group.
.EXAMPLE
    .\Deploy-ActionGroup.ps1 -ResourceGroupName "Production-RG" -TemplateFolder ".\Templates\Prod"
    Deploys to a specific resource group using templates from a custom location.
.EXAMPLE
    .\Deploy-ActionGroup.ps1 -SubscriptionId "00000000-0000-0000-0000-000000000000"
    Deploys to a specific Azure subscription.
.NOTES
    Security Level: High
    Required Permissions: Azure Contributor or specific RBAC roles
    Validation Requirements:
    - Verify Az PowerShell module is installed
    - Verify template files exist and are valid
    - Verify Azure authentication
    - Verify resource group exists
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$ResourceGroupName = "JB-TEST-RG2",

    [Parameter(Mandatory=$false)]
    [ValidateScript({
        if (-Not (Test-Path $_)) {
            throw "Template folder does not exist"
        }
        if (-Not (Test-Path (Join-Path $_ "template.json"))) {
            throw "template.json not found in specified folder"
        }
        if (-Not (Test-Path (Join-Path $_ "parameters.json"))) {
            throw "parameters.json not found in specified folder"
        }
        return $true
    })]
    [string]$TemplateFolder = ".\AGTemplate-JB-TEST-RG2",

    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$DeploymentName = "AGDeployment",

    [Parameter(Mandatory=$false)]
    [ValidatePattern('^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$')]
    [string]$SubscriptionId
)

# Initialize logging
$LogPath = Join-Path $PSScriptRoot "AGDeployment_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$ErrorActionPreference = "Stop"

function Write-Log {
    param($Message, $Level = "Information")
    
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC"
    $LogMessage = "$TimeStamp [$Level] $Message"
    Add-Content -Path $LogPath -Value $LogMessage
    
    switch ($Level) {
        "Information" { Write-Host $Message -ForegroundColor White }
        "Success"     { Write-Host $Message -ForegroundColor Green }
        "Warning"     { Write-Host $Message -ForegroundColor Yellow }
        "Error"       { Write-Host $Message -ForegroundColor Red }
        "Process"     { Write-Host $Message -ForegroundColor Cyan }
    }
}

try {
    Write-Log "Starting Action Group deployment process" "Process"
    
    # Verify Az module is installed
    if (-not (Get-Module -ListAvailable -Name Az)) {
        Write-Log "Az PowerShell module not found. Please install using: Install-Module -Name Az" "Error"
        throw "Required Az module not installed"
    }
    
    # Connect to Azure if needed
    try {
        $null = Get-AzContext -ErrorAction Stop
        Write-Log "Already connected to Azure" "Success"
    }
    catch {
        Write-Log "Not connected to Azure. Initiating sign-in..." "Process"
        Connect-AzAccount
    }
    
    # Set subscription context if provided
    if ($SubscriptionId) {
        Write-Log "Setting context to subscription: $SubscriptionId" "Process"
        Set-AzContext -SubscriptionId $SubscriptionId
    }
    
    # Verify resource group exists
    if (-not (Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction SilentlyContinue)) {
        throw "Resource group '$ResourceGroupName' not found"
    }
    
    # Prepare template paths
    $TemplateFile = Join-Path $TemplateFolder "template.json"
    $ParameterFile = Join-Path $TemplateFolder "parameters.json"
    
    Write-Log "Deploying Action Group..." "Process"
    Write-Log "Resource Group: $ResourceGroupName" "Information"
    Write-Log "Template File: $TemplateFile" "Information"
    Write-Log "Parameter File: $ParameterFile" "Information"
    
    # Deploy ARM template
    $deployment = New-AzResourceGroupDeployment `
        -Name $DeploymentName `
        -ResourceGroupName $ResourceGroupName `
        -TemplateFile $TemplateFile `
        -TemplateParameterFile $ParameterFile
    
    if ($deployment.ProvisioningState -eq "Succeeded") {
        Write-Log "Action Group deployment completed successfully" "Success"
    }
    else {
        throw "Deployment failed with state: $($deployment.ProvisioningState)"
    }
}
catch {
    Write-Log "Deployment failed: $_" "Error"
    throw
}
finally {
    Write-Log "Deployment process finished" "Information"
}
