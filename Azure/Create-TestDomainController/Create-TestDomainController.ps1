# =============================================================================
# Script: Create-TestDomainController.ps1
# Created: 2025-02-11 23:45:10 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-12 18:08:19 UTC
# Updated By: jdyer-nuvodia
# Version: 3.6
# Additional Info: Fixed module path joining logic
# =============================================================================

<#
.SYNOPSIS
    Creates a test domain controller as a Trusted Launch VM in Azure.
.DESCRIPTION
    This script orchestrates the creation of a domain controller VM in Azure using modular components.
    The process is split into validation, deployment, and configuration phases, each handled by
    separate modules for better maintainability and clarity.
.PARAMETER resourceGroupName
    The name of the resource group where the VM and related resources will be created.
.PARAMETER location
    The Azure region (location) to deploy the resources.
.PARAMETER vmName
    The name of the VM to create.
.PARAMETER VMSize
    The size of the VM (e.g., 'Standard_DS2_v2').
.PARAMETER vnetName
    The virtual network name for the VM.
.PARAMETER subnetName
    The name of the subnet within the virtual network.
.PARAMETER adminUsername
    The administrator username for the VM.
.PARAMETER adminPassword
    The administrator password for the VM.
.PARAMETER timeZoneId
    The timezone ID for VM auto-shutdown scheduling. Defaults to US Mountain Standard Time (Arizona).
.PARAMETER ValidateOnly
    Performs validation only without deploying resources.
.EXAMPLE
    .\Create-TestDomainController.ps1 -ValidateOnly
    Performs validation of all components without deployment.
#>


[CmdletBinding(SupportsShouldProcess=$true)]
param (
    [Parameter(Mandatory = $false)]
    [string]$resourceGroupName,
    [Parameter(Mandatory = $false)]
    [string]$location,
    [Parameter(Mandatory = $false)]
    [string]$vmName,
    [Parameter(Mandatory = $false)]
    [string]$VMSize,
    [Parameter(Mandatory = $false)]
    [string]$vnetName,
    [Parameter(Mandatory = $false)]
    [string]$subnetName,
    [Parameter(Mandatory = $false)]
    [string]$adminUsername,
    [Parameter(Mandatory = $false)]
    [string]$adminPassword,
    [Parameter(Mandatory = $false)]
    [string]$timeZoneId = 'US Mountain Standard Time',
    [Parameter(Mandatory = $false)]
    [switch]$ValidateOnly
)

# Get the script path for relative module imports
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path $ScriptPath "Modules"
$LogFile = Join-Path $ScriptPath "Create-TestDomainController.log"

# Initialize logging
Set-Content -Path $LogFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Log file reset. New log starting."

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'ERROR', 'VALIDATION', 'WARNING')]
        [string]$Level = 'INFO'
    )
    $logMessage = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    Add-Content -Path $LogFile -Value $logMessage
    Write-Host $logMessage
}

# Import required modules with error handling
try {
    $requiredModules = @(
        "Configuration\DC-Configuration",
        "Validation\DC-Validation",
        "Deployment\DC-Deployment"
    )

    foreach ($module in $requiredModules) {
        $modulePath = Join-Path $ModulePath $module
        $manifestPath = "$modulePath.psd1"
        
        if (Test-Path $manifestPath) {
            Import-Module $manifestPath -Force -ErrorAction Stop
            Write-Log "Successfully imported module: $module.psd1" -Level INFO
        } else {
            throw "Module manifest not found: $manifestPath"
        }
    }
} catch {
    Write-Log "Failed to import required modules: $_" -Level ERROR
    return
}

try {
    # Initialize default configuration
    $config = Initialize-DCConfiguration
    
    # Override defaults with provided parameters
    if ($resourceGroupName) { $config.ResourceGroupName = $resourceGroupName }
    if ($location)          { $config.Location = $location }
    if ($vmName)            { $config.VmName = $vmName }
    if ($VMSize)            { $config.VMSize = $VMSize }
    if ($vnetName)          { $config.VnetName = $vnetName }
    if ($subnetName)        { $config.SubnetName = $subnetName }
    if ($adminUsername)     { $config.AdminUsername = $adminUsername }
    if ($adminPassword)     { $config.AdminPassword = $adminPassword }
    if ($timeZoneId)        { $config.TimeZoneId = $timeZoneId }
    
    # Phase 1: Validation
    Write-Log "Starting validation phase..." -Level VALIDATION
    $validation = Test-DCPrerequisites -Config $config
    
    if (-not $validation.Success) {
        Write-Log "Validation failed:" -Level ERROR
        foreach ($message in $validation.Messages) {
            Write-Log $message -Level ERROR
        }
        throw "Resource validation failed. Please review the validation messages above."
    }
    Write-Log "All validations passed successfully." -Level VALIDATION
    
    if ($ValidateOnly) {
        Write-Log "Validation only mode - stopping before deployment." -Level INFO
        return
    }
    
    # Phase 2: Deployment
    if ($PSCmdlet.ShouldProcess("Azure Resources", "Deploy")) {
        Write-Log "Starting deployment phase..." -Level INFO
        New-DCEnvironment -Config $config
        Write-Log "Domain Controller VM creation completed successfully." -Level INFO
    } else {
        Write-Log "Deployment cancelled by user." -Level INFO
    }
} catch {
    Write-Log $_.Exception.Message -Level ERROR
    throw
} finally {
    Write-Log "Script execution completed." -Level INFO
}