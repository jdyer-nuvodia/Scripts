# =============================================================================
# Script: Create-TestDomainController.ps1
# Created: 2025-02-11 23:45:10 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-12 00:41:56 UTC
# Updated By: jdyer-nuvodia
# Version: 3.2
# Additional Info: Updated timezone handling for Arizona operations
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

# Import required modules
$modulePath = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$modulePath\Modules\DC-Configuration.ps1"
. "$modulePath\Modules\DC-Validation.ps1"
. "$modulePath\Modules\DC-Deployment.ps1"

# Initialize logging
$LogFile = Join-Path $PSScriptRoot "Create-TestDomainController.log"
Set-Content -Path $LogFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Log file reset. New log starting."

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