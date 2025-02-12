# =============================================================================
# Script: Create-TestDomainController.ps1
# Created: 2025-02-11 23:45:10 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-12 20:48:32 UTC
# Updated By: jdyer-nuvodia
# Version: 4.6
# Additional Info: Added function validation and Az module dependency check
# =============================================================================

<#
.SYNOPSIS
    Creates a test domain controller in Azure with proper configuration and auto-shutdown.
.DESCRIPTION
    This script creates and configures a domain controller in Azure with the following:
    - Automated deployment of VM infrastructure
    - Domain controller role installation and configuration
    - Auto-shutdown scheduling
    - Proper time zone configuration
    - Comprehensive error handling and logging
    
    Prerequisites:
    - Azure PowerShell Az module
    - Required custom modules in the Modules directory
    - Azure subscription with proper permissions
.PARAMETER resourceGroupName
    The name of the resource group where the domain controller will be created
.PARAMETER location
    The Azure region where resources will be deployed
.PARAMETER vmName
    The name of the virtual machine to be created
.PARAMETER VMSize
    The size/SKU of the virtual machine
.PARAMETER vnetName
    The name of the virtual network to use or create
.PARAMETER subnetName
    The name of the subnet within the virtual network
.PARAMETER adminUsername
    The administrator username for the domain controller
.PARAMETER adminPassword
    The administrator password for the domain controller
.PARAMETER timeZoneId
    The timezone ID for the VM (default: US Mountain Standard Time)
.PARAMETER ValidateOnly
    Switch to only validate the configuration without deployment
.EXAMPLE
    .\Create-TestDomainController.ps1 -resourceGroupName "RG-DC-Test" -location "westus2" -vmName "DC01-Test"
    Creates a domain controller with specified parameters
.EXAMPLE
    .\Create-TestDomainController.ps1 -ValidateOnly
    Validates the configuration without deploying resources
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

$LogFile = Join-Path -Path $PSScriptRoot -ChildPath "Create-TestDomainController.log"
$BaseModulePath = Join-Path -Path $PSScriptRoot -ChildPath "Modules"
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Set-Content -Path $LogFile -Value "[$timestamp] Log file reset. New log starting."

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'ERROR', 'VALIDATION', 'WARNING')]
        [string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $logMessage
    Write-Host $logMessage
}

function Test-ModulePath {
    param (
        [string]$Path,
        [string]$ModuleName
    )
    if (-not (Test-Path -Path $Path)) {
        Write-Log "Module path not found: $Path" -Level ERROR
        return $false
    }
    $manifestPath = Join-Path -Path $Path -ChildPath "$ModuleName.psd1"
    $modulePath = Join-Path -Path $Path -ChildPath "$ModuleName.psm1"
    if (-not (Test-Path -Path $manifestPath)) {
        Write-Log "Module manifest not found: $manifestPath" -Level ERROR
        return $false
    }
    if (-not (Test-Path -Path $modulePath)) {
        Write-Log "Module file not found: $modulePath" -Level ERROR
        return $false
    }
    return $true
}

function Import-RequiredModule {
    param (
        [string]$ModulePath,
        [string]$ModuleName,
        [string[]]$RequiredFunctions = @()
    )
    $fullPath = $ModulePath
    if (-not (Test-ModulePath -Path $fullPath -ModuleName $ModuleName)) {
        return $false
    }
    try {
        $manifestPath = Join-Path -Path $fullPath -ChildPath "$ModuleName.psd1"
        if ($ModuleName -eq 'DC-Deployment') {
            if (-not (Get-Module -ListAvailable -Name Az.Compute)) {
                Write-Log "The Az.Compute module is required but not installed. Installing..." -Level WARNING
                Install-Module -Name Az -Scope CurrentUser -Force -AllowClobber
            }
        }
        Import-Module -Name $manifestPath -Force -ErrorAction Stop
        Write-Log "Successfully imported module: $ModuleName from $manifestPath" -Level INFO
        if (-not (Get-Module -Name $ModuleName -ErrorAction SilentlyContinue)) {
            throw "Module $ModuleName was not properly loaded after import"
        }
        foreach ($function in $RequiredFunctions) {
            if (-not (Get-Command -Name "$ModuleName\$function" -ErrorAction SilentlyContinue)) {
                Write-Log "Required function '$function' not found in module $ModuleName" -Level ERROR
                return $false
            }
        }
        if (Get-Command -Name "$ModuleName\Set-DCLogFile" -ErrorAction SilentlyContinue) {
            try {
                & "$ModuleName\Set-DCLogFile" -Path $LogFile
                Write-Log "Successfully initialized logging for module: $ModuleName" -Level INFO
            }
            catch {
                Write-Log "Warning: Could not initialize module logging for $ModuleName. Continuing anyway." -Level WARNING
            }
        }
        else {
            Write-Log "Note: Set-DCLogFile not found in module $ModuleName. Continuing with default logging." -Level INFO
        }
        return $true
    }
    catch {
        Write-Log ("Failed to import module {0}: {1}" -f $ModuleName, $_.Exception.Message) -Level ERROR
        return $false
    }
}

try {
    $modules = @(
        @{
            Path = Join-Path -Path $BaseModulePath -ChildPath "Configuration"
            Name = "DC-Configuration"
            RequiredFunctions = @('Initialize-DCConfiguration')
        },
        @{
            Path = Join-Path -Path $BaseModulePath -ChildPath "Validation"
            Name = "DC-Validation"
            RequiredFunctions = @('Test-DCPrerequisites')
        },
        @{
            Path = Join-Path -Path $BaseModulePath -ChildPath "Deployment"
            Name = "DC-Deployment"
            RequiredFunctions = @('New-DCEnvironment', 'Enable-AzVmAutoShutdown')
        }
    )
    $failedImports = 0
    foreach ($module in $modules) {
        Write-Log "Attempting to import module: $($module.Name)" -Level INFO
        if (-not (Import-RequiredModule -ModulePath $module.Path -ModuleName $module.Name -RequiredFunctions $module.RequiredFunctions)) {
            $failedImports++
        }
    }
    if ($failedImports -gt 0) {
        throw "Failed to import $failedImports required module(s). Please check the log for details."
    }
}
catch {
    Write-Log "Critical error during module import: $_" -Level ERROR
    return
}

try {
    Write-Log "Initializing configuration..." -Level INFO
    $config = Initialize-DCConfiguration
    if ($resourceGroupName) { 
        $config.ResourceGroupName = $resourceGroupName 
        Write-Log "Setting ResourceGroupName to: $resourceGroupName" -Level INFO
    }
    if ($location) { 
        $config.Location = $location 
        Write-Log "Setting Location to: $location" -Level INFO
    }
    if ($vmName) { 
        $config.VmName = $vmName 
        Write-Log "Setting VmName to: $vmName" -Level INFO
    }
    if ($VMSize) { 
        $config.VMSize = $VMSize 
        Write-Log "Setting VMSize to: $VMSize" -Level INFO
    }
    if ($vnetName) { 
        $config.VnetName = $vnetName 
        Write-Log "Setting VnetName to: $vnetName" -Level INFO
    }
    if ($subnetName) { 
        $config.SubnetName = $subnetName 
        Write-Log "Setting SubnetName to: $subnetName" -Level INFO
    }
    if ($adminUsername) { 
        $config.AdminUsername = $adminUsername 
        Write-Log "Setting AdminUsername to: $adminUsername" -Level INFO
    }
    if ($adminPassword) { 
        $config.AdminPassword = $adminPassword 
        Write-Log "Admin password has been set" -Level INFO
    }
    if ($timeZoneId) { 
        $config.TimeZoneId = $timeZoneId 
        Write-Log "Setting TimeZoneId to: $timeZoneId" -Level INFO
    }
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
    if ($PSCmdlet.ShouldProcess("Azure Resources", "Deploy")) {
        Write-Log "Starting deployment phase..." -Level INFO
        New-DCEnvironment -Config $config
        Write-Log "Domain Controller VM creation completed successfully." -Level INFO
    }
    else {
        Write-Log "Deployment cancelled by user." -Level INFO
    }
}
catch {
    Write-Log $_.Exception.Message -Level ERROR
    throw
}
finally {
    Write-Log "Script execution completed." -Level INFO
}