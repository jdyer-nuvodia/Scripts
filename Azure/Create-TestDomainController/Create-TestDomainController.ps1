# =============================================================================
# Script: Create-TestDomainController.ps1
# Created: 2025-02-11 23:45:10 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-12 18:58:20 UTC
# Updated By: jdyer-nuvodia
# Version: 4.4
# Additional Info: Fixed module loading and logging initialization
# =============================================================================

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

# Initialize paths
$LogFile = Join-Path -Path $PSScriptRoot -ChildPath "Create-TestDomainController.log"
$BaseModulePath = "C:\Users\jdyer\OneDrive - Nuvodia\Documents\GitHub\Scripts\Azure\Create-TestDomainController\Modules"

# Initialize logging with timestamp
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
        [string]$ModuleName
    )
    
    $fullPath = $ModulePath
    
    if (-not (Test-ModulePath -Path $fullPath -ModuleName $ModuleName)) {
        return $false
    }
    
    try {
        $manifestPath = Join-Path -Path $fullPath -ChildPath "$ModuleName.psd1"
        Import-Module -Name $manifestPath -Force -ErrorAction Stop
        # Initialize module logging after import
        & "$ModuleName\Set-DCLogFile" -Path $LogFile
        Write-Log "Successfully imported module: $ModuleName from $manifestPath" -Level INFO
        return $true
    }
    catch {
        Write-Log ("Failed to import module {0}: {1}" -f $ModuleName, $_) -Level ERROR
        return $false
    }
}

# Verify and import required modules
try {
    $modules = @(
        @{
            Path = Join-Path -Path $BaseModulePath -ChildPath "Configuration"
            Name = "DC-Configuration"
        },
        @{
            Path = Join-Path -Path $BaseModulePath -ChildPath "Validation"
            Name = "DC-Validation"
        },
        @{
            Path = Join-Path -Path $BaseModulePath -ChildPath "Deployment"
            Name = "DC-Deployment"
        }
    )
    
    $failedImports = 0
    
    foreach ($module in $modules) {
        if (-not (Import-RequiredModule -ModulePath $module.Path -ModuleName $module.Name)) {
            $failedImports++
        }
    }
    
    if ($failedImports -gt 0) {
        throw "Failed to import $failedImports required module(s). Please check the log for details."
    }
} catch {
    Write-Log "Critical error during module import: $_" -Level ERROR
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