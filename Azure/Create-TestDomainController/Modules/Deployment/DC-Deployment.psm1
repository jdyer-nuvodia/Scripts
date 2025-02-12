# =============================================================================
# Script: DC-Deployment.psm1
# Created: 2025-02-11 23:45:10 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-12 20:51:01 UTC
# Updated By: jdyer-nuvodia
# Version: 2.2
# Additional Info: Complete module implementation with all required functions
# =============================================================================

<#
.SYNOPSIS
    Provides deployment functionality for Azure Domain Controller creation
.DESCRIPTION
    This module contains functions for deploying and configuring domain controllers in Azure,
    including VM creation, network configuration, domain services installation, and auto-shutdown scheduling.
    
    Required Azure PowerShell modules:
    - Az.Compute
    - Az.Network
    - Az.Resources
    - Az.Storage
    
    Functions included:
    - New-DCEnvironment: Creates the complete DC environment
    - Enable-AzVmAutoShutdown: Configures VM auto-shutdown schedule
    - Install-DomainServices: Installs and configures AD DS roles
    - Set-DCNetworkConfiguration: Configures network settings
    - Initialize-DCStorage: Sets up required storage configuration
#>

function New-DCEnvironment {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [PSCustomObject]$Config
    )
    try {
        # Validate Azure connection
        $context = Get-AzContext
        if (-not $context) {
            throw "Not connected to Azure. Please run Connect-AzAccount first."
        }
        # Create Resource Group if it doesn't exist
        if (-not (Get-AzResourceGroup -Name $Config.ResourceGroupName -ErrorAction SilentlyContinue)) {
            New-AzResourceGroup -Name $Config.ResourceGroupName -Location $Config.Location
        }
        # Initialize network configuration
        $networkConfig = Set-DCNetworkConfiguration -Config $Config
        # Initialize storage configuration
        $storageConfig = Initialize-DCStorage -Config $Config
        # Create the VM
        $vmParams = @{
            ResourceGroupName = $Config.ResourceGroupName
            Location = $Config.Location
            Name = $Config.VmName
            Size = $Config.VMSize
            NetworkInterface = $networkConfig.NetworkInterface
            Credential = New-Object System.Management.Automation.PSCredential ($Config.AdminUsername, (ConvertTo-SecureString $Config.AdminPassword -AsPlainText -Force))
            ImageName = "Win2022Datacenter"
            OSDiskType = "Premium_LRS"
            DataDiskSizeInGb = 128
        }
        $vm = New-AzVM @vmParams
        # Install Domain Services
        Install-DomainServices -Config $Config -VM $vm
        # Configure auto-shutdown
        Enable-AzVmAutoShutdown -ResourceGroupName $Config.ResourceGroupName -VmName $Config.VmName
        return $vm
    }
    catch {
        Write-Error "Failed to create DC environment: $_"
        throw
    }
}

function Set-DCNetworkConfiguration {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [PSCustomObject]$Config
    )
    try {
        # Create or get Virtual Network
        $vnetParams = @{
            Name = $Config.VnetName
            ResourceGroupName = $Config.ResourceGroupName
            Location = $Config.Location
            AddressPrefix = "10.0.0.0/16"
        }
        $vnet = Get-AzVirtualNetwork -Name $Config.VnetName -ResourceGroupName $Config.ResourceGroupName -ErrorAction SilentlyContinue
        if (-not $vnet) {
            $vnet = New-AzVirtualNetwork @vnetParams
        }
        # Create or get Subnet
        $subnetConfig = @{
            Name = $Config.SubnetName
            AddressPrefix = "10.0.0.0/24"
            VirtualNetwork = $vnet
        }
        $subnet = Get-AzVirtualNetworkSubnetConfig -VirtualNetwork $vnet -Name $Config.SubnetName -ErrorAction SilentlyContinue
        if (-not $subnet) {
            $subnet = Add-AzVirtualNetworkSubnetConfig @subnetConfig
            $vnet | Set-AzVirtualNetwork
        }
        # Create Network Interface
        $nicParams = @{
            Name = "$($Config.VmName)-nic"
            ResourceGroupName = $Config.ResourceGroupName
            Location = $Config.Location
            SubnetId = $subnet.Id
        }
        $nic = New-AzNetworkInterface @nicParams
        return @{
            VirtualNetwork = $vnet
            Subnet = $subnet
            NetworkInterface = $nic
        }
    }
    catch {
        Write-Error "Failed to configure network: $_"
        throw
    }
}

function Initialize-DCStorage {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [PSCustomObject]$Config
    )
    try {
        # Create storage account for diagnostics
        $saName = ($Config.VmName + "diag").ToLower() -replace '[^a-z0-9]', ''
        $storageParams = @{
            ResourceGroupName = $Config.ResourceGroupName
            Name = $saName
            Location = $Config.Location
            SkuName = "Standard_LRS"
            Kind = "StorageV2"
        }
        $storageAccount = New-AzStorageAccount @storageParams
        return @{
            StorageAccount = $storageAccount
            StorageAccountName = $saName
        }
    }
    catch {
        Write-Error "Failed to initialize storage: $_"
        throw
    }
}

function Install-DomainServices {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [PSCustomObject]$Config,
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [Microsoft.Azure.Commands.Compute.Models.PSVirtualMachine]$VM
    )
    try {
        # Prepare custom script extension parameters
        $scriptContent = @'
Install-WindowsFeature -Name AD-Domain-Services -IncludeManagementTools
Install-ADDSForest `
    -DomainName "test.local" `
    -InstallDns `
    -Force `
    -SafeModeAdministratorPassword (ConvertTo-SecureString "TemporaryPassword123!" -AsPlainText -Force)
'@
        $scriptParams = @{
            ResourceGroupName = $Config.ResourceGroupName
            VMName = $Config.VmName
            Location = $Config.Location
            Name = "InstallADDS"
            TypeHandlerVersion = "1.10"
            Publisher = "Microsoft.Compute"
            FileUri = "script.ps1"
            Run = $scriptContent
            Argument = ""
        }
        Set-AzVMCustomScriptExtension @scriptParams
    }
    catch {
        Write-Error "Failed to install domain services: $_"
        throw
    }
}

function Enable-AzVmAutoShutdown {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ResourceGroupName,
        [Parameter(Mandatory = $true)]
        [string]$VmName,
        [Parameter(Mandatory = $false)]
        [string]$ScheduledShutdownTime = "2200",
        [Parameter(Mandatory = $false)]
        [string]$TimeZoneId = "US Mountain Standard Time"
    )
    try {
        if (-not (Get-Module -ListAvailable -Name Az.Compute)) {
            throw "The Az.Compute module is required for this operation. Please install it using: Install-Module -Name Az -Scope CurrentUser"
        }
        $vm = Get-AzVM -ResourceGroupName $ResourceGroupName -Name $VmName
        if (-not $vm) {
            throw "VM '$VmName' not found in resource group '$ResourceGroupName'"
        }
        $scheduleParams = @{
            Location = $vm.Location
            Name = "shutdown-computevm-$VmName"
            ResourceGroupName = $ResourceGroupName
            TargetResourceId = $vm.Id
            DailyRecurrence = "true"
            TimeZoneId = $TimeZoneId
            Time = $ScheduledShutdownTime
            NotificationSettings = @{
                Enabled = $false
            }
        }
        $schedule = New-AzAutomationSchedule @scheduleParams
        Write-Verbose "Auto-shutdown schedule created successfully for VM '$VmName'"
        return $schedule
    }
    catch {
        Write-Error "Failed to enable auto-shutdown for VM '$VmName': $_"
        throw
    }
}

# Export module members
Export-ModuleMember -Function @(
    'New-DCEnvironment',
    'Enable-AzVmAutoShutdown',
    'Install-DomainServices',
    'Set-DCNetworkConfiguration',
    'Initialize-DCStorage'
)