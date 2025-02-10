# =============================================================================
# Script: Create-TestDomainController.ps1
# Created: 2025-02-10 22:50:04 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-10 22:50:04 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2
# Additional Info: Fixed Public IP SKU and allocation method
# =============================================================================

<# 
.SYNOPSIS
    Creates a test domain controller as a Trusted Launch VM in Azure.
.DESCRIPTION
    This script provisions a domain controller VM configured as a Trusted Launch VM in Azure.
    It creates or verifies a resource group, storage account, network resources (virtual network, subnet, public IP,
    network security group), and provisions a Windows Server VM with Trusted Launch security features (Secure Boot and vTPM enabled).
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
.EXAMPLE
    PS C:\> .\Create-TestDomainController.ps1 -resourceGroupName 'JB-TEST-RG2' `
           -location 'westus2' -vmName 'JB-TEST-DC01' -VMSize 'Standard_DS2_v2' `
           -vnetName 'JB-TEST-VNET' -subnetName 'JB-TEST-SUBNET1' `
           -adminUsername 'jbadmin' -adminPassword 'TS-pGxB~8m^A~WH^[yB8'
#>

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
    [string]$adminPassword
)

# Start logging
$LogFile = Join-Path $PSScriptRoot "Create-TestDomainController.log"
function Write-Log {
    param($Message)
    $LogMessage = "[$([DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))] $Message"
    Add-Content -Path $LogFile -Value $LogMessage
    Write-Host $LogMessage
}

# Clear log file
Set-Content -Path $LogFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Log file reset. New log starting."

# Explicitly defined default variables
$DefaultResourceGroupName    = 'JB-TEST-RG2'
$DefaultLocation            = 'westus2'
$DefaultStorageAccountName  = 'jbteststorage0'
$DefaultVnetName           = 'JB-TEST-VNET'
$DefaultSubnetName         = 'JB-TEST-SUBNET1'
$DefaultVmName             = 'JB-TEST-DC01'
$DefaultAdminUsername      = 'jbadmin'
$DefaultAdminPassword      = 'TS-pGxB~8m^A~WH^[yB8'
$DefaultDomainName         = 'JB-TEST.local'
$DefaultPublicIpName       = "$DefaultVmName-PUBIP"
$DefaultNsgName           = 'JB-TEST-NSG'
$DefaultVnetAddressSpace   = '10.0.0.0/16'
$DefaultSubnetAddressSpace = '10.0.1.0/24'

# If parameters are not provided, assign defaults
if (-not $resourceGroupName) { $resourceGroupName = $DefaultResourceGroupName }
if (-not $location)          { $location = $DefaultLocation }
if (-not $vmName)            { $vmName = $DefaultVmName }
if (-not $VMSize)            { $VMSize = 'Standard_DS2_v2' }
if (-not $vnetName)          { $vnetName = $DefaultVnetName }
if (-not $subnetName)        { $subnetName = $DefaultSubnetName }
if (-not $adminUsername)     { $adminUsername = $DefaultAdminUsername }
if (-not $adminPassword)     { $adminPassword = $DefaultAdminPassword }

try {
    # Verify Az modules are loaded
    $requiredModules = @('Az.Accounts', 'Az.Resources', 'Az.Network', 'Az.Storage', 'Az.Compute')
    foreach ($module in $requiredModules) {
        if (!(Get-Module -Name $module -ListAvailable)) {
            throw "Required module $module is not installed"
        }
    }
    Write-Log "[INFO] Successfully loaded required Az modules."

    # Check/Create Resource Group
    $rg = Get-AzResourceGroup -Name $resourceGroupName -ErrorAction SilentlyContinue
    if (!$rg) {
        Write-Log "[INFO] Creating resource group '$resourceGroupName'..."
        New-AzResourceGroup -Name $resourceGroupName -Location $location
    } else {
        Write-Log "[INFO] Resource group '$resourceGroupName' already exists."
    }

    # Create Storage Account if it doesn't exist
    $storageAccount = Get-AzStorageAccount -ResourceGroupName $resourceGroupName -Name $DefaultStorageAccountName -ErrorAction SilentlyContinue
    if (!$storageAccount) {
        Write-Log "[INFO] Creating Storage Account '$DefaultStorageAccountName'..."
        $storageAccount = New-AzStorageAccount -ResourceGroupName $resourceGroupName `
            -Name $DefaultStorageAccountName `
            -Location $location `
            -SkuName Standard_LRS `
            -Kind StorageV2
        Write-Log "[INFO] Storage Account '$DefaultStorageAccountName' created successfully."
    }

    # Create Network Security Group
    $nsg = Get-AzNetworkSecurityGroup -Name $DefaultNsgName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue
    if (!$nsg) {
        Write-Log "[INFO] Creating Network Security Group '$DefaultNsgName'..."
		        $nsgRules = @(
            @{
                Name = 'AllowRDP'
                Protocol = 'Tcp'
                SourcePortRange = '*'
                DestinationPortRange = '3389'
                SourceAddressPrefix = '*'
                DestinationAddressPrefix = '*'
                Access = 'Allow'
                Priority = 100
                Direction = 'Inbound'
            }
        )
        
        $nsg = New-AzNetworkSecurityGroup -ResourceGroupName $resourceGroupName `
            -Location $location `
            -Name $DefaultNsgName
        
        foreach ($rule in $nsgRules) {
            Add-AzNetworkSecurityRuleConfig @rule -NetworkSecurityGroup $nsg
        }
        $nsg | Set-AzNetworkSecurityGroup
        Write-Log "[INFO] Network Security Group created successfully."
    }

    # Create Virtual Network and Subnet if they don't exist
    $vnet = Get-AzVirtualNetwork -Name $vnetName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue
    if (!$vnet) {
        Write-Log "[INFO] Creating Virtual Network '$vnetName'..."
        $subnetConfig = New-AzVirtualNetworkSubnetConfig -Name $subnetName `
            -AddressPrefix $DefaultSubnetAddressSpace `
            -NetworkSecurityGroup $nsg

        $vnet = New-AzVirtualNetwork -ResourceGroupName $resourceGroupName `
            -Location $location `
            -Name $vnetName `
            -AddressPrefix $DefaultVnetAddressSpace `
            -Subnet $subnetConfig
        Write-Log "[INFO] Virtual Network and Subnet created successfully."
    }

    # Create Public IP with Standard SKU and Static allocation
    $publicIp = Get-AzPublicIpAddress -Name $DefaultPublicIpName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue
    if (!$publicIp) {
        Write-Log "[INFO] Creating Public IP '$DefaultPublicIpName'..."
        $publicIp = New-AzPublicIpAddress -ResourceGroupName $resourceGroupName `
            -Location $location `
            -Name $DefaultPublicIpName `
            -Sku Standard `
            -AllocationMethod Static
        Write-Log "[INFO] Public IP created successfully."
    }

    # Create NIC
    $nicName = "$vmName-NIC"
    $subnet = Get-AzVirtualNetwork -Name $vnetName -ResourceGroupName $resourceGroupName |
        Get-AzVirtualNetworkSubnetConfig -Name $subnetName
    $nic = New-AzNetworkInterface -Name $nicName -ResourceGroupName $resourceGroupName `
        -Location $location -SubnetId $subnet.Id -PublicIpAddressId $publicIp.Id
    Write-Log "[INFO] Network interface created successfully."

    # Create VM configuration
    $vmConfig = New-AzVMConfig -VMName $vmName -VMSize $VMSize |
        Set-AzVMOperatingSystem -Windows -ComputerName $vmName -Credential $credential |
        Set-AzVMSourceImage -PublisherName 'MicrosoftWindowsServer' `
            -Offer 'WindowsServer' `
            -Skus '2022-Datacenter' `
            -Version latest |
        Add-AzVMNetworkInterface -Id $nic.Id
    
    # Enable Trusted Launch
    $vmConfig = Set-AzVMSecurityProfile -VM $vmConfig -SecurityType "TrustedLaunch"
    $vmConfig = Set-AzVMUefi -VM $vmConfig -EnableVtpm $true -EnableSecureBoot $true

    # Create the VM
    Write-Log "[INFO] Creating VM '$vmName'..."
    New-AzVM -ResourceGroupName $resourceGroupName -Location $location -VM $vmConfig
    Write-Log "[INFO] VM '$vmName' created successfully."

    Write-Log "[INFO] Domain Controller VM creation completed successfully."
} catch {
    Write-Log "[ERROR] An error occurred: $($_.Exception.Message)"
    throw
} finally {
    Write-Log "[INFO] Script execution completed."
}