# =============================================================================
# Script: Create-TestDomainController.ps1
# Created: 2025-02-07 21:21:53 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-09 15:42:57 UTC
# Updated By: jdyer-nuvodia
# Version: 2.5
# Purpose: Creates a test domain controller in Azure with existence checks,
#          error handling, logging, and an option for verbose output.
# =============================================================================

[CmdletBinding()]
Param()

# Enable strict mode and set error action preference
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Function to write timestamped log messages
function Write-Log {
    param($Message)
    $timestamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
    Write-Host "[$timestamp UTC] $Message"
}

Write-Log "Script execution started."
Write-Verbose "Verbose mode activated."

# Import required Azure modules with error checking
try {
    Write-Log "Importing required Azure modules..."
    Import-Module Az.Resources -ErrorAction Stop
    Import-Module Az.Compute -ErrorAction Stop
    Import-Module Az.Network -ErrorAction Stop
    Import-Module Az.Storage -ErrorAction Stop
    Write-Verbose "Azure modules imported successfully."
} catch {
    Write-Log "ERROR: Failed to import required Azure modules. $_"
    exit 1
}

# Script Parameters (defaults; these could be parameterized as needed)
$resourceGroupName    = "JB-TEST-RG2"
$location             = "westus2"
$storageAccountName   = "jbteststorage0"
$vnetName             = "JB-TEST-VNET"
$subnetName           = "JB-TEST-SUBNET1"
$vmName               = "JB-TEST-DC01"
$adminUsername        = "jbadmin"
$adminPassword        = "TS=pGxB~8m^A~WH^[yB8"
$domainName           = "JB-TEST.local"
$publicIpName         = "$vmName-PUBIP"
$nsgName              = "JB-TEST-NSG"

# Function: Test if a Resource Group exists
function Test-ResourceGroupExists {
    param($ResourceGroupName)
    try {
        Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction Stop | Out-Null
        return $true
    } catch {
        return $false
    }
}

# Begin resource creation steps with error checking and logging

# 1. Check and create Resource Group if missing
try {
    Write-Log "Checking for resource group '$resourceGroupName'..."
    if (-not (Test-ResourceGroupExists -ResourceGroupName $resourceGroupName)) {
        Write-Log "Resource group '$resourceGroupName' not found. Creating resource group..."
        New-AzResourceGroup -Name $resourceGroupName -Location $location -ErrorAction Stop | Out-Null
        Write-Log "Resource group '$resourceGroupName' created."
    } else {
        Write-Log "Resource group '$resourceGroupName' exists."
    }
} catch {
    Write-Log "ERROR: Failed to verify or create resource group. $_"
    exit 1
}

# 1.1 Check and create Storage Account if missing
try {
    Write-Log "Checking for storage account '$storageAccountName'..."
    $storageAccount = Get-AzStorageAccount -Name $storageAccountName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue
    if (-not $storageAccount) {
        Write-Log "Storage account '$storageAccountName' not found. Creating storage account..."
        $storageAccount = New-AzStorageAccount -Name $storageAccountName -ResourceGroupName $resourceGroupName -Location $location -SkuName Standard_LRS -Kind StorageV2 -ErrorAction Stop
        Write-Log "Storage account '$storageAccountName' created."
    } else {
        Write-Log "Storage account '$storageAccountName' exists."
    }
} catch {
    Write-Log "ERROR: Failed to verify or create storage account. $_"
    exit 1
}

# 2. Create Public IP
try {
    Write-Log "Creating Public IP '$publicIpName'..."
    $publicIp = New-AzPublicIpAddress -Name $publicIpName -ResourceGroupName $resourceGroupName -Location $location -AllocationMethod Static -ErrorAction Stop
    Write-Log "Public IP '$publicIpName' created."
} catch {
    Write-Log "ERROR: Failed to create Public IP. $_"
    exit 1
}

# 3. Create Virtual Network and Subnet
try {
    Write-Log "Checking for Virtual Network '$vnetName'..."
    $vnet = Get-AzVirtualNetwork -Name $vnetName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue
    if (-not $vnet) {
        Write-Log "Virtual Network '$vnetName' not found. Creating Virtual Network with subnet '$subnetName'..."
        $subnetConfig = New-AzVirtualNetworkSubnetConfig -Name $subnetName -AddressPrefix "10.0.0.0/24"
        $vnet = New-AzVirtualNetwork -Name $vnetName -ResourceGroupName $resourceGroupName -Location $location -AddressPrefix "10.0.0.0/16" -Subnet $subnetConfig -ErrorAction Stop
        Write-Log "Virtual Network '$vnetName' created."
    } else {
        Write-Log "Virtual Network '$vnetName' exists."
    }
} catch {
    Write-Log "ERROR: Failed to verify or create Virtual Network. $_"
    exit 1
}

# 4. Create Network Interface
try {
    Write-Log "Creating Network Interface for VM '$vmName'..."
    $subnet = $vnet.Subnets | Where-Object { $_.Name -eq $subnetName }
    if (-not $subnet) {
        throw "Subnet '$subnetName' could not be found in Virtual Network '$vnetName'."
    }
    $nic = New-AzNetworkInterface -Name "$vmName-NIC" -ResourceGroupName $resourceGroupName -Location $location -SubnetId $subnet.Id -PublicIpAddressId $publicIp.Id -ErrorAction Stop
    Write-Log "Network Interface for VM '$vmName' created."
} catch {
    Write-Log "ERROR: Failed to create Network Interface. $_"
    exit 1
}

# 5. Create Virtual Machine
try {
    Write-Log "Creating Virtual Machine '$vmName'..."
    $securePassword = ConvertTo-SecureString $adminPassword -AsPlainText -Force
    $cred = New-Object System.Management.Automation.PSCredential ($adminUsername, $securePassword)
    $vmConfig = New-AzVMConfig -VMName $vmName -VMSize "Standard_DS1_v2" -ErrorAction Stop |
      Set-AzVMOperatingSystem -Windows -ComputerName $vmName -Credential $cred -ProvisionVMAgent -EnableAutoUpdate -ErrorAction Stop |
      Set-AzVMSourceImage -PublisherName "MicrosoftWindowsServer" -Offer "WindowsServer" -Skus "2019-Datacenter" -Version "latest" -ErrorAction Stop |
      Add-AzVMNetworkInterface -Id $nic.Id -Primary -ErrorAction Stop

    # Configure boot diagnostics to use the specified storage account using the correct cmdlet
    $vmConfig = Enable-AzVMBootDiagnostics -VM $vmConfig -ResourceGroupName $resourceGroupName -StorageAccountName $storageAccountName -ErrorAction Stop

    New-AzVM -ResourceGroupName $resourceGroupName -Location $location -VM $vmConfig -ErrorAction Stop | Out-Null
    Write-Log "Virtual Machine '$vmName' created successfully."
} catch {
    Write-Log "ERROR: Failed to create Virtual Machine. $_"
    exit 1
}

Write-Log "Script execution completed successfully."