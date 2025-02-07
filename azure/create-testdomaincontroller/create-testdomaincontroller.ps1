# =============================================================================
# Script: Create-TestDomainController.ps1
# Created: 2025-02-07 21:21:53 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-07 22:45:13 UTC
# Updated By: jdyer-nuvodia
# Version: 2.0
# Purpose: Creates a test domain controller in Azure with existence checks
# =============================================================================

# Function to write timestamped log messages
function Write-Log {
    param($Message)
    $timestamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
    Write-Host "[$timestamp UTC] $Message"
}

# Function to check if a resource group exists
function Test-ResourceGroupExists {
    param($ResourceGroupName)
    try {
        $null = Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

# Function to check if a storage account exists
function Test-StorageAccountExists {
    param($StorageAccountName, $ResourceGroupName)
    try {
        $null = Get-AzStorageAccount -ResourceGroupName $ResourceGroupName -Name $StorageAccountName -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

# Function to check if a virtual network exists
function Test-VNetExists {
    param($VNetName, $ResourceGroupName)
    try {
        $null = Get-AzVirtualNetwork -Name $VNetName -ResourceGroupName $ResourceGroupName -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

# Function to check if a subnet exists
function Test-SubnetExists {
    param($VNetName, $SubnetName, $ResourceGroupName)
    try {
        $vnet = Get-AzVirtualNetwork -Name $VNetName -ResourceGroupName $ResourceGroupName
        $null = $vnet.Subnets | Where-Object { $_.Name -eq $SubnetName }
        return $true
    }
    catch {
        return $false
    }
}

# Function to check if a public IP exists
function Test-PublicIPExists {
    param($PublicIPName, $ResourceGroupName)
    try {
        $null = Get-AzPublicIpAddress -Name $PublicIPName -ResourceGroupName $ResourceGroupName -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

# Function to check if a network interface exists
function Test-NetworkInterfaceExists {
    param($NICName, $ResourceGroupName)
    try {
        $null = Get-AzNetworkInterface -Name $NICName -ResourceGroupName $ResourceGroupName -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

# Function to check if a virtual machine exists
function Test-VMExists {
    param($VMName, $ResourceGroupName)
    try {
        $null = Get-AzVM -Name $VMName -ResourceGroupName $ResourceGroupName -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

# Script Parameters
$resourceGroupName = "TestDC-RG"
$location = "eastus"
$storageAccountName = "jbteststorage0"
$vnetName = "TestDC-VNet"
$subnetName = "TestDC-Subnet"
$publicIPName = "TestDC-PIP"
$nicName = "TestDC-NIC"
$vmName = "TestDC01"

Write-Log "Importing Azure modules..."
Import-Module Az.Accounts
Import-Module Az.Resources
Import-Module Az.Storage
Import-Module Az.Network
Import-Module Az.Compute

# Check and create Resource Group
Write-Log "Setting up Resource Group..."
if (-not (Test-ResourceGroupExists -ResourceGroupName $resourceGroupName)) {
    try {
        New-AzResourceGroup -Name $resourceGroupName -Location $location
        Write-Log "Resource Group created successfully"
    }
    catch {
        Write-Log "ERROR: Failed to create Resource Group: $_"
        exit 1
    }
}
else {
    Write-Log "Resource Group already exists - skipping creation"
}

# Check and create Storage Account
Write-Log "Configuring Storage..."
if (-not (Test-StorageAccountExists -StorageAccountName $storageAccountName -ResourceGroupName $resourceGroupName)) {
    Write-Log "Checking storage account name availability..."
    $storageNameAvailable = Get-AzStorageAccountNameAvailability -Name $storageAccountName
    
    if ($storageNameAvailable.NameAvailable) {
        try {
            New-AzStorageAccount -ResourceGroupName $resourceGroupName `
                                -Name $storageAccountName `
                                -Location $location `
                                -SkuName Standard_LRS
            Write-Log "Storage Account created successfully"
        }
        catch {
            Write-Log "ERROR: Failed to create Storage Account: $_"
            exit 1
        }
    }
    else {
        Write-Log "ERROR: Storage account name '$storageAccountName' is not available. Reason: $($storageNameAvailable.Message)"
        exit 1
    }
}
else {
    Write-Log "Storage Account already exists - skipping creation"
}

# Check and create Virtual Network
Write-Log "Setting up Virtual Network..."
if (-not (Test-VNetExists -VNetName $vnetName -ResourceGroupName $resourceGroupName)) {
    try {
        $subnetConfig = New-AzVirtualNetworkSubnetConfig -Name $subnetName -AddressPrefix "10.0.0.0/24"
        New-AzVirtualNetwork -ResourceGroupName $resourceGroupName `
                            -Name $vnetName `
                            -Location $location `
                            -AddressPrefix "10.0.0.0/16" `
                            -Subnet $subnetConfig
        Write-Log "Virtual Network created successfully"
    }
    catch {
        Write-Log "ERROR: Failed to create Virtual Network: $_"
        exit 1
    }
}
else {
    Write-Log "Virtual Network already exists - skipping creation"
}

# Check and create Public IP
Write-Log "Creating Public IP..."
if (-not (Test-PublicIPExists -PublicIPName $publicIPName -ResourceGroupName $resourceGroupName)) {
    try {
        New-AzPublicIpAddress -Name $publicIPName `
                             -ResourceGroupName $resourceGroupName `
                             -Location $location `
                             -AllocationMethod Dynamic
        Write-Log "Public IP created successfully"
    }
    catch {
        Write-Log "ERROR: Failed to create Public IP: $_"
        exit 1
    }
}
else {
    Write-Log "Public IP already exists - skipping creation"
}

# Check and create Network Interface
Write-Log "Setting up Network Interface..."
if (-not (Test-NetworkInterfaceExists -NICName $nicName -ResourceGroupName $resourceGroupName)) {
    try {
        $vnet = Get-AzVirtualNetwork -Name $vnetName -ResourceGroupName $resourceGroupName
        $subnet = $vnet.Subnets[0]
        $pip = Get-AzPublicIpAddress -Name $publicIPName -ResourceGroupName $resourceGroupName
        
        New-AzNetworkInterface -Name $nicName `
                              -ResourceGroupName $resourceGroupName `
                              -Location $location `
                              -SubnetId $subnet.Id `
                              -PublicIpAddressId $pip.Id
        Write-Log "Network Interface created successfully"
    }
    catch {
        Write-Log "ERROR: Failed to create Network Interface: $_"
        exit 1
    }
}
else {
    Write-Log "Network Interface already exists - skipping creation"
}

# Check and create Virtual Machine
Write-Log "Creating Virtual Machine..."
if (-not (Test-VMExists -VMName $vmName -ResourceGroupName $resourceGroupName)) {
    try {
        $nic = Get-AzNetworkInterface -Name $nicName -ResourceGroupName $resourceGroupName
        
        $vmConfig = New-AzVMConfig -VMName $vmName -VMSize "Standard_DS2_v2"
        $vmConfig = Set-AzVMOperatingSystem -VM $vmConfig `
                                          -Windows `
                                          -ComputerName $vmName `
                                          -Credential (Get-Credential) `
                                          -ProvisionVMAgent
        
        $vmConfig = Add-AzVMNetworkInterface -VM $vmConfig -Id $nic.Id
        $vmConfig = Set-AzVMSourceImage -VM $vmConfig `
                                      -PublisherName "MicrosoftWindowsServer" `
                                      -Offer "WindowsServer" `
                                      -Skus "2019-Datacenter" `
                                      -Version "latest"
        
        New-AzVM -ResourceGroupName $resourceGroupName -Location $location -VM $vmConfig
        Write-Log "Virtual Machine created successfully"
    }
    catch {
        Write-Log "ERROR: Failed to create Virtual Machine: $_"
        exit 1
    }
}
else {
    Write-Log "Virtual Machine already exists - skipping creation"
}

Write-Log "Script completed successfully"