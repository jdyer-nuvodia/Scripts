<# 
.SYNOPSIS
    Creates a test domain controller as a Trusted Launch VM in Azure.
.DESCRIPTION
    This script provisions a domain controller VM configured as a Trusted Launch VM in Azure.
    It creates or verifies a resource group, storage account, network resources (virtual network, subnet, public IP,
    network security group), and provisions a Windows Server VM with Trusted Launch security features (Secure Boot and vTPM enabled).
    Additionally, it preserves explicitly defined variable values exactly as found in the repository.
    Note: The $domainName variable is provided for future domain join or configuration purposes.
.PARAMETER resourceGroupName
    The name of the resource group where the VM and related resources will be created.
.PARAMETER location
    The Azure region (location) to deploy the resources.
.PARAMETER vmName
    The name of the VM to create.
.PARAMETER VMSize
    The size of the VM (e.g., "Standard_DS2_v2").
.PARAMETER vnetName
    The virtual network name for the VM.
.PARAMETER subnetName
    The name of the subnet within the virtual network.
.PARAMETER adminUsername
    The administrator username for the VM.
.PARAMETER adminPassword
    The administrator password for the VM.
.EXAMPLE
    PS C:\> .\Create-TestDomainController.ps1 -resourceGroupName "JB-TEST-RG2" `
           -location "westus2" -vmName "JB-TEST-DC01" -VMSize "Standard_DS2_v2" `
           -vnetName "JB-TEST-VNET" -subnetName "JB-TEST-SUBNET1" `
           -adminUsername "jbadmin" -adminPassword "TS=pGxB~8m^A~WH^[yB8"
#>

# Explicitly defined variables (values preserved from the repository)
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

param(
    [Parameter(Mandatory = $true)]
    [string]$resourceGroupName = $resourceGroupName,
    [Parameter(Mandatory = $true)]
    [string]$location = $location,
    [Parameter(Mandatory = $true)]
    [string]$vmName = $vmName,
    [Parameter(Mandatory = $true)]
    [string]$VMSize = "Standard_DS2_v2",
    [Parameter(Mandatory = $true)]
    [string]$vnetName = $vnetName,
    [Parameter(Mandatory = $true)]
    [string]$subnetName = $subnetName,
    [Parameter(Mandatory = $true)]
    [string]$adminUsername = $adminUsername,
    [Parameter(Mandatory = $true)]
    [string]$adminPassword = $adminPassword
)

# ---------------------------------------------------------------------------
# Load Az Modules and Verify
# ---------------------------------------------------------------------------
try {
    if (-not (Get-Module -ListAvailable -Name Az.Compute)) { throw "Az.Compute module not found" }
    Import-Module Az.Compute -ErrorAction Stop
    Import-Module Az.Network -ErrorAction Stop
    Import-Module Az.Resources -ErrorAction Stop
    Write-Host "Successfully loaded required Az modules."
}
catch {
    Write-Error "Failed to load required Az modules: $_"
    exit 1
}

# ---------------------------------------------------------------------------
# Create/Verify Resource Group
# ---------------------------------------------------------------------------
try {
    if (-not (Get-AzResourceGroup -Name $resourceGroupName -ErrorAction SilentlyContinue)) {
        Write-Host "Creating resource group '$resourceGroupName' in location '$location'..."
        New-AzResourceGroup -Name $resourceGroupName -Location $location -ErrorAction Stop
        Write-Host "Resource group '$resourceGroupName' created successfully."
    }
    else {
        Write-Host "Resource group '$resourceGroupName' already exists."
    }
}
catch {
    Write-Error "Error during resource group creation or verification: $_"
    exit 1
}

# ---------------------------------------------------------------------------
# Create/Verify Storage Account
# ---------------------------------------------------------------------------
try {
    if (-not (Get-AzStorageAccount -ResourceGroupName $resourceGroupName -Name $storageAccountName -ErrorAction SilentlyContinue)) {
        Write-Host "Creating Storage Account '$storageAccountName'..."
        New-AzStorageAccount -ResourceGroupName $resourceGroupName -Name $storageAccountName -Location $location -SkuName Standard_LRS -Kind StorageV2 -ErrorAction Stop
        Write-Host "Storage Account '$storageAccountName' created successfully."
    }
    else {
        Write-Host "Storage Account '$storageAccountName' already exists."
    }
}
catch {
    Write-Error "Error creating Storage Account: $_"
    exit 1
}

# ---------------------------------------------------------------------------
# Configure Virtual Network and Subnet
# ---------------------------------------------------------------------------
try {
    $VNet = Get-AzVirtualNetwork -Name $vnetName -ResourceGroupName $resourceGroupName -ErrorAction Stop
    $Subnet = Get-AzVirtualNetworkSubnetConfig -Name $subnetName -VirtualNetwork $VNet -ErrorAction Stop
    Write-Host "Virtual network '$vnetName' and subnet '$subnetName' fetched successfully."
}
catch {
    Write-Error "Error fetching virtual network/subnet: $_"
    exit 1
}

# ---------------------------------------------------------------------------
# Create Public IP Address
# ---------------------------------------------------------------------------
try {
    $PublicIP = New-AzPublicIpAddress -Name $publicIpName -ResourceGroupName $resourceGroupName `
        -Location $location -AllocationMethod Dynamic -ErrorAction Stop
    Write-Host "Public IP address '$($PublicIP.Name)' created successfully."
}
catch {
    Write-Error "Error creating Public IP Address: $_"
    exit 1
}

# ---------------------------------------------------------------------------
# Create Network Security Group (NSG)
# ---------------------------------------------------------------------------
try {
    if (-not (Get-AzNetworkSecurityGroup -Name $nsgName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue)) {
        Write-Host "Creating Network Security Group '$nsgName'..."
        $nsg = New-AzNetworkSecurityGroup -Name $nsgName -ResourceGroupName $resourceGroupName -Location $location -ErrorAction Stop
        Write-Host "NSG '$nsgName' created successfully."
    }
    else {
        $nsg = Get-AzNetworkSecurityGroup -Name $nsgName -ResourceGroupName $resourceGroupName -ErrorAction Stop
        Write-Host "NSG '$nsgName' already exists."
    }
}
catch {
    Write-Error "Error creating or retrieving NSG: $_"
    exit 1
}

# ---------------------------------------------------------------------------
# Create Network Interface
# ---------------------------------------------------------------------------
try {
    $NIC = New-AzNetworkInterface -Name "$vmName-nic" -ResourceGroupName $resourceGroupName `
        -Location $location -SubnetId $Subnet.Id -PublicIpAddressId $PublicIP.Id -ErrorAction Stop
    Write-Host "Network interface '$($NIC.Name)' created successfully."
}
catch {
    Write-Error "Error creating Network Interface: $_"
    exit 1
}

# ---------------------------------------------------------------------------
# Associate NSG with Network Interface
# ---------------------------------------------------------------------------
try {
    Set-AzNetworkInterface -NetworkInterface $NIC -NetworkSecurityGroup $nsg -ErrorAction Stop
    Write-Host "Associated NSG '$nsgName' with NIC '$($NIC.Name)'."
}
catch {
    Write-Error "Error associating NSG with NIC: $_"
    exit 1
}

# ---------------------------------------------------------------------------
# Define Trusted Launch Settings
# ---------------------------------------------------------------------------
$TrustedLaunchProfile = @{
    SecurityType = "TrustedLaunch"
    UefiSettings = @{
        SecureBootEnabled = $true
        VtpmEnabled       = $true
    }
}
Write-Host "Trusted Launch settings defined."

# ---------------------------------------------------------------------------
# Create VM Configuration with Extended Settings
# ---------------------------------------------------------------------------
try {
    $SecureCredential = New-Object System.Management.Automation.PSCredential(
        $adminUsername, (ConvertTo-SecureString $adminPassword -AsPlainText -Force)
    )
    $VMConfig = New-AzVMConfig -VMName $vmName -VMSize $VMSize -ErrorAction Stop | `
        Set-AzVMOperatingSystem -Windows -ComputerName $vmName -Credential $SecureCredential `
            -ProvisionVMAgent -EnableAutoUpdate -ErrorAction Stop | `
        Set-AzVMSourceImage -PublisherName MicrosoftWindowsServer -Offer WindowsServer `
            -Skus 2019-Datacenter -Version latest -ErrorAction Stop | `
        Add-AzVMNetworkInterface -Id $NIC.Id -ErrorAction Stop

    # Apply Trusted Launch settings without altering explicitly defined variable values or extra functionality
    $VMConfig.AdditionalCapabilities = @{ UefiSettings = $TrustedLaunchProfile.UefiSettings }
    $VMConfig.SecurityProfile = @{ SecurityType = $TrustedLaunchProfile.SecurityType }
    Write-Host "VM configuration prepared with Trusted Launch settings."
}
catch {
    Write-Error "Error configuring VM settings: $_"
    exit 1
}

# ---------------------------------------------------------------------------
# Create the Trusted Launch VM
# ---------------------------------------------------------------------------
try {
    Write-Host "Initiating creation of Trusted Launch VM '$vmName'..."
    New-AzVM -ResourceGroupName $resourceGroupName -Location $location -VM $VMConfig -ErrorAction Stop
    Write-Host "Trusted Launch VM '$vmName' created successfully."
    Write-Host "Note: The provided domain name '$domainName' is ready for future domain join configuration."
}
catch {
    Write-Error "Error during VM creation: $_"
    exit 1
}