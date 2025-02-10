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

# Explicitly defined default variables
$DefaultResourceGroupName    = 'JB-TEST-RG2'
$DefaultLocation             = 'westus2'
$DefaultStorageAccountName   = 'jbteststorage0'
$DefaultVnetName             = 'JB-TEST-VNET'
$DefaultSubnetName           = 'JB-TEST-SUBNET1'
$DefaultVmName               = 'JB-TEST-DC01'
$DefaultAdminUsername        = 'jbadmin'
$DefaultAdminPassword        = 'TS-pGxB~8m^A~WH^[yB8'
$DefaultDomainName           = 'JB-TEST.local'
$DefaultPublicIpName         = "$DefaultVmName-PUBIP"
$DefaultNsgName              = 'JB-TEST-NSG'

# Set up log file
$logFile = Join-Path -Path $PSScriptRoot -ChildPath 'Create-TestDomainController.log'
if (Test-Path $logFile) { Remove-Item $logFile -Force }
"[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Log file reset. New log starting." | Out-File -FilePath $logFile

# Logging function
function Write-Log {
    param (
        [string]$Message,
        [ValidateSet('INFO','ERROR')]
        [string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$timestamp] [$Level] $Message"
    Write-Host $entry
    $entry | Out-File -FilePath $logFile -Append
}

param (
    [Parameter(Mandatory = $true)]
    [string]$resourceGroupName,
    [Parameter(Mandatory = $true)]
    [string]$location,
    [Parameter(Mandatory = $true)]
    [string]$vmName,
    [Parameter(Mandatory = $true)]
    [string]$VMSize,
    [Parameter(Mandatory = $true)]
    [string]$vnetName,
    [Parameter(Mandatory = $true)]
    [string]$subnetName,
    [Parameter(Mandatory = $true)]
    [string]$adminUsername,
    [Parameter(Mandatory = $true)]
    [string]$adminPassword
)

# ---------------------------------------------------------------------------
# Load Az Modules and Verify
# ---------------------------------------------------------------------------
try {
    if (-not (Get-Module -ListAvailable -Name Az.Compute)) { throw 'Az.Compute module not found' }
    Import-Module Az.Compute -ErrorAction Stop
    Import-Module Az.Network -ErrorAction Stop
    Import-Module Az.Resources -ErrorAction Stop
    Write-Log 'Successfully loaded required Az modules.'
}
catch {
    Write-Log "Failed to load required Az modules: $_" 'ERROR'
    exit 1
}

# ---------------------------------------------------------------------------
# Create/Verify Resource Group
# ---------------------------------------------------------------------------
try {
    if (-not (Get-AzResourceGroup -Name $resourceGroupName -ErrorAction SilentlyContinue)) {
        Write-Log "Creating resource group '$resourceGroupName' in location '$location'..."
        New-AzResourceGroup -Name $resourceGroupName -Location $location -ErrorAction Stop
        Write-Log "Resource group '$resourceGroupName' created successfully."
    }
    else {
        Write-Log "Resource group '$resourceGroupName' already exists."
    }
}
catch {
    Write-Log "Error during resource group creation or verification: $_" 'ERROR'
    exit 1
}

# ---------------------------------------------------------------------------
# Create/Verify Storage Account
# ---------------------------------------------------------------------------
try {
    if (-not (Get-AzStorageAccount -ResourceGroupName $resourceGroupName -Name $DefaultStorageAccountName -ErrorAction SilentlyContinue)) {
        Write-Log "Creating Storage Account '$DefaultStorageAccountName'..."
        New-AzStorageAccount -ResourceGroupName $resourceGroupName -Name $DefaultStorageAccountName -Location $location -SkuName Standard_LRS -Kind StorageV2 -ErrorAction Stop
        Write-Log "Storage Account '$DefaultStorageAccountName' created successfully."
    }
    else {
        Write-Log "Storage Account '$DefaultStorageAccountName' already exists."
    }
}
catch {
    Write-Log "Error creating Storage Account: $_" 'ERROR'
    exit 1
}

# ---------------------------------------------------------------------------
# Configure Virtual Network and Subnet
# ---------------------------------------------------------------------------
try {
    $VNet = Get-AzVirtualNetwork -Name $vnetName -ResourceGroupName $resourceGroupName -ErrorAction Stop
    $Subnet = Get-AzVirtualNetworkSubnetConfig -Name $subnetName -VirtualNetwork $VNet -ErrorAction Stop
    Write-Log "Virtual network '$vnetName' and subnet '$subnetName' fetched successfully."
}
catch {
    Write-Log "Error fetching virtual network/subnet: $_" 'ERROR'
    exit 1
}

# ---------------------------------------------------------------------------
# Create Public IP Address
# ---------------------------------------------------------------------------
try {
    $PublicIP = New-AzPublicIpAddress -Name $DefaultPublicIpName -ResourceGroupName $resourceGroupName `
        -Location $location -AllocationMethod Dynamic -ErrorAction Stop
    Write-Log "Public IP address '$($PublicIP.Name)' created successfully."
}
catch {
    Write-Log "Error creating Public IP Address: $_" 'ERROR'
    exit 1
}

# ---------------------------------------------------------------------------
# Create Network Security Group (NSG)
# ---------------------------------------------------------------------------
try {
    if (-not (Get-AzNetworkSecurityGroup -Name $DefaultNsgName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue)) {
        Write-Log "Creating Network Security Group '$DefaultNsgName'..."
        $nsg = New-AzNetworkSecurityGroup -Name $DefaultNsgName -ResourceGroupName $resourceGroupName -Location $location -ErrorAction Stop
        Write-Log "NSG '$DefaultNsgName' created successfully."
    }
    else {
        $nsg = Get-AzNetworkSecurityGroup -Name $DefaultNsgName -ResourceGroupName $resourceGroupName -ErrorAction Stop
        Write-Log "NSG '$DefaultNsgName' already exists."
    }
}
catch {
    Write-Log "Error creating or retrieving NSG: $_" 'ERROR'
    exit 1
}

# ---------------------------------------------------------------------------
# Create Network Interface
# ---------------------------------------------------------------------------
try {
    $NIC = New-AzNetworkInterface -Name "$vmName-nic" -ResourceGroupName $resourceGroupName `
        -Location $location -SubnetId $Subnet.Id -PublicIpAddressId $PublicIP.Id -ErrorAction Stop
    Write-Log "Network interface '$($NIC.Name)' created successfully."
}
catch {
    Write-Log "Error creating Network Interface: $_" 'ERROR'
    exit 1
}

# ---------------------------------------------------------------------------
# Associate NSG with Network Interface
# ---------------------------------------------------------------------------
try {
    Set-AzNetworkInterface -NetworkInterface $NIC -NetworkSecurityGroup $nsg -ErrorAction Stop
    Write-Log "Associated NSG '$DefaultNsgName' with NIC '$($NIC.Name)'."
}
catch {
    Write-Log "Error associating NSG with NIC: $_" 'ERROR'
    exit 1
}

# ---------------------------------------------------------------------------
# Define Trusted Launch Settings
# ---------------------------------------------------------------------------
$TrustedLaunchProfile = @{
    SecurityType = 'TrustedLaunch'
    UefiSettings = @{
        SecureBootEnabled = $true
        VtpmEnabled       = $true
    }
}
Write-Log 'Trusted Launch settings defined.'

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
    Write-Log 'VM configuration prepared with Trusted Launch settings.'
}
catch {
    Write-Log "Error configuring VM settings: $_" 'ERROR'
    exit 1
}

# ---------------------------------------------------------------------------
# Create the Trusted Launch VM
# ---------------------------------------------------------------------------
try {
    Write-Log "Initiating creation of Trusted Launch VM '$vmName'..."
    New-AzVM -ResourceGroupName $resourceGroupName -Location $location -VM $VMConfig -ErrorAction Stop
    Write-Log "Trusted Launch VM '$vmName' created successfully."
    Write-Log "Note: The provided domain name '$DefaultDomainName' is ready for future domain join configuration."
}
catch {
    Write-Log "Error during VM creation: $_" 'ERROR'
    exit 1
}