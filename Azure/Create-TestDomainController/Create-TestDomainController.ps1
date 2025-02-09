# =============================================================================
# Script: Create-TestDomainController.ps1
# Created: 2025-02-07 21:21:53 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-09 16:30:00 UTC
# Updated By: jdyer-nuvodia
# Version: 2.11
# Purpose: Creates a test domain controller in Azure with existence checks,
#          error handling, NSG creation with an RDP rule on port 10443 and an explicit deny on port 3389,
#          overwrites existing resources automatically, logs execution via transcript, and backs up the current script.
# =============================================================================

[CmdletBinding()]
Param()

# Enable strict mode and set error action preference.
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Determine the folder where the script resides.
if ($PSScriptRoot) {
    $scriptFolder = $PSScriptRoot
} else {
    $scriptFolder = Get-Location
}

# Delete previous transcript log file if it exists.
$logPattern = "Create-TestDomainController-*.log"
$existingLogs = Get-ChildItem -Path $scriptFolder -Filter $logPattern -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
if ($existingLogs -and $existingLogs.Count -gt 0) {
    Write-Host "Deleting previous log file: $($existingLogs[0].FullName)"
    Remove-Item $existingLogs[0].FullName -Force
}

# Create a new transcript log file with a timestamp.
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$logFile = Join-Path $scriptFolder "Create-TestDomainController-$timestamp.log"
Start-Transcript -Path $logFile

# ---------------------------------------------------------------------------
# Backup mechanism: Delete previous backup and create a new backup of this script.
# ---------------------------------------------------------------------------
$backupPattern = "Create-TestDomainController_Backup-*.ps1"
$existingBackups = Get-ChildItem -Path $scriptFolder -Filter $backupPattern -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
if ($existingBackups -and $existingBackups.Count -gt 0) {
    Write-Host "Deleting previous backup file: $($existingBackups[0].FullName)"
    Remove-Item $existingBackups[0].FullName -Force
}
$backupTimestamp = Get-Date -Format "yyyyMMddHHmmss"
$backupFile = Join-Path $scriptFolder "Create-TestDomainController_Backup-$backupTimestamp.ps1"

# Determine the current script file path using a fallback mechanism.
$currentScriptPath = $MyInvocation.MyCommand.Path
if (-not $currentScriptPath) {
    # Fallback: Assume the script filename is "Create-TestDomainController.ps1" in the script folder.
    $expectedScriptName = "Create-TestDomainController.ps1"
    $fallbackScriptPath = Join-Path $scriptFolder $expectedScriptName
    if (Test-Path $fallbackScriptPath) {
        $currentScriptPath = $fallbackScriptPath
    }
    else {
        Write-Host "ERROR: Unable to determine the current script file path."
        Stop-Transcript
        exit 1
    }
}
try {
    Write-Host "Creating backup of the current script: $currentScriptPath"
    Copy-Item -Path $currentScriptPath -Destination $backupFile -Force
    Write-Host "Backup created successfully: $backupFile"
} catch {
    Write-Host "ERROR: Failed to create backup file. $_"
}

# Function to write timestamped log messages.
function Write-Log {
    param($Message)
    $timeStamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
    Write-Host "[$timeStamp UTC] $Message"
}

Write-Log "Script execution started."
Write-Verbose "Verbose mode activated."

# Import required Azure modules.
try {
    Write-Log "Importing required Azure modules..."
    Import-Module Az.Resources -ErrorAction Stop
    Import-Module Az.Compute -ErrorAction Stop
    Import-Module Az.Network -ErrorAction Stop
    Import-Module Az.Storage -ErrorAction Stop
    Write-Verbose "Azure modules imported successfully."
} catch {
    Write-Log "ERROR: Failed to import required Azure modules. $_"
    Stop-Transcript
    exit 1
}

# Script Parameters (defaults can be parameterized as needed)
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

# Function: Test if a Resource Group exists.
function Test-ResourceGroupExists {
    param($ResourceGroupName)
    try {
        Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction Stop | Out-Null
        return $true
    } catch {
        return $false
    }
}

# Begin resource creation steps with error handling.

# 1. Check and create Resource Group if missing.
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
    Stop-Transcript
    exit 1
}

# 1.1. Check and remove existing Storage Account then create a new one.
try {
    Write-Log "Checking for storage account '$storageAccountName'..."
    $existingStorage = Get-AzStorageAccount -Name $storageAccountName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue
    if ($existingStorage) {
        Write-Log "Storage account '$storageAccountName' already exists. Removing..."
        Remove-AzStorageAccount -Name $storageAccountName -ResourceGroupName $resourceGroupName -Force -Confirm:$false
        Write-Log "Storage account '$storageAccountName' removed."
    }
    Write-Log "Creating storage account '$storageAccountName'..."
    $storageAccount = New-AzStorageAccount -Name $storageAccountName -ResourceGroupName $resourceGroupName -Location $location -SkuName Standard_LRS -Kind StorageV2 -ErrorAction Stop
    Write-Log "Storage account '$storageAccountName' created."
} catch {
    Write-Log "ERROR: Failed to verify or create storage account. $_"
    Stop-Transcript
    exit 1
}

# 1.2. Check and remove existing Network Security Group (NSG) then create a new one with RDP rules.
try {
    Write-Log "Checking for Network Security Group '$nsgName'..."
    $existingNsg = Get-AzNetworkSecurityGroup -Name $nsgName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue
    if ($existingNsg) {
        Write-Log "NSG '$nsgName' already exists. Removing..."
        Remove-AzNetworkSecurityGroup -Name $nsgName -ResourceGroupName $resourceGroupName -Force -Confirm:$false
        Write-Log "NSG '$nsgName' removed."
    }
    Write-Log "Creating NSG '$nsgName' with rules:"
    Write-Log " - Denying inbound RDP on port 3389"
    Write-Log " - Allowing inbound RDP on port 10443"
    # Deny rule for TCP port 3389.
    $denyRule = New-AzNetworkSecurityRuleConfig -Name "Deny-RDP-3389" -Protocol Tcp -Direction Inbound -Priority 900 `
                 -SourceAddressPrefix * -SourcePortRange * -DestinationAddressPrefix * -DestinationPortRange 3389 `
                 -Access Deny -Description "Explicitly deny RDP traffic on port 3389"
    # Allow rule for TCP port 10443.
    $allowRule = New-AzNetworkSecurityRuleConfig -Name "Allow-RDP-10443" -Protocol Tcp -Direction Inbound -Priority 1000 `
                 -SourceAddressPrefix * -SourcePortRange * -DestinationAddressPrefix * -DestinationPortRange 10443 `
                 -Access Allow -Description "Allow RDP traffic on port 10443"
    $nsg = New-AzNetworkSecurityGroup -ResourceGroupName $resourceGroupName -Location $location -Name $nsgName `
           -SecurityRules @($denyRule, $allowRule) -ErrorAction Stop
    Write-Log "NSG '$nsgName' created."
} catch {
    Write-Log "ERROR: Failed to verify or create NSG. $_"
    Stop-Transcript
    exit 1
}

# 2. Check and remove existing Public IP then create a new one.
try {
    Write-Log "Checking for Public IP '$publicIpName'..."
    $existingPublicIp = Get-AzPublicIpAddress -Name $publicIpName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue
    if ($existingPublicIp) {
        Write-Log "Public IP '$publicIpName' already exists. Removing..."
        Remove-AzPublicIpAddress -Name $publicIpName -ResourceGroupName $resourceGroupName -Force -Confirm:$false
        Write-Log "Public IP '$publicIpName' removed."
    }
    Write-Log "Creating Public IP '$publicIpName'..."
    $publicIp = New-AzPublicIpAddress -Name $publicIpName -ResourceGroupName $resourceGroupName -Location $location -AllocationMethod Static -ErrorAction Stop
    Write-Log "Public IP '$publicIpName' created."
} catch {
    Write-Log "ERROR: Failed to create Public IP. $_"
    Stop-Transcript
    exit 1
}

# 3. Check and remove existing Virtual Network, then create a new one with subnet.
try {
    Write-Log "Checking for Virtual Network '$vnetName'..."
    $existingVnet = Get-AzVirtualNetwork -Name $vnetName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue
    if ($existingVnet) {
        Write-Log "Virtual Network '$vnetName' already exists. Removing..."
        Remove-AzVirtualNetwork -Name $vnetName -ResourceGroupName $resourceGroupName -Force -Confirm:$false
        Write-Log "Virtual Network '$vnetName' removed."
    }
    Write-Log "Creating Virtual Network '$vnetName' with subnet '$subnetName'..."
    $subnetConfig = New-AzVirtualNetworkSubnetConfig -Name $subnetName -AddressPrefix "10.0.0.0/24"
    $vnet = New-AzVirtualNetwork -Name $vnetName -ResourceGroupName $resourceGroupName -Location $location -AddressPrefix "10.0.0.0/16" -Subnet $subnetConfig -ErrorAction Stop
    Write-Log "Virtual Network '$vnetName' created."
} catch {
    Write-Log "ERROR: Failed to verify or create Virtual Network. $_"
    Stop-Transcript
    exit 1
}

# 4. Check and remove existing Network Interface, then create a new one; associate it with the NSG.
try {
    $nicName = "$vmName-NIC"
    Write-Log "Checking for Network Interface '$nicName'..."
    $existingNic = Get-AzNetworkInterface -Name $nicName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue
    if ($existingNic) {
        Write-Log "Network Interface '$nicName' already exists. Removing..."
        Remove-AzNetworkInterface -Name $nicName -ResourceGroupName $resourceGroupName -Force -Confirm:$false
        Write-Log "Network Interface '$nicName' removed."
    }
    Write-Log "Creating Network Interface for VM '$vmName' and associating NSG '$nsgName'..."
    $subnet = $vnet.Subnets | Where-Object { $_.Name -eq $subnetName }
    if (-not $subnet) {
        throw "Subnet '$subnetName' could not be found in Virtual Network '$vnetName'."
    }
    $nic = New-AzNetworkInterface -Name $nicName -ResourceGroupName $resourceGroupName -Location $location `
           -SubnetId $subnet.Id -PublicIpAddressId $publicIp.Id -NetworkSecurityGroupId $nsg.Id -ErrorAction Stop
    Write-Log "Network Interface for VM '$vmName' created."
} catch {
    Write-Log "ERROR: Failed to create Network Interface. $_"
    Stop-Transcript
    exit 1
}

# 5. Check and remove existing Virtual Machine, then create a new one.
try {
    Write-Log "Checking for Virtual Machine '$vmName'..."
    $existingVm = Get-AzVM -Name $vmName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue
    if ($existingVm) {
        Write-Log "Virtual Machine '$vmName' already exists. Removing..."
        Remove-AzVM -Name $vmName -ResourceGroupName $resourceGroupName -Force -Confirm:$false
        Write-Log "Virtual Machine '$vmName' removed."
    }
    Write-Log "Creating Virtual Machine '$vmName'..."
    $securePassword = ConvertTo-SecureString $adminPassword -AsPlainText -Force
    $cred = New-Object System.Management.Automation.PSCredential ($adminUsername, $securePassword)
    $vmConfig = New-AzVMConfig -VMName $vmName -VMSize "Standard_DS1_v2" -ErrorAction Stop |
      Set-AzVMOperatingSystem -Windows -ComputerName $vmName -Credential $cred -ProvisionVMAgent -EnableAutoUpdate -ErrorAction Stop |
      Set-AzVMSourceImage -PublisherName "MicrosoftWindowsServer" -Offer "WindowsServer" -Skus "2019-Datacenter" -Version "latest" -ErrorAction Stop |
      Add-AzVMNetworkInterface -Id $nic.Id -Primary -ErrorAction Stop

    # Configure boot diagnostics using the specified storage account.
    $vmConfig = Enable-AzVMBootDiagnostics -VM $vmConfig -ResourceGroupName $resourceGroupName -StorageAccountName $storageAccountName -ErrorAction Stop

    New-AzVM -ResourceGroupName $resourceGroupName -Location $location -VM $vmConfig -ErrorAction Stop | Out-Null
    Write-Log "Virtual Machine '$vmName' created successfully."
} catch {
    Write-Log "ERROR: Failed to create Virtual Machine. $_"
    Stop-Transcript
    exit 1
}

Write-Log "Script execution completed successfully."

# Stop transcript logging.
Stop-Transcript
Stop