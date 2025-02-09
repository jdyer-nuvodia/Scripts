# =============================================================================
# Script: Create-TestDomainController.ps1
# Created: 2025-02-07 21:21:53 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-09 17:53:03 UTC
# Updated By: jdyer-nuvodia
# Version: 1.5
# Additional Info: Enhanced storage account creation with retry logic and timeout
# =============================================================================
<#
.SYNOPSIS
    Creates a test domain controller in Azure with existence checks, logging, and backup capabilities.
.DESCRIPTION
    This script provisions a test domain controller in Azure and performs the following actions:
    - Checks and creates the required Azure Resource Group.
    - Removes any existing conflicting resources (Storage Account, NSG, Public IP, Virtual Network, NIC, Virtual Machine) before re-creation.
    - Configures boot diagnostics and a backup mechanism.
    - Logs execution details via PowerShell transcript.
.PARAMETER None
    This script does not require parameters by default; parameters can be added as needed.
.EXAMPLE
    .\Create-TestDomainController.ps1
    This command runs the script to provision a test domain controller in Azure.
#>
[CmdletBinding()] Param()
Set-StrictMode -Version Latest; $ErrorActionPreference = "Stop"
if ($PSScriptRoot) { $scriptFolder = $PSScriptRoot } else { $scriptFolder = Get-Location }
$logPattern = "Create-TestDomainController-*.log"
$existingLogs = @(Get-ChildItem -Path $scriptFolder -Filter $logPattern -ErrorAction SilentlyContinue)
if ($existingLogs) {
    $latestLog = $existingLogs | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($latestLog) {
        Write-Host "Deleting previous log file: $($latestLog.FullName)"
        Remove-Item -Path $latestLog.FullName -Force -ErrorAction SilentlyContinue
    }
}
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$logFile = Join-Path $scriptFolder "Create-TestDomainController-$timestamp.log"
Start-Transcript -Path $logFile
$backupPattern = "Create-TestDomainController_Backup-*.ps1"
$existingBackups = @(Get-ChildItem -Path $scriptFolder -Filter $backupPattern -ErrorAction SilentlyContinue)
if ($existingBackups) {
    $latestBackup = $existingBackups | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($latestBackup) {
        Write-Host "Deleting previous backup file: $($latestBackup.FullName)"
        Remove-Item -Path $latestBackup.FullName -Force -ErrorAction SilentlyContinue
    }
}
$backupTimestamp = Get-Date -Format "yyyyMMddHHmmss"
$backupFile = Join-Path $scriptFolder "Create-TestDomainController_Backup-$backupTimestamp.ps1"
$currentScriptPath = $MyInvocation.MyCommand.Path
if (-not $currentScriptPath) {
    $expectedScriptName = "Create-TestDomainController.ps1"
    $fallbackScriptPath = Join-Path $scriptFolder $expectedScriptName
    if (Test-Path $fallbackScriptPath) {
        $currentScriptPath = $fallbackScriptPath
    } else {
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
    Write-Host "ERROR: Failed to create script backup: $_"
    Stop-Transcript
    exit 1
}
function Write-Log { param($Message); $timeStamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss"); Write-Host "[$timeStamp UTC] $Message" }
Write-Log "Script execution started."
Write-Verbose "Verbose mode activated."
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
$resourceGroupName    = "JB-TEST-RG2"
$location            = "westus2"
$storageAccountName  = "jbteststorage0"
$vnetName            = "JB-TEST-VNET"
$subnetName          = "JB-TEST-SUBNET1"
$vmName              = "JB-TEST-DC01"
$adminUsername       = "jbadmin"
$adminPassword       = "TS=pGxB~8m^A~WH^[yB8"
$domainName          = "JB-TEST.local"
$publicIpName        = "$vmName-PUBIP"
$nsgName             = "JB-TEST-NSG"

function Test-ResourceGroupExists {
    param($ResourceGroupName)
    try {
        Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction Stop | Out-Null
        return $true
    } catch {
        return $false
    }
}

function Wait-ForStorageAccount {
    param(
        [string]$ResourceGroupName,
        [string]$StorageAccountName,
        [int]$TimeoutSeconds = 300,
        [int]$RetryIntervalSeconds = 10
    )
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $completed = $false
    Write-Log "Waiting for storage account '$StorageAccountName' to be fully provisioned..."
    while (-not $completed -and $timer.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        try {
            $storageAccount = Get-AzStorageAccount -ResourceGroupName $ResourceGroupName -Name $StorageAccountName -ErrorAction Stop
            if ($storageAccount.ProvisioningState -eq "Succeeded") {
                Write-Log "Storage account '$StorageAccountName' is fully provisioned."
                $completed = $true
                return $true
            }
            Write-Log "Current provisioning state: $($storageAccount.ProvisioningState)"
        } catch {
            Write-Log "Waiting for storage account to be accessible... ($($timer.Elapsed.TotalSeconds) seconds elapsed)"
        }
        Start-Sleep -Seconds $RetryIntervalSeconds
    }
    $timer.Stop()
    if (-not $completed) {
        Write-Log "ERROR: Timeout waiting for storage account creation after $TimeoutSeconds seconds."
        return $false
    }
}

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

try {
    Write-Log "Checking for storage account '$storageAccountName'..."
    $maxRetries = 3
    $retryCount = 0
    $storageAccount = $null
    
    while ($retryCount -lt $maxRetries) {
        try {
            $existingStorage = Get-AzStorageAccount -Name $storageAccountName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue
            if ($existingStorage) {
                Write-Log "Storage account '$storageAccountName' already exists. Removing..."
                Remove-AzStorageAccount -Name $storageAccountName -ResourceGroupName $resourceGroupName -Force -Confirm:$false
                Write-Log "Storage account removed."
                Start-Sleep -Seconds 30  # Wait for removal to complete
            }
            
            Write-Log "Creating storage account '$storageAccountName' (Attempt $($retryCount + 1) of $maxRetries)..."
            $storageAccount = New-AzStorageAccount `
                -Name $storageAccountName `
                -ResourceGroupName $resourceGroupName `
                -Location $location `
                -SkuName Standard_LRS `
                -Kind StorageV2 `
                -ErrorAction Stop
            
            if (Wait-ForStorageAccount -ResourceGroupName $resourceGroupName -StorageAccountName $storageAccountName) {
                Write-Log "Storage account '$storageAccountName' created and verified successfully."
                break
            } else {
                throw "Storage account creation verification timed out."
            }
        } catch {
            $retryCount++
            Write-Log "WARNING: Attempt $retryCount failed to create storage account. Error: $_"
            if ($retryCount -ge $maxRetries) {
                throw "Failed to create storage account after $maxRetries attempts."
            }
            Write-Log "Waiting 60 seconds before retry..."
            Start-Sleep -Seconds 60
        }
    }
    
    if (-not $storageAccount) {
        throw "Failed to create storage account after all attempts."
    }
} catch {
    Write-Log "ERROR: Failed to verify or create storage account. $_"
    Stop-Transcript
    exit 1
}

try {
    Write-Log "Checking for Network Security Group '$nsgName'..."
    $existingNsg = Get-AzNetworkSecurityGroup -Name $nsgName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue
    if ($existingNsg) {
        Write-Log "NSG '$nsgName' already exists. Removing..."
        Remove-AzNetworkSecurityGroup -Name $nsgName -ResourceGroupName $resourceGroupName -Force -Confirm:$false
        Write-Log "NSG removed."
    }
    Write-Log "Creating NSG '$nsgName' with rules..."
    $denyRule = New-AzNetworkSecurityRuleConfig -Name "Deny-RDP-3389" -Protocol Tcp -Direction Inbound -Priority 900 -SourceAddressPrefix * -SourcePortRange * -DestinationAddressPrefix * -DestinationPortRange 3389 -Access Deny
    $allowRule = New-AzNetworkSecurityRuleConfig -Name "Allow-RDP-10443" -Protocol Tcp -Direction Inbound -Priority 1000 -SourceAddressPrefix * -SourcePortRange * -DestinationAddressPrefix * -DestinationPortRange 10443 -Access Allow
    $nsg = New-AzNetworkSecurityGroup -ResourceGroupName $resourceGroupName -Location $location -Name $nsgName -SecurityRules @($denyRule, $allowRule) -ErrorAction Stop
    Write-Log "NSG '$nsgName' created."
} catch {
    Write-Log "ERROR: Failed to verify or create NSG. $_"
    Stop-Transcript
    exit 1
}

try {
    Write-Log "Checking for Public IP '$publicIpName'..."
    $existingPublicIp = Get-AzPublicIpAddress -Name $publicIpName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue
    if ($existingPublicIp) {
        Write-Log "Public IP '$publicIpName' already exists. Removing..."
        Remove-AzPublicIpAddress -Name $publicIpName -ResourceGroupName $resourceGroupName -Force -Confirm:$false
        Write-Log "Public IP removed."
    }
    Write-Log "Creating Public IP '$publicIpName'..."
    $publicIp = New-AzPublicIpAddress -Name $publicIpName -ResourceGroupName $resourceGroupName -Location $location -AllocationMethod Static -ErrorAction Stop
    Write-Log "Public IP '$publicIpName' created."
} catch {
    Write-Log "ERROR: Failed to create Public IP. $_"
    Stop-Transcript
    exit 1
}

try {
    Write-Log "Checking for Virtual Network '$vnetName'..."
    $existingVnet = Get-AzVirtualNetwork -Name $vnetName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue
    if ($existingVnet) {
        Write-Log "Virtual Network '$vnetName' already exists. Removing..."
        Remove-AzVirtualNetwork -Name $vnetName -ResourceGroupName $resourceGroupName -Force -Confirm:$false
        Write-Log "Virtual Network removed."
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

try {
    $nicName = "$vmName-NIC"
    Write-Log "Checking for Network Interface '$nicName'..."
    $existingNic = Get-AzNetworkInterface -Name $nicName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue
    if ($existingNic) {
        Write-Log "Network Interface '$nicName' already exists. Removing..."
        Remove-AzNetworkInterface -Name $nicName -ResourceGroupName $resourceGroupName -Force -Confirm:$false
        Write-Log "Network Interface removed."
    }
    Write-Log "Creating Network Interface for VM '$vmName' and associating NSG '$nsgName'..."
    $subnet = $vnet.Subnets | Where-Object { $_.Name -eq $subnetName }
    if (-not $subnet) {
        throw "Subnet '$subnetName' could not be found in Virtual Network '$vnetName'."
    }
    $nic = New-AzNetworkInterface -Name $nicName -ResourceGroupName $resourceGroupName -Location $location -SubnetId $subnet.Id -PublicIpAddressId $publicIp.Id -NetworkSecurityGroupId $nsg.Id -ErrorAction Stop
    Write-Log "Network Interface for VM '$vmName' created."
} catch {
    Write-Log "ERROR: Failed to create Network Interface. $_"
    Stop-Transcript
    exit 1
    $nic = New-AzNetworkInterface -Name $nicName -ResourceGroupName $resourceGroupName -Location $location -SubnetId $subnet.Id -PublicIpAddressId $publicIp.Id -NetworkSecurityGroupId $nsg.Id -ErrorAction Stop
    Write-Log "Network Interface for VM '$vmName' created."
} catch {
    Write-Log "ERROR: Failed to create Network Interface. $_"
    Stop-Transcript
    exit 1
}

try {
    Write-Log "Checking for Virtual Machine '$vmName'..."
    $existingVm = Get-AzVM -Name $vmName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue
    if ($existingVm) {
        Write-Log "Virtual Machine '$vmName' already exists. Removing..."
        Remove-AzVM -Name $vmName -ResourceGroupName $resourceGroupName -Force -Confirm:$false
        Write-Log "Virtual Machine removed."
    }
    Write-Log "Creating Virtual Machine '$vmName'..."
    $securePassword = ConvertTo-SecureString $adminPassword -AsPlainText -Force
    $cred = New-Object System.Management.Automation.PSCredential ($adminUsername, $securePassword)
    $vmConfig = New-AzVMConfig -VMName $vmName -VMSize "Standard_DS1_v2" -ErrorAction Stop |
        Set-AzVMOperatingSystem -Windows -ComputerName $vmName -Credential $cred -ProvisionVMAgent -EnableAutoUpdate -ErrorAction Stop |
        Set-AzVMSourceImage -PublisherName "MicrosoftWindowsServer" -Offer "WindowsServer" -Skus "2019-Datacenter" -Version "latest" -ErrorAction Stop |
        Add-AzVMNetworkInterface -Id $nic.Id -Primary -ErrorAction Stop |
        Set-AzVMBootDiagnostic -Enable -ResourceGroupName $resourceGroupName -StorageAccountName $storageAccountName
    New-AzVM -ResourceGroupName $resourceGroupName -Location $location -VM $vmConfig -ErrorAction Stop | Out-Null
    Write-Log "Virtual Machine '$vmName' created successfully."
} catch {
    Write-Log "ERROR: Failed to create Virtual Machine. $_"
    Stop-Transcript
    exit 1
}

Write-Log "Script execution completed successfully."
Stop-Transcript