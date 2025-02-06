# =============================================================================
# Script: create-testdomaincontroller.ps1
# Created: 2025-02-05 01:27:32 UTC
# Author: jdyer-nuvodia
# Purpose: Creates a test domain controller in Azure with automated shutdown
#
# Repository Information:
#   Repo: jdyer-nuvodia/Scripts
#   Repo ID: 924269019
#   Language Composition: PowerShell (100%)
#
# Version: 1.7
# =============================================================================

# Enable strict mode and stop on errors
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Initialize logging function
function Write-Log {
    param($Message)
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
}

Write-Log "Script execution started"
Write-Log "Initializing variables and checking prerequisites..."

# Variables
$resourceGroup       = "JB-TEST-RG2"
$location           = "westus2"
$vnetName           = "JB-TEST-VNET"
$subnetName         = "JB-TEST-SUBNET1"
$vmName             = "JB-TEST-DC01"
$adminUsername      = "jbadmin"
$adminPassword      = "TS=pGxB~8m^A~WH^[yB8"
$domainName         = "JB-TEST.local"
$publicIpName       = "$vmName-PUBIP"
$storageAccountName = "JB-TEST-STORAGE"
$containerName      = "RUNBOOKS"
$blobName           = "AutoShutdownRunbook.ps1"
$nsgName            = "JB-TEST-NSG"
$automationAccountName = "JB-TEST-AUTOMATION"
$runbookName        = "AutoShutdownRunbook"

# Import required modules with error handling
try {
    Write-Log "Importing required Azure modules..."
    Import-Module Az.Automation -ErrorAction Stop
    Import-Module Az.Storage -ErrorAction Stop
    Import-Module Az.Network -ErrorAction Stop
    Import-Module Az.Compute -ErrorAction Stop
} catch {
    Write-Log "ERROR: Failed to import required modules. Error: $_"
    Write-Log "Please ensure all required Az modules are installed."
    exit 1
}

# Function to wait for VM creation with improved error handling
function Wait-ForVM {
    param (
        [string]$ResourceGroupName,
        [string]$VMName,
        [int]$Timeout = 600
    )
    Write-Log "Waiting for VM $VMName to be ready..."
    $timer = 0
    while ($timer -lt $Timeout) {
        try {
            $vm = Get-AzVM -ResourceGroupName $ResourceGroupName -Name $VMName -ErrorAction Stop
            if ($vm) {
                Write-Log "VM $VMName is ready"
                return $true
            }
        } catch {
            Write-Log "Still waiting for VM $VMName... ($timer seconds elapsed)"
        }
        Start-Sleep -Seconds 10
        $timer += 10
    }
    Write-Log "ERROR: Timeout waiting for VM $VMName"
    return $false
}

# Resource Group Creation/Validation
try {
    Write-Log "Checking resource group..."
    if (-not (Get-AzResourceGroup -Name $resourceGroup -ErrorAction SilentlyContinue)) {
        Write-Log "Creating resource group $resourceGroup"
        New-AzResourceGroup -Name $resourceGroup -Location $location -ErrorAction Stop
    } else {
        Write-Log "Resource group $resourceGroup already exists"
    }
} catch {
    Write-Log "ERROR: Failed to create/validate resource group. Error: $_"
    exit 1
}

# Storage Account Creation/Validation with reuse
try {
    Write-Log "Checking storage account..."
    $storageAccount = Get-AzStorageAccount -ResourceGroupName $resourceGroup -Name $storageAccountName -ErrorAction SilentlyContinue
    
    if (-not $storageAccount) {
        Write-Log "Creating new storage account $storageAccountName"
        $storageAccount = New-AzStorageAccount -ResourceGroupName $resourceGroup `
                                           -Name $storageAccountName `
                                           -Location $location `
                                           -SkuName Standard_LRS
    } else {
        Write-Log "Reusing existing storage account $storageAccountName"
    }

    Write-Log "Creating storage context..."
    $storageKey = (Get-AzStorageAccountKey -ResourceGroupName $resourceGroup -Name $storageAccountName)[0].Value
    $storageAccountContext = New-AzStorageContext -StorageAccountName $storageAccountName `
                                               -StorageAccountKey $storageKey `
                                               -Protocol 'HTTPS'

    Write-Log "Checking container..."
    $container = Get-AzStorageContainer -Name $containerName -Context $storageAccountContext -ErrorAction SilentlyContinue
    if (-not $container) {
        Write-Log "Creating new container $containerName"
        New-AzStorageContainer -Name $containerName -Context $storageAccountContext -ErrorAction Stop
    } else {
        Write-Log "Reusing existing container $containerName"
    }
} catch {
    Write-Log "ERROR: Storage account setup failed. Error: $_"
    exit 1
}

# Create runbook content
try {
    Write-Log "Checking temporary directory..."
    if (-not (Test-Path -Path "C:\Temp")) {
        New-Item -ItemType Directory -Path "C:\Temp" -ErrorAction Stop
    }

    Write-Log "Creating runbook content..."
    $runbookContent = @"
workflow $runbookName {
    param (
        [string] `$resourceGroupName,
        [string] `$vmName
    )

    `$connection = Get-AutomationConnection -Name AzureRunAsConnection
    Add-AzAccount -ServicePrincipal -TenantId `$connection.TenantId -ApplicationId `$connection.ApplicationId -CertificateThumbprint `$connection.CertificateThumbprint

    Stop-AzVM -ResourceGroupName `$resourceGroupName -Name `$vmName -Force
}
"@

    $tempRunbookFilePath = "C:\Temp\$runbookName.ps1"
    Set-Content -Path $tempRunbookFilePath -Value $runbookContent -ErrorAction Stop
    Write-Log "Runbook content created successfully"

    Write-Log "Uploading runbook to blob storage..."
    Set-AzStorageBlobContent -Context $storageAccountContext `
                         -Container $containerName `
                         -File $tempRunbookFilePath `
                         -Blob $blobName `
                         -Force | Out-Null
    Write-Log "Runbook uploaded to blob storage successfully"
    
    Write-Log "Temporary file cleaned up"
    Remove-Item -Path $tempRunbookFilePath -ErrorAction SilentlyContinue

} catch {
    Write-Log "ERROR: Failed to create or upload runbook content. Error: $_"
    if (Test-Path -Path $tempRunbookFilePath) {
        Remove-Item -Path $tempRunbookFilePath -ErrorAction SilentlyContinue
    }
    exit 1
}

# Network Configuration with reuse
try {
    Write-Log "Configuring network components..."
    
    # Virtual Network
    $virtualNetwork = Get-AzVirtualNetwork -Name $vnetName -ResourceGroupName $resourceGroup -ErrorAction SilentlyContinue
    if (-not $virtualNetwork) {
        Write-Log "Creating new virtual network $vnetName"
        $vnet = @{
            ResourceGroupName = $resourceGroup
            Location = $location
            Name = $vnetName
            AddressPrefix = "10.0.0.0/16"
        }
        $virtualNetwork = New-AzVirtualNetwork @vnet -ErrorAction Stop
    } else {
        Write-Log "Reusing existing virtual network $vnetName"
    }

    # Subnet Configuration
    $subnet = Get-AzVirtualNetworkSubnetConfig -VirtualNetwork $virtualNetwork -Name $subnetName -ErrorAction SilentlyContinue
    if (-not $subnet) {
        Write-Log "Creating new subnet $subnetName"
        Add-AzVirtualNetworkSubnetConfig -VirtualNetwork $virtualNetwork `
                                     -AddressPrefix "10.0.1.0/24" `
                                     -Name $subnetName | Out-Null
        $virtualNetwork = $virtualNetwork | Set-AzVirtualNetwork -ErrorAction Stop
        $subnet = Get-AzVirtualNetworkSubnetConfig -VirtualNetwork $virtualNetwork -Name $subnetName -ErrorAction Stop
    } else {
        Write-Log "Reusing existing subnet $subnetName"
    }

    # Network Security Group
    $nsg = Get-AzNetworkSecurityGroup -ResourceGroupName $resourceGroup -Name $nsgName -ErrorAction SilentlyContinue
    if (-not $nsg) {
        Write-Log "Creating new Network Security Group $nsgName"
        $nsg = New-AzNetworkSecurityGroup -ResourceGroupName $resourceGroup `
                                       -Location $location `
                                       -Name $nsgName
    } else {
        Write-Log "Reusing existing Network Security Group $nsgName"
    }

    # Network Interface
    $nic = Get-AzNetworkInterface -ResourceGroupName $resourceGroup -Name "$($vmName)VMNic" -ErrorAction SilentlyContinue
    if (-not $nic) {
        Write-Log "Creating new network interface $($vmName)VMNic"
        $nic = New-AzNetworkInterface -Name "$($vmName)VMNic" `
                                   -ResourceGroupName $resourceGroup `
                                   -Location $location `
                                   -SubnetId $subnet.Id
    } else {
        Write-Log "Reusing existing network interface $($vmName)VMNic"
    }

    # Public IP
    $publicIp = Get-AzPublicIpAddress -ResourceGroupName $resourceGroup -Name $publicIpName -ErrorAction SilentlyContinue
    if (-not $publicIp) {
        Write-Log "Creating new public IP address $publicIpName"
        $publicIp = New-AzPublicIpAddress -Name $publicIpName `
                                       -ResourceGroupName $resourceGroup `
                                       -Location $location `
                                       -AllocationMethod Static `
                                       -Sku Standard
    } else {
        Write-Log "Reusing existing public IP address $publicIpName"
    }

    # Only update NIC if it's new or needs the public IP updated
    if (-not $nic.IpConfigurations[0].PublicIpAddress -or 
        $nic.IpConfigurations[0].PublicIpAddress.Id -ne $publicIp.Id) {
        Write-Log "Updating network interface with public IP..."
        $nic.IpConfigurations[0].PublicIpAddress = $publicIp
        Set-AzNetworkInterface -NetworkInterface $nic | Out-Null
    }

} catch {
    Write-Log "ERROR: Network configuration failed. Error: $_"
    exit 1
}

# VM Configuration and Creation
try {
    Write-Log "Checking VM configuration..."
    
    Write-Log "Creating VM $vmName"
    try {
        # VM Configuration
        $vmConfig = New-AzVMConfig -VMName $vmName -VMSize "Standard_B2s"
        
        # Operating System Configuration
        Write-Log "Configuring VM operating system..."
        $securePassword = ConvertTo-SecureString $adminPassword -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential($adminUsername, $securePassword)
        
        $vmConfig = Set-AzVMOperatingSystem -VM $vmConfig `
                                        -Windows `
                                        -ComputerName $vmName `
                                        -Credential $credential `
                                        -ErrorAction Stop

        Write-Log "Setting VM source image..."
        $vmConfig = Set-AzVMSourceImage -VM $vmConfig `
                                    -PublisherName "MicrosoftWindowsServer" `
                                    -Offer "WindowsServer" `
                                    -Skus "2025-datacenter-core-g2" `
                                    -Version "latest" `
                                    -ErrorAction Stop

        Write-Log "Adding network interface to VM configuration..."
        $vmConfig = Add-AzVMNetworkInterface -VM $vmConfig -Id $nic.Id -ErrorAction Stop

        Write-Log "Configuring VM OS disk..."
        $vmConfig = Set-AzVMOSDisk -VM $vmConfig `
                                -Windows `
                                -Caching ReadWrite `
                                -CreateOption FromImage `
                                -DiskSizeInGB 128 `
                                -Name "$($vmName)OSDisk" `
                                -ErrorAction Stop

        Write-Log "Configuring boot diagnostics..."
        $vmConfig = Set-AzVMBootDiagnostic -VM $vmConfig `
                                        -Enable `
                                        -StorageAccountName $storageAccountName `
                                        -ResourceGroupName $resourceGroup `
                                        -ErrorAction Stop

        Write-Log "Configuring security profile..."
        $vmConfig.SecurityProfile = @{
            SecurityType = "TrustedLaunch"
            UefiSettings = @{
                SecureBootEnabled = $true
                VTpmEnabled = $true
            }
        }

        Write-Log "Creating VM with Azure Hybrid Benefit..."
        $vmCreation = New-AzVM -ResourceGroupName $resourceGroup `
                            -Location $location `
                            -VM $vmConfig `
                            -LicenseType "Windows_Server" `
                            -ErrorAction Stop

        if (-not (Wait-ForVM -ResourceGroupName $resourceGroup -VMName $vmName)) {
            throw "Timeout waiting for VM creation"
        }
        Write-Log "VM created successfully"

        # Configure RDP Port
        Write-Log "Configuring custom RDP port..."
        $portvalue = 10443
        $scriptBlock = {
            Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -Name "PortNumber" -Value $using:portvalue
            New-NetFirewallRule -DisplayName 'RDPPORTLatest-TCP-In' -Profile 'Public' -Direction Inbound -Action Allow -Protocol TCP -LocalPort $using:portvalue
            New-NetFirewallRule -DisplayName 'RDPPORTLatest-UDP-In' -Profile 'Public' -Direction Inbound -Action Allow -Protocol UDP -LocalPort $using:portvalue
            Restart-Service -Name TermService -Force
        }

        Write-Log "Executing RDP port configuration..."
        $rdpConfig = Invoke-AzVMRunCommand -ResourceGroupName $resourceGroup `
                                       -VMName $vmName `
                                       -CommandId 'RunPowerShellScript' `
                                       -ScriptString $scriptBlock `
                                       -ErrorAction Stop

        Write-Log "Checking NSG rules for custom RDP port..."
        $rdpRule = Get-AzNetworkSecurityRuleConfig -NetworkSecurityGroup $nsg -Name "Allow_RDP_10443" -ErrorAction SilentlyContinue
        if (-not $rdpRule) {
            Write-Log "Creating new NSG rule Allow_RDP_10443"
            Add-AzNetworkSecurityRuleConfig -NetworkSecurityGroup $nsg `
                                        -Name "Allow_RDP_10443" `
                                        -Description "Allow RDP" `
                                        -Access Allow `
                                        -Protocol Tcp `
                                        -Direction Inbound `
                                        -Priority 1001 `
                                        -SourceAddressPrefix * `
                                        -SourcePortRange * `
                                        -DestinationAddressPrefix * `
                                        -DestinationPortRange 10443 | Out-Null
            $nsg | Set-AzNetworkSecurityGroup -ErrorAction Stop | Out-Null
            Write-Log "NSG rule created successfully"
        } else {
            Write-Log "Reusing existing NSG rule Allow_RDP_10443"
        }
    }
    catch {
        Write-Log "ERROR: VM creation failed. Error: $_"
        throw
    }
}
catch {
    Write-Log "ERROR: VM configuration failed. Error: $_"
    exit 1
}

# Azure Automation Account Configuration with reuse
try {
    Write-Log "Configuring Azure Automation account..."
    
    # Check if automation account exists
    $automationAccount = Get-AzAutomationAccount -ResourceGroupName $resourceGroup `
                                             -Name $automationAccountName `
                                             -ErrorAction SilentlyContinue
    
    if (-not $automationAccount) {
        Write-Log "Creating new Automation Account $automationAccountName"
        $automationAccount = New-AzAutomationAccount -ResourceGroupName $resourceGroup `
                                                 -Name $automationAccountName `
                                                 -Location $location `
                                                 -ErrorAction Stop
    } else {
        Write-Log "Reusing existing Automation Account $automationAccountName"
    }

    # Check for existing runbook but only update if content is different
    $existingRunbook = Get-AzAutomationRunbook -AutomationAccountName $automationAccountName `
                                            -Name $runbookName `
                                            -ResourceGroupName $resourceGroup `
                                            -ErrorAction SilentlyContinue

    $shouldUpdateRunbook = $true
    if ($existingRunbook) {
        Write-Log "Existing runbook found. Checking if update is needed..."
        $exportPath = "C:\Temp\ExistingRunbook.ps1"
        Export-AzAutomationRunbook -ResourceGroupName $resourceGroup `
                               -AutomationAccountName $automationAccountName `
                               -Name $runbookName `
                               -OutputFolder (Split-Path $exportPath) `
                               -Slot "Published" `
                               -ErrorAction SilentlyContinue

        if (Test-Path $exportPath) {
            $existingContent = Get-Content $exportPath -Raw
            $newContent = Get-AzStorageBlobContent -Container $containerName `
                                               -Blob $blobName `
                                               -Context $storageAccountContext `
                                               -Force `
                                               -AsString
            if ($existingContent -eq $newContent) {
                Write-Log "Runbook content is unchanged. Skipping update."
                $shouldUpdateRunbook = $false
            }
            Remove-Item $exportPath -Force
        }
    }

    if ($shouldUpdateRunbook) {
        Write-Log "Updating runbook content..."
        if ($existingRunbook) {
            Remove-AzAutomationRunbook -AutomationAccountName $automationAccountName `
                                   -Name $runbookName `
                                   -ResourceGroupName $resourceGroup `
                                   -Force `
                                   -ErrorAction Stop
        }

        # Create and import new runbook
        New-AzAutomationRunbook -AutomationAccountName $automationAccountName `
                             -Name $runbookName `
                             -ResourceGroupName $resourceGroup `
                             -Type PowerShellWorkflow `
                             -ErrorAction Stop | Out-Null

        # Import updated content
        $downloadPath = "C:\Temp\$runbookName.ps1"
        Get-AzStorageBlobContent -Container $containerName `
                             -Blob $blobName `
                             -Destination $downloadPath `
                             -Context $storageAccountContext `
                             -ErrorAction Stop

        Import-AzAutomationRunbook -Path $downloadPath `
                                -Name $runbookName `
                                -Type PowerShellWorkflow `
                                -ResourceGroupName $resourceGroup `
                                -AutomationAccountName $automationAccountName `
                                -Force `
                                -ErrorAction Stop

        Write-Log "Publishing updated runbook..."
        Publish-AzAutomationRunbook -Name $runbookName `
                                 -ResourceGroupName $resourceGroup `
                                 -AutomationAccountName $automationAccountName `
                                 -ErrorAction Stop

        Remove-Item -Path $downloadPath -ErrorAction SilentlyContinue
    }

    # Daily Automation Schedule Creation for 9pm MST (Phoenix)
    # Phoenix is fixed as UTC-7 (no daylight savings)
    Write-Log "Creating daily automation schedule..."
    $nowUtc = (Get-Date).ToUniversalTime()
    $phxOffsetHours = -7
    $nowPhx = $nowUtc.AddHours($phxOffsetHours)

    # Get 9pm today in Phoenix time
    $today9pmPhx = Get-Date -Year $nowPhx.Year -Month $nowPhx.Month -Day $nowPhx.Day -Hour 21 -Minute 0 -Second 0

    if ($nowPhx -ge $today9pmPhx) {
        # If it's already past 9pm in Phoenix, schedule for tomorrow
        $next9pmPhx = $today9pmPhx.AddDays(1)
    } else {
        $next9pmPhx = $today9pmPhx
    }

    # Convert next 9pm Phoenix time (UTC-7) to UTC
    $startTimeUtc = $next9pmPhx.AddHours(-$phxOffsetHours)

    Write-Log "Calculated next daily shutdown time (UTC): $($startTimeUtc.ToString('yyyy-MM-dd HH:mm:ss'))"

    $scheduleName = "DailyAutoShutdownSchedule"
    try {
        $schedule = New-AzAutomationSchedule -AutomationAccountName $automationAccountName `
                                         -Name $scheduleName `
                                         -StartTime $startTimeUtc `
                                         -ExpiryTime ($startTimeUtc.AddYears(5)) `
                                         -Interval 1 `
                                         -Frequency Day `
                                         -ResourceGroupName $resourceGroup `
                                         -ErrorAction Stop
        Write-Log "Schedule '$scheduleName' created successfully for daily execution at 9pm MST (Phoenix)."
    
        Register-AzAutomationScheduledRunbook -AutomationAccountName $automationAccountName `
                                          -Name $runbookName `
                                          -ScheduleName $scheduleName `
                                          -ResourceGroupName $resourceGroup `
                                          -Parameters @{ "resourceGroupName" = $resourceGroup; "vmName" = $vmName } `
                                          -ErrorAction Stop
        Write-Log "Runbook '$runbookName' scheduled successfully to shut down the VM daily at 9pm MST (Phoenix)."
    }
    catch {
        Write-Log "ERROR: Failed to create or register the daily schedule. Error: $_"
        exit 1
    }
    
    # Output Configuration Summary
    $publicIpAddress = (Get-AzPublicIpAddress -ResourceGroupName $resourceGroup -Name $publicIpName).IpAddress
    
    Write-Log "`nConfiguration Complete!"
    Write-Log "===================="
    Write-Log "Domain Controller Configuration Summary:"
    Write-Log "VM Name: $vmName"
    Write-Log "Public IP: $publicIpAddress"
    Write-Log "RDP Port: 10443"
    Write-Log "Domain: $domainName"
    Write-Log "Admin Username: $adminUsername"
    Write-Log "Resource Group: $resourceGroup"
    Write-Log "Location: $location"
    Write-Log "`nNOTE: The VM will automatically shut down daily at 9pm MST (Phoenix)."
    Write-Log "===================="

} catch {
    Write-Log "ERROR: Automation account configuration failed. Error: $_"
    exit 1
} finally {
    # Cleanup
    Get-ChildItem -Path "C:\Temp" -Filter "$runbookName*.ps1" | Remove-Item -Force -ErrorAction SilentlyContinue
}

Write-Log "Script execution completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"