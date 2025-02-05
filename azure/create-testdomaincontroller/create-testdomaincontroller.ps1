# =============================================================================
# Script: create-testdomaincontroller.ps1
# Created: 2025-02-05 00:57:20 UTC
# Author: jdyer-nuvodia
# Purpose: Creates a test domain controller in Azure with automated shutdown
# Version: 1.2
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
$resourceGroup       = "JB-TEST-RG"
$location           = "westus"
$automationLocation = "westus2"
$vnetName           = "JB-TEST-VNET"
$subnetName         = "JB-TEST-SUBNET1"
$vmName             = "JB-TEST-DC01"
$adminUsername      = "jbadmin"
$adminPassword      = "TS=pGxB~8m^A~WH^[yB8"
$domainName         = "JB-TEST.local"
$publicIpName       = "$vmName-PublicIP"
$storageAccountName = "jbteststorage0"
$containerName      = "runbooks"
$blobName           = "AutoShutdownRunbook.ps1"
$nsgName           = "JB-TEST-NSG"
$automationAccountName = "JB-TEST-Automation"
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

# Storage Account Creation/Validation
try {
    Write-Log "Checking storage account..."
    $storageAccount = Get-AzStorageAccount -ResourceGroupName $resourceGroup -Name $storageAccountName -ErrorAction SilentlyContinue
    
    if (-not $storageAccount) {
        Write-Log "Creating storage account $storageAccountName"
        $storageAccount = New-AzStorageAccount -ResourceGroupName $resourceGroup `
                                             -Name $storageAccountName `
                                             -Location $location `
                                             -SkuName Standard_LRS
    } else {
        Write-Log "Storage account $storageAccountName already exists"
    }

    Write-Log "Creating storage context..."
    $storageKey = (Get-AzStorageAccountKey -ResourceGroupName $resourceGroup -Name $storageAccountName)[0].Value
    $storageAccountContext = New-AzStorageContext -StorageAccountName $storageAccountName `
                                                -StorageAccountKey $storageKey `
                                                -Protocol 'HTTPS'

    Write-Log "Checking container..."
    if (-not (Get-AzStorageContainer -Name $containerName -Context $storageAccountContext -ErrorAction SilentlyContinue)) {
        Write-Log "Creating container $containerName"
        New-AzStorageContainer -Name $containerName -Context $storageAccountContext -ErrorAction Stop
    } else {
        Write-Log "Container $containerName already exists"
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

# Network Configuration
try {
    Write-Log "Configuring network components..."
    
    # Virtual Network
    $virtualNetwork = Get-AzVirtualNetwork -Name $vnetName -ResourceGroupName $resourceGroup -ErrorAction SilentlyContinue
    if (-not $virtualNetwork) {
        Write-Log "Creating virtual network $vnetName"
        $vnet = @{
            ResourceGroupName = $resourceGroup
            Location = $location
            Name = $vnetName
            AddressPrefix = "10.0.0.0/16"
        }
        $virtualNetwork = New-AzVirtualNetwork @vnet -ErrorAction Stop
    } else {
        Write-Log "Virtual network $vnetName already exists"
    }

    # Subnet Configuration
    $subnet = Get-AzVirtualNetworkSubnetConfig -VirtualNetwork $virtualNetwork -Name $subnetName -ErrorAction SilentlyContinue
    if (-not $subnet) {
        Write-Log "Creating subnet $subnetName in virtual network $vnetName"
        Add-AzVirtualNetworkSubnetConfig -VirtualNetwork $virtualNetwork `
                                        -AddressPrefix "10.0.1.0/24" `
                                        -Name $subnetName | Out-Null
        $virtualNetwork | Set-AzVirtualNetwork -ErrorAction Stop | Out-Null
    } else {
        Write-Log "Subnet $subnetName already exists in virtual network $vnetName"
    }

    # Network Security Group
    Write-Log "Creating Network Security Group $nsgName"
    $nsg = New-AzNetworkSecurityGroup -ResourceGroupName $resourceGroup `
                                     -Location $location `
                                     -Name $nsgName
    Write-Log "Network Security Group created successfully"

    # Network Interface
    Write-Log "Creating network interface $($vmName)VMNic"
    $nic = New-AzNetworkInterface -Name "$($vmName)VMNic" `
                                 -ResourceGroupName $resourceGroup `
                                 -Location $location `
                                 -SubnetId $virtualNetwork.Subnets[0].Id
    Write-Log "Network interface created successfully"

    # Public IP
    Write-Log "Creating public IP address $publicIpName"
    $publicIp = New-AzPublicIpAddress -Name $publicIpName `
                                     -ResourceGroupName $resourceGroup `
                                     -Location $location `
                                     -AllocationMethod Static `
                                     -Sku Standard
    Write-Log "Public IP address created successfully"

    Write-Log "Connecting public IP to network interface..."
    $nic = Get-AzNetworkInterface -ResourceGroupName $resourceGroup -Name "$($vmName)VMNic"
    $nic.IpConfigurations[0].PublicIpAddress = $publicIp
    Set-AzNetworkInterface -NetworkInterface $nic | Out-Null
    Write-Log "Network interface updated successfully"

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

        Write-Log "Updating NSG rules for custom RDP port..."
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
        Write-Log "NSG rules updated successfully"
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

# Azure Automation Account Configuration
try {
    Write-Log "Configuring Azure Automation account..."
    
    # Function to validate automation prerequisites
    function Test-AutomationPrerequisites {
        param(
            [string]$ResourceGroupName,
            [string]$AutomationAccountName
        )
        return Get-AzAutomationAccount -ResourceGroupName $ResourceGroupName `
                                     -Name $AutomationAccountName `
                                     -ErrorAction SilentlyContinue
    }

    Write-Log "Validating Automation Account prerequisites..."
    if (-not (Test-AutomationPrerequisites -ResourceGroupName $resourceGroup -AutomationAccountName $automationAccountName)) {
        Write-Log "Automation account does not exist. Will create new one."
        try {
            $automationAccount = New-AzAutomationAccount -ResourceGroupName $resourceGroup `
                                                       -Name $automationAccountName `
                                                       -Location $automationLocation `
                                                       -ErrorAction Stop
            Write-Log "Automation account created successfully"
        } catch {
            Write-Log "ERROR: Failed to create Automation account. Error: $_"
            throw
        }
    }

    # Import runbook from blob storage
    Write-Log "Checking for existing runbook..."
    $existingRunbook = Get-AzAutomationRunbook -AutomationAccountName $automationAccountName `
                                              -Name $runbookName `
                                              -ResourceGroupName $resourceGroup `
                                              -ErrorAction SilentlyContinue

    if ($existingRunbook) {
        Write-Log "Removing existing runbook $runbookName"
        Remove-AzAutomationRunbook -AutomationAccountName $automationAccountName `
                                  -Name $runbookName `
                                  -ResourceGroupName $resourceGroup `
                                  -Force `
                                  -ErrorAction Stop
    }

    Write-Log "Creating new runbook..."
    New-AzAutomationRunbook -AutomationAccountName $automationAccountName `
                           -Name $runbookName `
                           -ResourceGroupName $resourceGroup `
                           -Type PowerShellWorkflow `
                           -ErrorAction Stop | Out-Null

    Write-Log "Importing runbook content..."
    $maxAttempts = 3
    $attempt = 1
    $success = $false

    while (-not $success -and $attempt -le $maxAttempts) {
        Write-Log "Attempting to download blob content (Attempt $attempt of $maxAttempts)..."
        try {
            $downloadPath = "C:\Temp\$runbookName.ps1"
            Get-AzStorageBlobContent -Container $containerName `
                                   -Blob $blobName `
                                   -Destination $downloadPath `
                                   -Context $storageAccountContext `
                                   -ErrorAction Stop

            Write-Log "Successfully downloaded blob content"
            
            Import-AzAutomationRunbook -Path $downloadPath `
                                      -Name $runbookName `
                                      -Type PowerShellWorkflow `
                                      -ResourceGroupName $resourceGroup `
                                      -AutomationAccountName $automationAccountName `
                                      -Force `
                                      -ErrorAction Stop

            Write-Log "Publishing runbook..."
            Publish-AzAutomationRunbook -Name $runbookName `
                                       -ResourceGroupName $resourceGroup `
                                       -AutomationAccountName $automationAccountName `
                                       -ErrorAction Stop

            $success = $true
            Remove-Item -Path $downloadPath -ErrorAction SilentlyContinue
        } catch {
            if ($attempt -eq $maxAttempts) {
                Write-Log "ERROR: Failed to import or publish runbook. Error: $_"
                throw
            }
            $attempt++
            Start-Sleep -Seconds 5
        }
    }

    # Create and register schedule
    Write-Log "Creating automation schedule..."
    $scheduleName = "AutoShutdownSchedule"
    $startTime = (Get-Date).AddMinutes(5)
    
    New-AzAutomationSchedule -AutomationAccountName $automationAccountName `
                            -Name $scheduleName `
                            -StartTime $startTime `
                            -OneTime `
                            -ResourceGroupName $resourceGroup `
                            -ErrorAction Stop | Out-Null

    Register-AzAutomationScheduledRunbook -AutomationAccountName $automationAccountName `
                                         -Name $runbookName `
                                         -ScheduleName $scheduleName `
                                         -ResourceGroupName $resourceGroup `
                                         -Parameters @{"resourceGroupName"=$resourceGroup; "vmName"=$vmName} `
                                         -ErrorAction Stop

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
    Write-Log "`nNOTE: The VM will automatically shut down in approximately 5 minutes."
    Write-Log "===================="

} catch {
    Write-Log "ERROR: Automation account configuration failed. Error: $_"
    exit 1
} finally {
    # Cleanup
    Get-ChildItem -Path "C:\Temp" -Filter "$runbookName*.ps1" | Remove-Item -Force -ErrorAction SilentlyContinue
}

Write-Log "Script execution completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"