# =============================================================================
# Script: create-testdomaincontroller.ps1
# Created: 2025-02-05 00:22:44 UTC
# Author: jdyer-nuvodia
# Purpose: Creates a test domain controller in Azure with automated shutdown
# Version: 1.1
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
$tempRunbookFilePath = "C:\Temp\$([System.Guid]::NewGuid().ToString()).ps1"
$nsgName            = "JB-TEST-NSG"
$automationAccountName = "JB-TEST-Automation"
$runbookName        = "AutoShutdownRunbook"

# Import required modules with error handling
try {
    Write-Log "Importing required Azure modules..."
    Import-Module Az.Automation -ErrorAction Stop
    Import-Module Az.Storage -ErrorAction Stop
} catch {
    Write-Log "ERROR: Failed to import required modules. Error: $_"
    Write-Log "Please ensure Az.Automation and Az.Storage modules are installed."
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

# Function to validate Azure Automation prerequisites
function Test-AutomationPrerequisites {
    param (
        [string]$ResourceGroupName,
        [string]$AutomationAccountName
    )
    
    try {
        Write-Log "Validating Automation Account prerequisites..."
        
        # Check if automation account exists
        $automationAccount = Get-AzAutomationAccount -ResourceGroupName $ResourceGroupName `
                                                   -Name $AutomationAccountName `
                                                   -ErrorAction SilentlyContinue
        
        if (-not $automationAccount) {
            Write-Log "Automation account does not exist. Will create new one."
            return $false
        }
        
        # Validate permissions
        $currentContext = Get-AzContext
        if (-not $currentContext) {
            throw "No Azure context found. Please run Connect-AzAccount first."
        }
        
        Write-Log "Automation Account prerequisites validated successfully"
        return $true
    } catch {
        Write-Log "ERROR: Failed to validate Automation prerequisites: $_"
        throw
    }
}

# Function to handle blob storage operations with retries
function Get-BlobContent {
    param (
        [string]$ContainerName,
        [string]$BlobName,
        [string]$DestinationPath,
        $StorageContext
    )
    
    $maxRetries = 3
    $retryCount = 0
    $success = $false
    
    while (-not $success -and $retryCount -lt $maxRetries) {
        try {
            Write-Log "Attempting to download blob content (Attempt $($retryCount + 1) of $maxRetries)..."
            Get-AzStorageBlobContent -Context $StorageContext `
                                   -Container $ContainerName `
                                   -Blob $BlobName `
                                   -Destination $DestinationPath `
                                   -Force | Out-Null
            $success = $true
            Write-Log "Successfully downloaded blob content"
        } catch {
            $retryCount++
            if ($retryCount -lt $maxRetries) {
                Write-Log "Failed to download blob content. Retrying in 5 seconds..."
                Start-Sleep -Seconds 5
            } else {
                Write-Log "ERROR: Failed to download blob content after $maxRetries attempts. Error: $_"
                throw
            }
        }
    }
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

# Storage Account Creation/Validation with improved error handling
try {
    Write-Log "Checking storage account..."
    $storageAccount = Get-AzStorageAccount -ResourceGroupName $resourceGroup -Name $storageAccountName -ErrorAction SilentlyContinue
    
    if (-not $storageAccount) {
        Write-Log "Creating storage account $storageAccountName"
        $storageAccount = New-AzStorageAccount -ResourceGroupName $resourceGroup `
                                             -Name $storageAccountName `
                                             -Location $location `
                                             -SkuName Standard_LRS `
                                             -ErrorAction Stop
    } else {
        Write-Log "Storage account $storageAccountName already exists"
    }

    # Create storage context with validation
    Write-Log "Creating storage context..."
    $storageKey = (Get-AzStorageAccountKey -ResourceGroupName $resourceGroup -Name $storageAccountName -ErrorAction Stop)[0].Value
    if (-not $storageKey) {
        throw "Failed to retrieve storage account key"
    }
    
    $storageAccountContext = New-AzStorageContext -StorageAccountName $storageAccountName `
                                                -StorageAccountKey $storageKey `
                                                -Protocol 'HTTPS' `
                                                -ErrorAction Stop

    # Container validation/creation
    Write-Log "Checking container..."
    $container = Get-AzStorageContainer -Name $containerName -Context $storageAccountContext -ErrorAction SilentlyContinue
    if (-not $container) {
        Write-Log "Creating container $containerName"
        $container = New-AzStorageContainer -Name $containerName -Context $storageAccountContext -ErrorAction Stop
    } else {
        Write-Log "Container $containerName already exists"
    }

} catch {
    Write-Log "ERROR: Storage account setup failed. Error: $_"
    exit 1
}

# Ensure temp directory exists
try {
    Write-Log "Checking temporary directory..."
    if (-not (Test-Path -Path "C:\Temp")) {
        New-Item -ItemType Directory -Path "C:\Temp" -ErrorAction Stop
        Write-Log "Created temporary directory"
    }
} catch {
    Write-Log "ERROR: Failed to create temporary directory. Error: $_"
    exit 1
}

# Create runbook content with improved error handling
try {
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

    Set-Content -Path $tempRunbookFilePath -Value $runbookContent -ErrorAction Stop
    Write-Log "Runbook content created successfully"

    # Upload runbook to blob storage with validation
    Write-Log "Uploading runbook to blob storage..."
    Set-AzStorageBlobContent -Context $storageAccountContext `
                            -Container $containerName `
                            -File $tempRunbookFilePath `
                            -Blob $blobName `
                            -Force `
                            -ErrorAction Stop | Out-Null
    Write-Log "Runbook uploaded to blob storage successfully"

    # Clean up temporary file
    Remove-Item -Path $tempRunbookFilePath -ErrorAction SilentlyContinue
    Write-Log "Temporary file cleaned up"

} catch {
    Write-Log "ERROR: Failed to create or upload runbook content. Error: $_"
    if (Test-Path -Path $tempRunbookFilePath) {
        Remove-Item -Path $tempRunbookFilePath -ErrorAction SilentlyContinue
    }
    exit 1
}

# Network Configuration with improved error handling
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
        Write-Log "Virtual network created successfully"
    } else {
        Write-Log "Virtual network $vnetName already exists"
    }

    # Subnet Configuration
    $existingSubnet = Get-AzVirtualNetworkSubnetConfig -VirtualNetwork $virtualNetwork -Name $subnetName -ErrorAction SilentlyContinue
    if (-not $existingSubnet) {
        Write-Log "Adding subnet $subnetName to virtual network $vnetName"
        try {
            Add-AzVirtualNetworkSubnetConfig -VirtualNetwork $virtualNetwork `
                                           -AddressPrefix "10.0.1.0/24" `
                                           -Name $subnetName | Out-Null
            $virtualNetwork | Set-AzVirtualNetwork -ErrorAction Stop | Out-Null
            Write-Log "Subnet added successfully"
        } catch {
            Write-Log "ERROR: Failed to add subnet. Error: $_"
            throw
        }
    } else {
        Write-Log "Subnet $subnetName already exists in virtual network $vnetName"
    }

    # Network Security Group
    try {
        $nsg = Get-AzNetworkSecurityGroup -ResourceGroupName $resourceGroup -Name $nsgName -ErrorAction Stop
        Write-Log "Network Security Group $nsgName already exists"
    } catch {
        Write-Log "Creating Network Security Group $nsgName"
        $nsg = New-AzNetworkSecurityGroup -ResourceGroupName $resourceGroup `
                                        -Location $location `
                                        -Name $nsgName `
                                        -ErrorAction Stop
        Write-Log "Network Security Group created successfully"
    }

    # Network Interface
    $nicName = "$($vmName)VMNic"
    try {
        $nic = Get-AzNetworkInterface -ResourceGroupName $resourceGroup -Name $nicName -ErrorAction Stop
        Write-Log "Network interface $nicName already exists"
    } catch {
        Write-Log "Creating network interface $nicName"
        $subnetId = (Get-AzVirtualNetwork -ResourceGroupName $resourceGroup -Name $vnetName).Subnets[0].Id
        $nic = New-AzNetworkInterface -Name $nicName `
                                    -ResourceGroupName $resourceGroup `
                                    -Location $location `
                                    -SubnetId $subnetId `
                                    -ErrorAction Stop
        Write-Log "Network interface created successfully"
    }

    # Public IP
    try {
        $publicIp = Get-AzPublicIpAddress -ResourceGroupName $resourceGroup -Name $publicIpName -ErrorAction Stop
        Write-Log "Public IP address $publicIpName already exists"
    } catch {
        Write-Log "Creating public IP address $publicIpName"
        $publicIp = New-AzPublicIpAddress -Name $publicIpName `
                                        -ResourceGroupName $resourceGroup `
                                        -Location $location `
                                        -AllocationMethod Static `
                                        -Sku Standard `
                                        -ErrorAction Stop
        Write-Log "Public IP address created successfully"
    }

    # Connect Public IP to Network Interface
    Write-Log "Connecting public IP to network interface..."
    $nic = Get-AzNetworkInterface -ResourceGroupName $resourceGroup -Name $nicName
    $nic.IpConfigurations[0].PublicIpAddress = $publicIp
    Set-AzNetworkInterface -NetworkInterface $nic -ErrorAction Stop | Out-Null
    Write-Log "Network interface updated successfully"

} catch {
    Write-Log "ERROR: Network configuration failed. Error: $_"
    exit 1
}

# VM Configuration and Creation with improved error handling
try {
    Write-Log "Checking VM configuration..."
    $vm = Get-AzVM -ResourceGroupName $resourceGroup -Name $vmName -ErrorAction SilentlyContinue
    
    if ($vm) {
        Write-Log "VM $vmName already exists"
    } else {
        Write-Log "Creating VM $vmName"
        try {
            # VM Configuration
            Write-Log "Configuring VM settings..."
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

            # Source Image Configuration
            Write-Log "Setting VM source image..."
            $vmConfig = Set-AzVMSourceImage -VM $vmConfig `
                                          -PublisherName "MicrosoftWindowsServer" `
                                          -Offer "WindowsServer" `
                                          -Skus "2025-datacenter-core-g2" `
                                          -Version "latest" `
                                          -ErrorAction Stop

            # Network Configuration
            Write-Log "Adding network interface to VM configuration..."
            $vmConfig = Add-AzVMNetworkInterface -VM $vmConfig -Id $nic.Id -ErrorAction Stop

            # OS Disk Configuration
            Write-Log "Configuring VM OS disk..."
            $vmConfig = Set-AzVMOSDisk -VM $vmConfig `
                                      -Windows `
                                      -Caching ReadWrite `
                                      -CreateOption FromImage `
                                      -DiskSizeInGB 128 `
                                      -Name "$($vmName)OSDisk" `
                                      -ErrorAction Stop

            # Boot Diagnostics Configuration
            Write-Log "Configuring boot diagnostics..."
            $vmConfig = Set-AzVMBootDiagnostic -VM $vmConfig `
                                              -Enable `
                                              -StorageAccountName $storageAccountName `
                                              -ResourceGroupName $resourceGroup `
                                              -ErrorAction Stop

            # Security Profile Configuration
            Write-Log "Configuring security profile..."
            $vmConfig.SecurityProfile = @{
                SecurityType = "TrustedLaunch"
                UefiSettings = @{
                    SecureBootEnabled = $true
                    VTpmEnabled = $true
                }
            }

            # Create the VM
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

            # Update NSG Rules
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
}
catch {
    Write-Log "ERROR: VM configuration failed. Error: $_"
    exit 1
}

# Azure Automation Account Configuration with improved error handling
try {
    Write-Log "Configuring Azure Automation account..."
    
    # Validate automation prerequisites
    if (-not (Test-AutomationPrerequisites -ResourceGroupName $resourceGroup -AutomationAccountName $automationAccountName)) {
        Write-Log "Creating new Azure Automation account..."
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

    # Handle existing runbook
    Write-Log "Checking for existing runbook..."
    try {
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
    } catch {
        Write-Log "WARNING: Error handling existing runbook: $_"
        # Continue execution as this is not critical
    }

    # Create new runbook
    Write-Log "Creating new runbook..."
    try {
        New-AzAutomationRunbook -AutomationAccountName $automationAccountName `
                               -Name $runbookName `
                               -ResourceGroupName $resourceGroup `
                               -Type PowerShellWorkflow `
                               -ErrorAction Stop | Out-Null
        Write-Log "Runbook created successfully"
    } catch {
        Write-Log "ERROR: Failed to create runbook. Error: $_"
        throw
    }

    # Download and import runbook content
    Write-Log "Importing runbook content..."
    $downloadPath = "C:\Temp\$([System.Guid]::NewGuid().ToString()).ps1"
    try {
        Get-BlobContent -ContainerName $containerName `
                       -BlobName $blobName `
                       -DestinationPath $downloadPath `
                       -StorageContext $storageAccountContext

        Import-AzAutomationRunbook -Path $downloadPath `
                                  -Name $runbookName `
                                  -Type PowerShellWorkflow `
                                  -ResourceGroupName $resourceGroup `
                                  -AutomationAccountName $automationAccountName `
                                  -ErrorAction Stop

        Write-Log "Publishing runbook..."
        Publish-AzAutomationRunbook -Name $runbookName `
                                   -ResourceGroupName $resourceGroup `
                                   -AutomationAccountName $automationAccountName `
                                   -ErrorAction Stop

        Remove-Item -Path $downloadPath -ErrorAction SilentlyContinue
        Write-Log "Runbook published successfully"
    } catch {
        Write-Log "ERROR: Failed to import or publish runbook. Error: $_"
        if (Test-Path -Path $downloadPath) {
            Remove-Item -Path $downloadPath -ErrorAction SilentlyContinue
        }
        throw
    }

    # Create and register schedule
    Write-Log "Creating automation schedule..."
    try {
        $scheduleName = "AutoShutdownSchedule"
        $startTime = (Get-Date).AddMinutes(5)
        
        Write-Log "Creating schedule $scheduleName"
        New-AzAutomationSchedule -AutomationAccountName $automationAccountName `
                                -Name $scheduleName `
                                -StartTime $startTime `
                                -OneTime `
                                -ResourceGroupName $resourceGroup `
                                -ErrorAction Stop | Out-Null

        Write-Log "Registering runbook with schedule"
        Register-AzAutomationScheduledRunbook -AutomationAccountName $automationAccountName `
                                             -Name $runbookName `
                                             -ScheduleName $scheduleName `
                                             -ResourceGroupName $resourceGroup `
                                             -ErrorAction Stop
        
        Write-Log "Auto-shutdown schedule created and registered successfully"
    } catch {
        Write-Log "ERROR: Failed to create or register schedule. Error: $_"
        throw
    }

} catch {
    Write-Log "ERROR: Automation account configuration failed. Error: $_"
    exit 1
}

# =============================================================================
# Script: create-testdomaincontroller.ps1
# Created: 2025-02-05 00:29:07 UTC
# Author: jdyer-nuvodia
# Purpose: Creates a test domain controller in Azure with automated shutdown
# Version: 1.1
# =============================================================================

# Domain Controller Configuration
try {
    Write-Log "Configuring Domain Controller..."
    
    # Prepare DC configuration script
    $dcConfigScript = @"
# Install AD DS Role
Write-Host "Installing AD DS Role..."
Install-WindowsFeature -Name AD-Domain-Services -IncludeManagementTools

# Configure Domain Controller
Write-Host "Configuring Domain Controller..."
`$securePassword = ConvertTo-SecureString "$adminPassword" -AsPlainText -Force

Import-Module ADDSDeployment
Install-ADDSForest ``
    -CreateDnsDelegation:`$false ``
    -DatabasePath "C:\Windows\NTDS" ``
    -DomainMode "WinThreshold" ``
    -DomainName "$domainName" ``
    -ForestMode "WinThreshold" ``
    -InstallDns:`$true ``
    -LogPath "C:\Windows\NTDS" ``
    -NoRebootOnCompletion:`$false ``
    -SysvolPath "C:\Windows\SYSVOL" ``
    -SafeModeAdministratorPassword `$securePassword ``
    -Force::`$true

# Set DNS Server to self
Set-DnsClientServerAddress -InterfaceIndex (Get-NetAdapter).InterfaceIndex -ServerAddresses ("127.0.0.1")
"@

    # Create temporary script file
    $dcConfigPath = "C:\Temp\ConfigureDC_$([System.Guid]::NewGuid().ToString()).ps1"
    try {
        Set-Content -Path $dcConfigPath -Value $dcConfigScript -ErrorAction Stop
        Write-Log "Domain Controller configuration script created successfully"

        # Execute DC configuration
        Write-Log "Executing Domain Controller configuration..."
        $dcConfig = Invoke-AzVMRunCommand -ResourceGroupName $resourceGroup `
                                        -VMName $vmName `
                                        -CommandId 'RunPowerShellScript' `
                                        -ScriptPath $dcConfigPath `
                                        -ErrorAction Stop

        # Clean up temporary script
        Remove-Item -Path $dcConfigPath -ErrorAction SilentlyContinue
        Write-Log "Domain Controller configuration initiated successfully"

    } catch {
        Write-Log "ERROR: Failed to configure Domain Controller. Error: $_"
        if (Test-Path -Path $dcConfigPath) {
            Remove-Item -Path $dcConfigPath -ErrorAction SilentlyContinue
        }
        throw
    }

    # Final Configuration and Summary
    Write-Log "Getting public IP address..."
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
    Write-Log "ERROR: Final configuration failed. Error: $_"
    exit 1
} finally {
    # Cleanup any remaining temporary files
    Get-ChildItem -Path "C:\Temp" -Filter "ConfigureDC_*.ps1" | Remove-Item -Force -ErrorAction SilentlyContinue
    Get-ChildItem -Path "C:\Temp" -Filter "*.ps1" | Where-Object { $_.Name -match "[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}" } | Remove-Item -Force -ErrorAction SilentlyContinue
}

Write-Log "Script execution completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"