# =============================================================================
# Script: DC-Deployment.ps1
# Created: 2025-02-12 00:07:30 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-12 00:07:30 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1
# Additional Info: Fixed syntax error and updated VM size for Trusted Launch
# =============================================================================

function New-DCEnvironment {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Config
    )
    
    try {
        # Create Resource Group if it doesn't exist
        $rg = Get-AzResourceGroup -Name $Config.ResourceGroupName -ErrorAction SilentlyContinue
        if (-not $rg) {
            Write-Log "Creating resource group '$($Config.ResourceGroupName)'..." -Level INFO
            New-AzResourceGroup -Name $Config.ResourceGroupName -Location $Config.Location -ErrorAction Stop
        }

        # Create Storage Account if it doesn't exist
        $storageAccount = Get-AzStorageAccount -ResourceGroupName $Config.ResourceGroupName `
            -Name $Config.StorageAccountName -ErrorAction SilentlyContinue
        if (-not $storageAccount) {
            Write-Log "Creating Storage Account '$($Config.StorageAccountName)'..." -Level INFO
            $storageAccountParams = @{
                ResourceGroupName = $Config.ResourceGroupName
                Name = $Config.StorageAccountName
                Location = $Config.Location
                SkuName = 'Standard_LRS'
                Kind = 'StorageV2'
            }
            $storageAccount = New-AzStorageAccount @storageAccountParams
        }

        # Create Network Security Group
        $nsg = Get-AzNetworkSecurityGroup -Name $Config.NsgName -ResourceGroupName $Config.ResourceGroupName -ErrorAction SilentlyContinue
        if (-not $nsg) {
            Write-Log "Creating Network Security Group '$($Config.NsgName)'..." -Level INFO
            $nsgRules = @{
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
            $nsg = New-AzNetworkSecurityGroup -ResourceGroupName $Config.ResourceGroupName -Location $Config.Location -Name $Config.NsgName -ErrorAction Stop
            Add-AzNetworkSecurityRuleConfig @nsgRules -NetworkSecurityGroup $nsg
            $nsg | Set-AzNetworkSecurityGroup
        }

        # Deploy VM and Configure
        Deploy-DomainController -Config $Config -Nsg $nsg
        
        # Configure auto-shutdown
        Set-VMAutoShutdown -Config $Config
        
    } catch {
        Write-Log "Deployment error: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

function Deploy-DomainController {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Config,
        [Parameter(Mandatory = $true)]
        [Microsoft.Azure.Commands.Network.Models.PSNetworkSecurityGroup]$Nsg
    )
    
    # Create or get Virtual Network and Subnet
    $vnet = Get-AzVirtualNetwork -Name $Config.VnetName -ResourceGroupName $Config.ResourceGroupName -ErrorAction SilentlyContinue
    if (-not $vnet) {
        $subnetConfig = New-AzVirtualNetworkSubnetConfig -Name $Config.SubnetName -AddressPrefix $Config.SubnetAddressSpace -NetworkSecurityGroup $nsg
        $vnet = New-AzVirtualNetwork -ResourceGroupName $Config.ResourceGroupName -Location $Config.Location -Name $Config.VnetName `
            -AddressPrefix $Config.VnetAddressSpace -Subnet $subnetConfig -ErrorAction Stop
    }

    # Create Public IP
    $publicIp = New-AzPublicIpAddress -ResourceGroupName $Config.ResourceGroupName -Location $Config.Location `
        -Name $Config.PublicIpName -Sku Standard -AllocationMethod Static -ErrorAction Stop

    # Create NIC
    $nicName = "$($Config.VmName)-NIC"
    $subnet = Get-AzVirtualNetworkSubnetConfig -VirtualNetwork $vnet -Name $Config.SubnetName
    $nic = New-AzNetworkInterface -Name $nicName -ResourceGroupName $Config.ResourceGroupName `
        -Location $Config.Location -SubnetId $subnet.Id -PublicIpAddressId $publicIp.Id -ErrorAction Stop

    # Create VM Configuration
    Write-Log "Creating VM configuration..." -Level INFO
    try {
        $vmConfig = New-AzVMConfig -VMName $Config.VmName -VMSize $Config.VMSize -SecurityType "TrustedLaunch"
        
        # Configure OS
        $credential = New-Object System.Management.Automation.PSCredential ($Config.AdminUsername, (ConvertTo-SecureString $Config.AdminPassword -AsPlainText -Force))
        $vmConfig = Set-AzVMOperatingSystem -VM $vmConfig -Windows -ComputerName $Config.VmName `
            -Credential $credential -ProvisionVMAgent -EnableAutoUpdate

        # Add network interface
        $vmConfig = Add-AzVMNetworkInterface -VM $vmConfig -Id $nic.Id

        # Set source image
        $vmConfig = Set-AzVMSourceImage -VM $vmConfig `
            -PublisherName $Config.ImagePublisher `
            -Offer $Config.ImageOffer `
            -Skus $Config.ImageSku `
            -Version $Config.ImageVersion

        # Configure boot diagnostics
        $vmConfig = Set-AzVMBootDiagnostic -VM $vmConfig -Enable `
            -ResourceGroupName $Config.ResourceGroupName `
            -StorageAccountName $Config.StorageAccountName

        # Configure Trusted Launch security settings
        $securityProfile = @{
            SecurityType = "TrustedLaunch"
            UefiSettings = @{
                SecureBootEnabled = $true
                VTpmEnabled = $true
            }
        }
        $vmConfig.SecurityProfile = $securityProfile

        # Create the VM
        Write-Log "Creating VM '$($Config.VmName)'..." -Level INFO
        $newVM = New-AzVM -ResourceGroupName $Config.ResourceGroupName -Location $Config.Location -VM $vmConfig
        if ($newVM) {
            Write-Log "VM created successfully" -Level INFO
        } else {
            throw "VM creation failed without specific error"
        }
    } catch {
        Write-Log "Error during VM configuration or creation: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

function Set-VMAutoShutdown {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Config
    )
    
    Write-Log "Configuring auto-shutdown schedule..." -Level INFO
    $scheduledShutdownResourceId = "/subscriptions/{0}/resourceGroups/{1}/providers/microsoft.devtestlab/schedules/shutdown-computevm-{2}" -f `
        (Get-AzContext).Subscription.Id, $Config.ResourceGroupName, $Config.VmName
    
    $properties = @{
        status = "Enabled"
        taskType = "ComputeVmShutdownTask"
        dailyRecurrence = @{ time = $Config.ShutdownTime }
        timeZoneId = $Config.TimeZone
        targetResourceId = (Get-AzVM -ResourceGroupName $Config.ResourceGroupName -Name $Config.VmName).Id
        notificationSettings = @{ status = "Disabled" }
        location = $Config.Location
    }
    
    New-AzResource -ResourceId $scheduledShutdownResourceId -Properties $properties -Force -Location $Config.Location
}