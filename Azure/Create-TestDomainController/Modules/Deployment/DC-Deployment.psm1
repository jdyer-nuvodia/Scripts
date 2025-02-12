# =============================================================================
# Script: DC-Deployment.psm1
# Created: 2025-02-12 00:25:18 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-12 20:38:20 UTC
# Updated By: jdyer-nuvodia
# Version: 2.1
# Additional Info: Updated Public IP configuration for Standard SKU compliance
# =============================================================================

# Script-scoped variables
$Script:LogFile = $null

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter()]
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'DEPLOYMENT')]
        [string]$Level = 'INFO',
        [Parameter()]
        [string]$LogFile = $Script:LogFile
    )
    if ([string]::IsNullOrEmpty($LogFile)) {
        $LogFile = Join-Path -Path $PSScriptRoot -ChildPath "DC-Deployment.log"
    }
    $LogMessage = "[$([DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))] [$Level] $Message"
    Add-Content -Path $LogFile -Value $LogMessage
    Write-Host $LogMessage
    if ($Level -eq 'ERROR') {
        Write-Error $Message
    } elseif ($VerbosePreference -eq 'Continue') {
        Write-Verbose $Message
    }
}

function Set-DCLogFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    $Script:LogFile = $Path
    Write-Log "Log file path set to: $Path" -Level INFO
}

function Initialize-DCStorage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Config
    )
    try {
        Write-Log "Initializing storage account: $($Config.StorageAccountName)" -Level DEPLOYMENT
        $storageAccount = Get-AzStorageAccount -ResourceGroupName $Config.ResourceGroupName `
            -Name $Config.StorageAccountName -ErrorAction SilentlyContinue
        if (-not $storageAccount) {
            Write-Log "Storage account does not exist. Creating new storage account." -Level DEPLOYMENT
            $storageAccount = New-AzStorageAccount -ResourceGroupName $Config.ResourceGroupName `
                -Name $Config.StorageAccountName `
                -Location $Config.Location `
                -SkuName Standard_LRS `
                -Kind StorageV2
            Write-Log "Storage account created successfully" -Level DEPLOYMENT
        } else {
            Write-Log "Using existing storage account" -Level DEPLOYMENT
        }
        return $storageAccount
    }
    catch {
        Write-Log "Failed to initialize storage account: $_" -Level ERROR
        throw
    }
}

function New-DCPublicIP {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Config
    )
    try {
        Write-Log "Creating Public IP with Standard SKU: $($Config.PublicIpName)" -Level DEPLOYMENT
        $publicIpParams = @{
            Name              = $Config.PublicIpName
            ResourceGroupName = $Config.ResourceGroupName
            Location         = $Config.Location
            Sku              = 'Standard'
            AllocationMethod = 'Static'
            IpAddressVersion = 'IPv4'
            Zone             = 1
        }
        if ($PSCmdlet.ShouldProcess("Public IP $($Config.PublicIpName)", "Create")) {
            $publicIp = New-AzPublicIpAddress @publicIpParams
            Write-Log "Successfully created Public IP: $($Config.PublicIpName)" -Level DEPLOYMENT
            return $publicIp
        }
    }
    catch {
        Write-Log "Failed to create Public IP: $_" -Level ERROR
        throw
    }
}

function New-DCEnvironment {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Config
    )
    try {
        Write-Log "Starting Domain Controller environment deployment" -Level DEPLOYMENT
        
        # Create Resource Group if it doesn't exist
        if (-not (Get-AzResourceGroup -Name $Config.ResourceGroupName -ErrorAction SilentlyContinue)) {
            if ($PSCmdlet.ShouldProcess("Resource Group $($Config.ResourceGroupName)", "Create")) {
                Write-Log "Creating Resource Group: $($Config.ResourceGroupName)" -Level DEPLOYMENT
                New-AzResourceGroup -Name $Config.ResourceGroupName -Location $Config.Location
            }
        }

        # Initialize Storage Account
        if ($PSCmdlet.ShouldProcess("Storage Account $($Config.StorageAccountName)", "Initialize")) {
            $storageAccount = Initialize-DCStorage -Config $Config
        }

        # Create Virtual Network
        if ($PSCmdlet.ShouldProcess("Virtual Network $($Config.VnetName)", "Create")) {
            Write-Log "Creating Virtual Network: $($Config.VnetName)" -Level DEPLOYMENT
            $subnetConfig = New-AzVirtualNetworkSubnetConfig -Name $Config.SubnetName `
                -AddressPrefix $Config.SubnetAddressSpace
            $vnet = New-AzVirtualNetwork -Name $Config.VnetName `
                -ResourceGroupName $Config.ResourceGroupName `
                -Location $Config.Location `
                -AddressPrefix $Config.VnetAddressSpace `
                -Subnet $subnetConfig
        }

        # Create Public IP
        if ($PSCmdlet.ShouldProcess("Public IP $($Config.PublicIpName)", "Create")) {
            $publicIp = New-DCPublicIP -Config $Config
        }

        # Create Network Security Group
        if ($PSCmdlet.ShouldProcess("NSG $($Config.NsgName)", "Create")) {
            Write-Log "Creating Network Security Group: $($Config.NsgName)" -Level DEPLOYMENT
            $nsgRuleRDP = New-AzNetworkSecurityRuleConfig -Name "Allow-RDP" `
                -Description "Allow RDP" `
                -Access Allow `
                -Protocol Tcp `
                -Direction Inbound `
                -Priority 100 `
                -SourceAddressPrefix Internet `
                -SourcePortRange * `
                -DestinationAddressPrefix * `
                -DestinationPortRange 3389
            
            $nsg = New-AzNetworkSecurityGroup -Name $Config.NsgName `
                -ResourceGroupName $Config.ResourceGroupName `
                -Location $Config.Location `
                -SecurityRules $nsgRuleRDP
        }

        # Create Network Interface
        if ($PSCmdlet.ShouldProcess("Network Interface for VM $($Config.VmName)", "Create")) {
            Write-Log "Creating Network Interface for VM: $($Config.VmName)" -Level DEPLOYMENT
            $subnet = $vnet.Subnets[0]
            $nicName = "$($Config.VmName)-NIC"
            $nic = New-AzNetworkInterface -Name $nicName `
                -ResourceGroupName $Config.ResourceGroupName `
                -Location $Config.Location `
                -SubnetId $subnet.Id `
                -PublicIpAddressId $publicIp.Id `
                -NetworkSecurityGroupId $nsg.Id
        }

        # Create VM Configuration
        Write-Log "Configuring VM settings" -Level DEPLOYMENT
        $vmConfig = New-AzVMConfig -VMName $Config.VmName -VMSize $Config.VMSize
        
        # Configure boot diagnostics with existing storage account
        $vmConfig = Set-AzVMBootDiagnostic -VM $vmConfig `
            -Enable `
            -ResourceGroupName $Config.ResourceGroupName `
            -StorageAccountName $Config.StorageAccountName
        
        $securePassword = ConvertTo-SecureString $Config.AdminPassword -AsPlainText -Force
        $cred = New-Object System.Management.Automation.PSCredential ($Config.AdminUsername, $securePassword)

        $vmConfig = Set-AzVMOperatingSystem -VM $vmConfig `
            -Windows `
            -ComputerName $Config.VmName `
            -Credential $cred `
            -ProvisionVMAgent `
            -EnableAutoUpdate

        $vmConfig = Set-AzVMSourceImage -VM $vmConfig `
            -PublisherName $Config.ImagePublisher `
            -Offer $Config.ImageOffer `
            -Skus $Config.ImageSku `
            -Version $Config.ImageVersion

        $vmConfig = Add-AzVMNetworkInterface -VM $vmConfig -Id $nic.Id

        # Create Virtual Machine
        if ($PSCmdlet.ShouldProcess("Virtual Machine $($Config.VmName)", "Create")) {
            Write-Log "Creating Virtual Machine: $($Config.VmName)" -Level DEPLOYMENT
            $vm = New-AzVM -ResourceGroupName $Config.ResourceGroupName `
                -Location $Config.Location `
                -VM $vmConfig
            
            Write-Log "Virtual Machine deployment completed successfully" -Level DEPLOYMENT
            
            # Configure Auto-Shutdown
            if ($Config.ShutdownTime) {
                Write-Log "Configuring auto-shutdown schedule" -Level DEPLOYMENT
                $scheduleConfig = @{
                    Location              = $Config.Location
                    Name                  = "$($Config.VmName)-Shutdown"
                    DailyRecurrence      = $Config.ShutdownTime
                    TargetResourceId     = $vm.Id
                    TimeZoneId           = $Config.TimeZone
                }
                Enable-AzVmAutoShutdown @scheduleConfig
            }
        }

        Write-Log "Domain Controller environment deployment completed successfully" -Level DEPLOYMENT
        return $true
    }
    catch {
        Write-Log "Deployment failed: $_" -Level ERROR
        throw
    }
}

# Export functions
Export-ModuleMember -Function Write-Log, Set-DCLogFile, New-DCEnvironment