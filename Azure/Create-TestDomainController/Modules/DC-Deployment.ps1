# =============================================================================
# Script: DC-Deployment.ps1
# Created: 2025-02-11 23:45:10 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-11 23:45:10 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0
# Additional Info: Deployment module for Domain Controller creation
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
    $subnet