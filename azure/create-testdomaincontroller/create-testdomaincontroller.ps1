# =============================================================================
# Script: create-testdomaincontroller.ps1
# Created: 2025-02-05 01:27:32 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-07 21:50:05 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2
# Purpose: Creates a test domain controller in Azure with automated shutdown
# =============================================================================

# Script Configuration and Error Handling
Set-StrictMode -Version Latest
$ErrorActionPreference="Stop"

# Logging Function
function Write-Log{param($m)Write-Host "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) UTC] $m"}

# Core Variables
$rg="JB-TEST-RG2";$loc="westus2";$vm="JB-TEST-DC01"
$vnet="JB-TEST-VNET";$subnet="JB-TEST-SUBNET1";$nsg="JB-TEST-NSG"
$pip="$vm-PUBIP";$stor="jbteststorage0";$cont="runbooks"
$auto="JB-TEST-AUTOMATION";$rb="AutoShutdownRunbook"
$cred=New-Object PSCredential("jbadmin",(ConvertTo-SecureString "TS=pGxB~8m^A~WH^[yB8" -AsPlainText -Force))

# Module Import with Error Handling
Write-Log "Importing Azure modules..."
try{
    @('Az.Automation','Az.Storage','Az.Network','Az.Compute')|%{Import-Module $_ -EA Stop}
}catch{Write-Log "ERROR: Module import failed: $_";exit 1}

# Resource Group Setup
Write-Log "Setting up Resource Group..."
try{
    if(!(Get-AzResourceGroup -Name $rg -EA SilentlyContinue)){
        New-AzResourceGroup -Name $rg -Location $loc -EA Stop
    }
}catch{Write-Log "ERROR: RG creation failed: $_";exit 1}

# Storage Account and Container Setup
Write-Log "Configuring Storage..."
try{
    $sa=Get-AzStorageAccount -ResourceGroupName $rg -Name $stor -EA SilentlyContinue
    if(!$sa){
        $sa=New-AzStorageAccount -ResourceGroupName $rg -Name $stor -Location $loc -SkuName Standard_LRS
    }
    $ctx=New-AzStorageContext -StorageAccountName $stor -StorageAccountKey (Get-AzStorageAccountKey -ResourceGroupName $rg -Name $stor)[0].Value
    if(!(Get-AzStorageContainer -Name $cont -Context $ctx -EA SilentlyContinue)){
        New-AzStorageContainer -Name $cont -Context $ctx
    }
}catch{Write-Log "ERROR: Storage setup failed: $_";exit 1}

# Runbook Creation and Upload
Write-Log "Creating Runbook..."
try{
    $rbContent=@"
workflow $rb {
    param([string]`$rg,[string]`$vm)
    `$conn=Get-AutomationConnection -Name AzureRunAsConnection
    Add-AzAccount -ServicePrincipal -TenantId `$conn.TenantId -ApplicationId `$conn.ApplicationId -CertificateThumbprint `$conn.CertificateThumbprint
    Stop-AzVM -ResourceGroupName `$rg -Name `$vm -Force
}
"@
    $rbPath="C:\Temp\$rb.ps1"
    if(!(Test-Path "C:\Temp")){mkdir "C:\Temp"}
    Set-Content $rbPath $rbContent
    Set-AzStorageBlobContent -Context $ctx -Container $cont -File $rbPath -Blob "$rb.ps1" -Force|Out-Null
    Remove-Item $rbPath -EA SilentlyContinue
}catch{Write-Log "ERROR: Runbook creation failed: $_";exit 1}

# Network Configuration
Write-Log "Setting up Network..."
try{
    # Virtual Network
    $vnetObj=Get-AzVirtualNetwork -Name $vnet -ResourceGroupName $rg -EA SilentlyContinue
    if(!$vnetObj){
        $vnetObj=New-AzVirtualNetwork -ResourceGroupName $rg -Location $loc -Name $vnet -AddressPrefix "10.0.0.0/16"
    }

    # Subnet
    $subnetObj=$vnetObj|Get-AzVirtualNetworkSubnetConfig -Name $subnet -EA SilentlyContinue
    if(!$subnetObj){
        $vnetObj|Add-AzVirtualNetworkSubnetConfig -Name $subnet -AddressPrefix "10.0.1.0/24"|Set-AzVirtualNetwork|Out-Null
        $subnetObj=$vnetObj|Get-AzVirtualNetworkSubnetConfig -Name $subnet
    }

    # NSG with RDP Rules
    $nsgObj=Get-AzNetworkSecurityGroup -ResourceGroupName $rg -Name $nsg -EA SilentlyContinue
    if(!$nsgObj){$nsgObj=New-AzNetworkSecurityGroup -ResourceGroupName $rg -Location $loc -Name $nsg}

    # Configure RDP Rules
    $rdpRules=@(
        @{Name="Allow_RDP_3389";Protocol="Tcp";Access="Allow";Priority=100;SourceAddressPrefix="Internet";DestinationPortRange="3389"},
        @{Name="Allow_RDP_10443";Protocol="Tcp";Access="Allow";Priority=101;SourceAddressPrefix="Internet";DestinationPortRange="10443"}
    )
    $updated=$false
    foreach($rule in $rdpRules){
        if(!(Get-AzNetworkSecurityRuleConfig -NetworkSecurityGroup $nsgObj -Name $rule.Name -EA SilentlyContinue)){
            $nsgObj|Add-AzNetworkSecurityRuleConfig -Direction "Inbound" -SourcePortRange "*" `
                -DestinationAddressPrefix "VirtualNetwork" @rule|Out-Null
            $updated=$true
        }
    }
    if($updated){$nsgObj|Set-AzNetworkSecurityGroup|Out-Null}

    # NIC and Public IP
    $nic=Get-AzNetworkInterface -ResourceGroupName $rg -Name "$($vm)VMNic" -EA SilentlyContinue
    $pubip=Get-AzPublicIpAddress -ResourceGroupName $rg -Name $pip -EA SilentlyContinue
    if(!$pubip){
        $pubip=New-AzPublicIpAddress -Name $pip -ResourceGroupName $rg -Location $loc -AllocationMethod Static -Sku Standard
    }
    if(!$nic){
        $nic=New-AzNetworkInterface -Name "$($vm)VMNic" -ResourceGroupName $rg -Location $loc `
            -SubnetId $subnetObj.Id -NetworkSecurityGroupId $nsgObj.Id -PublicIpAddressId $pubip.Id
    }
}catch{Write-Log "ERROR: Network setup failed: $_";exit 1}

# VM Creation and Configuration
Write-Log "Creating VM..."
try{
    $vmConfig=New-AzVMConfig -VMName $vm -VMSize "Standard_B2s"|
        Set-AzVMOperatingSystem -Windows -ComputerName $vm -Credential $cred|
        Set-AzVMSourceImage -PublisherName "MicrosoftWindowsServer" -Offer "WindowsServer" `
            -Skus "2025-datacenter-core-g2" -Version "latest"|
        Add-AzVMNetworkInterface -Id $nic.Id|
        Set-AzVMOSDisk -Name "$($vm)_OSDisk" -CreateOption FromImage -Windows

    New-AzVM -ResourceGroupName $rg -Location $loc -VM $vmConfig
    Write-Log "VM creation completed successfully"
}catch{Write-Log "ERROR: VM creation failed: $_";exit 1}

Write-Log "Script completed successfully"