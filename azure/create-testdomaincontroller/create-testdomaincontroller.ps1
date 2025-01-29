# Variables
$resourceGroup = "JB-TEST-RG"
$location = "westus"
$vnetName = "JB-TEST-VNET"
$subnetName = "JB-TEST-SUBNET1"
$vmName = "JB-TEST-DC01"
$adminUsername = "jbadmin"
$adminPassword = "TS=pGxB~8m^A~WH^[yB8"
$domainName = "JB-TEST.local"
$publicIpName = "$vmName-PublicIP"
$storageAccountName = "jbteststorage0"

# Function to wait for VM creation
function Wait-ForVM {
    param (
        [string]$ResourceGroupName,
        [string]$VMName,
        [int]$Timeout = 600
    )
    $timer = 0
    while ($timer -lt $Timeout) {
        $vm = Get-AzVM -ResourceGroupName $ResourceGroupName -Name $VMName -ErrorAction SilentlyContinue
        if ($vm) {
            return $true
        }
        Start-Sleep -Seconds 10
        $timer += 10
    }
    return $false
}

# Check if resource group exists
if (-not (Get-AzResourceGroup -Name $resourceGroup -ErrorAction SilentlyContinue)) {
    Write-Host "Creating resource group $resourceGroup"
    New-AzResourceGroup -Name $resourceGroup -Location $location
} else {
    Write-Host "Resource group $resourceGroup already exists"
}

# Check if storage account exists, create if it doesn't
$storageAccount = Get-AzStorageAccount -ResourceGroupName $resourceGroup -Name $storageAccountName -ErrorAction SilentlyContinue
if (-not $storageAccount) {
    Write-Host "Creating storage account $storageAccountName"
    $storageAccount = New-AzStorageAccount -ResourceGroupName $resourceGroup -Name $storageAccountName -Location $location -SkuName Standard_LRS
} else {
    Write-Host "Storage account $storageAccountName already exists"
}

# Create the virtual network
$vnet = @{
    ResourceGroupName = $resourceGroup
    Location = $location
    Name = $vnetName
    AddressPrefix = "10.0.0.0/16"
}
New-AzVirtualNetwork @vnet

# Retrieve the created virtual network object
$virtualNetwork = Get-AzVirtualNetwork -Name $vnetName -ResourceGroupName $resourceGroup

# Check if the subnet already exists before adding it
$existingSubnet = Get-AzVirtualNetworkSubnetConfig -VirtualNetwork $virtualNetwork -Name $subnetName -ErrorAction SilentlyContinue

if (-not $existingSubnet) {
    Write-Host "Adding subnet $subnetName to virtual network $vnetName"
    Add-AzVirtualNetworkSubnetConfig `
        -VirtualNetwork $virtualNetwork `
        -AddressPrefix "10.0.1.0/24" `
        -Name $subnetName | Out-Null

    # Apply changes to the virtual network
    Set-AzVirtualNetwork -VirtualNetwork $virtualNetwork | Out-Null
} else {
    Write-Host "Subnet $subnetName already exists in virtual network $vnetName"
}

# Check if Network Interface exists
$nicName = "$($vmName)VMNic"
$nic = Get-AzNetworkInterface -ResourceGroupName $resourceGroup -Name $nicName -ErrorAction SilentlyContinue
if (-not $nic) {
    Write-Host "Creating network interface $nicName"
    $subnetId = (Get-AzVirtualNetwork -ResourceGroupName $resourceGroup -Name $vnetName).Subnets[0].Id
    $nic = New-AzNetworkInterface -Name $nicName -ResourceGroupName $resourceGroup -Location $location -SubnetId $subnetId
} else {
    Write-Host "Network interface $nicName already exists"
}

# Check if Public IP exists
$publicIp = Get-AzPublicIpAddress -ResourceGroupName $resourceGroup -Name $publicIpName -ErrorAction SilentlyContinue
if (-not $publicIp) {
    Write-Host "Creating public IP address $publicIpName"
    $publicIp = New-AzPublicIpAddress -Name $publicIpName -ResourceGroupName $resourceGroup -Location $location -AllocationMethod Static -Sku Standard
} else {
    Write-Host "Public IP address $publicIpName already exists"
}

# Connect the public IP to the VMNic
$nic = Get-AzNetworkInterface -ResourceGroupName $resourceGroup -Name "$($vmName)VMNic"
$nic.IpConfigurations[0].PublicIpAddress = $publicIp
Set-AzNetworkInterface -NetworkInterface $nic

# Check if VM exists
$vm = Get-AzVM -ResourceGroupName $resourceGroup -Name $vmName -ErrorAction SilentlyContinue
if (-not $vm) {
    Write-Host "Creating VM $vmName"
    
    # Create VM Configuration
    $vmConfig = New-AzVMConfig -VMName $vmName -VMSize "Standard_B2s"
    $vmConfig = Set-AzVMOperatingSystem -VM $vmConfig -Windows -ComputerName $vmName -Credential (New-Object System.Management.Automation.PSCredential($adminUsername, (ConvertTo-SecureString $adminPassword -AsPlainText -Force))) -ProvisionVMAgent -EnableAutoUpdate
    $vmConfig = Set-AzVMSourceImage -VM $vmConfig -PublisherName "MicrosoftWindowsServer" -Offer "WindowsServer" -Skus "2022-Datacenter" -Version "latest"
    $vmConfig = Add-AzVMNetworkInterface -VM $vmConfig -Id $nic.Id
    $vmConfig = Set-AzVMOSDisk -VM $vmConfig -Windows -Caching ReadWrite -CreateOption FromImage -DiskSizeInGB 128 -Name "$($vmName)OSDisk"

    # Set boot diagnostics to use the specified storage account
    $vmConfig = Set-AzVMBootDiagnostic -VM $vmConfig -Enable -StorageAccountName $storageAccountName -ResourceGroupName $resourceGroup

    # Create the VM with Azure Hybrid Benefit
    New-AzVM -ResourceGroupName $resourceGroup -Location $location -VM $vmConfig -LicenseType "Windows_Server"

    # Wait for the VM to be created
    if (Wait-ForVM -ResourceGroupName $resourceGroup -VMName $vmName) {
        Write-Host "VM $vmName created successfully."
    } else {
        Write-Error "Failed to create VM $vmName."
        exit
    }

    # Change RDP port to 10443
    $portvalue = 10443
    $scriptBlock = {
        Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -name "PortNumber" -Value $using:portvalue
        New-NetFirewallRule -DisplayName 'RDPPORTLatest-TCP-In' -Profile 'Public' -Direction Inbound -Action Allow -Protocol TCP -LocalPort $using:portvalue
        New-NetFirewallRule -DisplayName 'RDPPORTLatest-UDP-In' -Profile 'Public' -Direction Inbound -Action Allow -Protocol UDP -LocalPort $using:portvalue
        Restart-Service -Name TermService -Force
    }
    Invoke-AzVMRunCommand -ResourceGroupName $resourceGroup -VMName $vmName -CommandId 'RunPowerShellScript' -ScriptString $scriptBlock

    # Update NSG to allow traffic on port 10443
    $nsg = Get-AzNetworkSecurityGroup -ResourceGroupName $resourceGroup
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
        -DestinationPortRange 10443
    $nsg | Set-AzNetworkSecurityGroup

    # Set up auto-shutdown
    $shutdownSchedule = @{
        "location" = $location
        "properties" = @{
            "status" = "Enabled"
            "taskType" = "ComputeVmShutdownTask"
            "dailyRecurrence" = @{"time" = "2100"}
            "timeZoneId" = "US Mountain Standard Time"
            "notificationSettings" = @{"status" = "Disabled"}
            "targetResourceId" = (Get-AzVM -ResourceGroupName $resourceGroup -Name $vmName).Id
        }
    }
    New-AzResource -ResourceId "/subscriptions/$((Get-AzContext).Subscription.Id)/resourceGroups/$resourceGroup/providers/microsoft.devtestlab/schedules/shutdown-computevm-$vmName" -Location $location -Properties $shutdownSchedule.properties -Force

    # Create PowerShell script for AD DS installation and configuration
    $script = @'
# Install AD DS role
Install-WindowsFeature -Name AD-Domain-Services -IncludeManagementTools

# Promote to Domain Controller
Import-Module ADDSDeployment
Install-ADDSForest `
    -CreateDnsDelegation:$false `
    -DatabasePath "C:\Windows\NTDS" `
    -DomainMode "WinThreshold" `
    -DomainName "$domainName" `
    -ForestMode "WinThreshold" `
    -InstallDns:$true `
    -LogPath "C:\Windows\NTDS" `
    -NoRebootOnCompletion:$false `
    -SysvolPath "C:\Windows\SYSVOL" `
    -Force:$true `
    -SafeModeAdministratorPassword (ConvertTo-SecureString '$adminPassword' -AsPlainText -Force)

# Wait for AD DS installation to complete
Start-Sleep -Seconds 300

# Create 10 test users in the domain
for ($i = 1; $i -le 10; $i++) {
    $username = "TestUser$i"
    $password = "TestPassword123!"
    New-ADUser -Name $username -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) -PasswordNeverExpires $true -Enabled $true
}
'@

    # Execute the PowerShell script on the VM using RunCommand
    Invoke-AzVMRunCommand -ResourceGroupName $resourceGroup -Name $vmName -CommandId "RunPowerShellScript" -ScriptString $script

    # Ensure VNet exists before updating DNS servers
    $vnet = Get-AzVirtualNetwork -ResourceGroupName $resourceGroup -Name $vnetName
    if ($vnet -ne $null) {
        Write-Host "Updating VNet DNS servers"
        $vnet.DhcpOptions.DnsServers.Add("10.0.1.4")
        $vnet | Set-AzVirtualNetwork
    } else {
        Write-Error "VNet $vnetName not found."
    }

    Write-Host "Domain Controller setup complete. The VM will restart to finish the AD DS installation and create test users."
} else {
    Write-Host "VM $vmName already exists"
}
