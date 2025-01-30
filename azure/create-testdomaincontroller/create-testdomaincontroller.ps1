# Set ErrorActionPreference to Stop to halt script execution on the first error
$ErrorActionPreference = "Stop"

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
$fileShareName = "runbooks"
$subdirectoryName = "AutoShutdownRunbook"
$runbookFileName = "runbook.ps1"
$tempRunbookFilePath = "C:\Temp\$([System.Guid]::NewGuid().ToString()).ps1"
$nsgName = "JB-TEST-NSG"
$automationAccountName = "JB-TEST-Automation"
$runbookName = "AutoShutdownRunbook"

# Import Az.Automation and Az.Storage modules
Import-Module Az.Automation
Import-Module Az.Storage

# Function to wait for VM creation
function Wait-ForVM {
    param (
        [string]$ResourceGroupName,
        [string]$VMName,
        [int]$Timeout = 600
    )
    $timer = 0
    while ($timer -lt $Timeout) {
        $vm = Get-AzVM -ResourceGroupName $ResourceGroupName -Name $VMName -ErrorAction Stop
        if ($vm) {
            return $true
        }
        Start-Sleep -Seconds 10
        $timer += 10
    }
    return $false
}

# Check if resource group exists
if (-not (Get-AzResourceGroup -Name $resourceGroup -ErrorAction Stop)) {
    Write-Host "Creating resource group $resourceGroup"
    New-AzResourceGroup -Name $resourceGroup -Location $location
} else {
    Write-Host "Resource group $resourceGroup already exists"
}

# Check if storage account exists, create if it doesn't
$storageAccount = Get-AzStorageAccount -ResourceGroupName $resourceGroup -Name $storageAccountName -ErrorAction Stop
if (-not $storageAccount) {
    Write-Host "Creating storage account $storageAccountName"
    $storageAccount = New-AzStorageAccount -ResourceGroupName $resourceGroup -Name $storageAccountName -Location $location -SkuName Standard_LRS
} else {
    Write-Host "Storage account $storageAccountName already exists"
}

# Get the storage account context
$storageAccountContext = $storageAccount.Context

# Ensure the directory exists
if (-not (Test-Path -Path "C:\Temp")) {
    New-Item -ItemType Directory -Path "C:\Temp"
}

# Write the runbook content to a temporary file
$runbookContent = @"
workflow $runbookName {
    param (
        [string] \$resourceGroupName,
        [string] \$vmName
    )

    \$connection = Get-AutomationConnection -Name AzureRunAsConnection
    Add-AzAccount -ServicePrincipal -TenantId \$connection.TenantId -ApplicationId \$connection.ApplicationId -CertificateThumbprint \$connection.CertificateThumbprint

    Stop-AzVM -ResourceGroupName \$resourceGroupName -Name \$vmName -Force
}
"@
Set-Content -Path $tempRunbookFilePath -Value $runbookContent

# Delete the existing runbook file if it exists
$remoteFilePath = "$subdirectoryName/$runbookFileName"
$fileExists = Get-AzStorageFile -Context $storageAccountContext -ShareName $fileShareName -Path $remoteFilePath -ErrorAction SilentlyContinue

if ($fileExists) {
    Remove-AzStorageFile -Context $storageAccountContext -ShareName $fileShareName -Path $remoteFilePath
    Write-Host "Existing runbook file deleted."
} else {
    Write-Host "No existing runbook file to delete."
}

# Upload the new runbook file to the file share
Set-AzStorageFileContent -Context $storageAccountContext -ShareName $fileShareName -Source $tempRunbookFilePath -Path $remoteFilePath

Write-Host "Runbook content written to file share successfully."

# Clean up the temporary file
Remove-Item -Path $tempRunbookFilePath

# Check if virtual network exists, create if it doesn't
$virtualNetwork = Get-AzVirtualNetwork -Name $vnetName -ResourceGroupName $resourceGroup -ErrorAction Stop
if (-not $virtualNetwork) {
    Write-Host "Creating virtual network $vnetName"
    $vnet = @{
        ResourceGroupName = $resourceGroup
        Location = $location
        Name = $vnetName
        AddressPrefix = "10.0.0.0/16"
    }
    New-AzVirtualNetwork @vnet

    # Retrieve the created virtual network object
    $virtualNetwork = Get-AzVirtualNetwork -Name $vnetName -ResourceGroupName $resourceGroup
} else {
    Write-Host "Virtual network $vnetName already exists"
}

# Check if the subnet already exists before adding it
$existingSubnet = Get-AzVirtualNetworkSubnetConfig -VirtualNetwork $virtualNetwork -Name $subnetName -ErrorAction Stop
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

# Check if Network Security Group exists, create if it doesn't
try {
    $nsg = Get-AzNetworkSecurityGroup -ResourceGroupName $resourceGroup -Name $nsgName -ErrorAction Stop
    Write-Host "Network Security Group $nsgName already exists"
} catch {
    Write-Host "Creating Network Security Group $nsgName"
    $nsg = New-AzNetworkSecurityGroup -ResourceGroupName $resourceGroup -Location $location -Name $nsgName
}

# Check if Network Interface exists, create if it doesn't
$nicName = "$($vmName)VMNic"
try {
    $nic = Get-AzNetworkInterface -ResourceGroupName $resourceGroup -Name $nicName -ErrorAction Stop
    Write-Host "Network interface $nicName already exists"
} catch {
    Write-Host "Creating network interface $nicName"
    $subnetId = (Get-AzVirtualNetwork -ResourceGroupName $resourceGroup -Name $vnetName).Subnets[0].Id
    $nic = New-AzNetworkInterface -Name $nicName -ResourceGroupName $resourceGroup -Location $location -SubnetId $subnetId
}

# Check if Public IP exists, create if it doesn't
try {
    $publicIp = Get-AzPublicIpAddress -ResourceGroupName $resourceGroup -Name $publicIpName -ErrorAction Stop
    Write-Host "Public IP address $publicIpName already exists"
} catch {
    Write-Host "Creating public IP address $publicIpName"
    $publicIp = New-AzPublicIpAddress -Name $publicIpName -ResourceGroupName $resourceGroup -Location $location -AllocationMethod Static -Sku Standard
}

# Connect the public IP to the VMNic
$nic = Get-AzNetworkInterface -ResourceGroupName $resourceGroup -Name "$($vmName)VMNic" -ErrorAction Stop
$nic.IpConfigurations[0].PublicIpAddress = $publicIp
Set-AzNetworkInterface -NetworkInterface $nic

# Check if VM exists, create if it doesn't
try {
    $vm = Get-AzVM -ResourceGroupName $resourceGroup -Name $vmName -ErrorAction Stop
    Write-Host "VM $vmName already exists"
} catch {
    Write-Host "Creating VM $vmName"

    # Create VM Configuration
    $vmConfig = New-AzVMConfig -VMName $vmName -VMSize "Standard_B2s"
    $vmConfig = Set-AzVMOperatingSystem -VM $vmConfig -Windows -ComputerName $vmName -Credential (New-Object System.Management.Automation.PSCredential($adminUsername, (ConvertTo-SecureString $adminPassword -AsPlainText -Force)))
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
}

# Check if Azure Automation account exists, create if it doesn't
try {
    $automationAccount = Get-AzAutomationAccount -ResourceGroupName $resourceGroup -Name $automationAccountName -ErrorAction Stop
    Write-Host "Azure Automation account $automationAccountName already exists"
} catch {
    Write-Host "Creating Azure Automation account $automationAccountName"
    
    # Ensure the location is valid for Azure Automation
    $validLocations = @("East US", "West US", "West US 2", "Central US", "North Central US", "South Central US", "East US 2", "Canada Central", "Canada East", "North Europe", "West Europe", "Germany Central", "Germany Northeast", "Switzerland North", "Switzerland West", "Norway East", "Norway West", "UK South", "UK West", "France Central", "France South", "Australia East", "Australia Southeast", "Australia Central", "Australia Central 2", "Southeast Asia", "East Asia", "Japan East", "Japan West", "Korea Central", "Korea South", "India Central", "India South", "India West", "South Africa North", "South Africa West", "Brazil South", "US Gov Virginia", "US Gov Arizona", "US Gov Texas", "US DoD Central", "US DoD East", "China East", "China East 2", "China North", "China North 2")
    
    if ($validLocations -contains $location) {
        $automationAccount = New-AzAutomationAccount -ResourceGroupName $resourceGroup -Name $automationAccountName -Location $location
    } else {
        Write-Error "Location $location is not valid for Azure Automation account."
        exit
    }
}

# Create the runbook
New-AzAutomationRunbook -AutomationAccountName $automationAccountName -Name $runbookName -ResourceGroupName $resourceGroup -Type PowerShellWorkflow -Force

# Download the runbook content
$headers = @{
    "x-ms-version" = "2022-04-11"
}
$downloadPath = "C:\Temp\$([System.Guid]::NewGuid().ToString()).ps1"
Invoke-WebRequest -Uri "https://$storageAccountName.file.core.windows.net/$fileShareName/$remoteFilePath" -Headers $headers -OutFile $downloadPath

# Import the runbook content from the downloaded file
Import-AzAutomationRunbook -Path $downloadPath -Name $runbookName -Type PowerShellWorkflow -ResourceGroupName $resourceGroup -AutomationAccountName $automationAccountName -Force

# Publish the runbook
Publish-AzAutomationRunbook -Name $runbookName -ResourceGroupName $resourceGroup -AutomationAccountName $automationAccountName -Force

# Define the schedule parameters
$scheduleName = "AutoShutdownSchedule"
$startTime = (Get-Date).AddMinutes(5) # Start in 5 minutes

# Create a new schedule
Write-Host "Creating schedule $scheduleName"
New-AzAutomationSchedule -AutomationAccountName $automationAccountName -Name $scheduleName -StartTime $startTime -OneTime

# Register the runbook with the schedule
Write-Host "Registering runbook $runbookName with schedule $scheduleName"
Register-AzAutomationScheduledRunbook -AutomationAccountName $automationAccountName -Name $runbookName -ScheduleName $scheduleName -ResourceGroupName $resourceGroup

Write-Host "Auto-shutdown schedule created successfully."

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
$vnet = Get-AzVirtualNetwork -ResourceGroupName $resourceGroup -Name $vnetName -ErrorAction Stop
if ($vnet -ne $null) {
    Write-Host "Updating VNet DNS servers"
    $vnet.DhcpOptions.DnsServers.Add("10.0.1.4")
    $vnet | Set-AzVirtualNetwork
} else {
    Write-Error "VNet $vnetName not found."
}

Write-Host "Domain Controller setup complete. The VM will restart to finish the AD DS installation and create test users."