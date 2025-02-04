# create-testdomaincontroller.ps1
# Updated script with explanations, contexts, and comments to help clarify each step.

# ==========================================================================================
# 1. SET ERROR PREFERENCES
# ==========================================================================================
# Halt script execution on the first error encountered to prevent partial/configuration
# that doesn't fully succeed.
$ErrorActionPreference = "Stop"

# ==========================================================================================
# 2. DEFINE VARIABLES
# ==========================================================================================
# These variables store basic configuration details, such as resource names and credentials.
# Adjust them to match your environment and requirements.
$resourceGroup       = "JB-TEST-RG"
$location            = "westus"
$automationLocation  = "westus2"       # Location where the Azure Automation account will be created
$vnetName            = "JB-TEST-VNET"  # Virtual Network name
$subnetName          = "JB-TEST-SUBNET1"
$vmName              = "JB-TEST-DC01"  # VM name (Domain Controller)
$adminUsername       = "jbadmin"       # Administrator username for the VM
$adminPassword       = "TS=pGxB~8m^A~WH^[yB8"  # Administrator password
$domainName          = "JB-TEST.local"       # Domain name for the AD DS environment
$publicIpName        = "$vmName-PublicIP"     # Public IP name for the domain controller
$storageAccountName  = "jbteststorage0"       # Existing or new Storage account name
$fileShareName       = "runbooks"             # Fileshare container for storing runbook
$subdirectoryName    = "AutoShutdownRunbook"
$runbookFileName     = "runbook.ps1"
$tempRunbookFilePath = "C:\Temp\$([System.Guid]::NewGuid().ToString()).ps1"
$nsgName             = "JB-TEST-NSG"          # Name for the Network Security Group
$automationAccountName = "JB-TEST-Automation" # Azure Automation account name
$runbookName         = "AutoShutdownRunbook"  # Name of the runbook
$policyName          = "RunbookAccessPolicy"  # Stored Access Policy name for generating SAS tokens

# ==========================================================================================
# 3. IMPORT REQUIRED MODULES
# ==========================================================================================
# Az.Automation and Az.Storage modules provide PowerShell commands for Azure Automation
# and Azure Storage interactions. Ensure these modules are installed or upgrade them if
# you encounter any parameter errors (e.g., the -ApiVersion parameter).
Import-Module Az.Automation
Import-Module Az.Storage

# ==========================================================================================
# 4. SUPPORTING FUNCTION
# ==========================================================================================
# This function waits until a VM appears in the Azure resource group or times out.
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

# ==========================================================================================
# 5. CREATE OR CONFIRM RESOURCE GROUP
# ==========================================================================================
# The resource group holds all your Azure resources. This checks if one exists already
# or creates it if needed.
if (-not (Get-AzResourceGroup -Name $resourceGroup -ErrorAction SilentlyContinue)) {
    Write-Host "Creating resource group $resourceGroup..."
    New-AzResourceGroup -Name $resourceGroup -Location $location
} else {
    Write-Host "Resource group $resourceGroup already exists."
}

# ==========================================================================================
# 6. CREATE OR CONFIRM STORAGE ACCOUNT
# ==========================================================================================
# Check if the storage account exists; if not, create it. 
$storageAccount = Get-AzStorageAccount -ResourceGroupName $resourceGroup -Name $storageAccountName -ErrorAction SilentlyContinue
if (-not $storageAccount) {
    Write-Host "Creating storage account $storageAccountName..."
    $storageAccount = New-AzStorageAccount -ResourceGroupName $resourceGroup -Name $storageAccountName -Location $location -SkuName Standard_LRS
} else {
    Write-Host "Storage account $storageAccountName already exists."
}

# ==========================================================================================
# 7. CREATE STORAGE CONTEXT (UPDATING x-ms-version)
# ==========================================================================================
# Retrieve the storage account key and create a storage context. If your Az.Storage module
# doesn't support the -ApiVersion parameter, consider upgrading the module or remove it.
$storageKey = (Get-AzStorageAccountKey -ResourceGroupName $resourceGroup -Name $storageAccountName)[0].Value
try {
    $storageAccountContext = New-AzStorageContext -StorageAccountName $storageAccountName -StorageAccountKey $storageKey -Protocol 'HTTPS' -ApiVersion '2021-08-06'
} catch {
    Write-Host "Your Az.Storage module might not support -ApiVersion. Falling back to default."
    $storageAccountContext = New-AzStorageContext -StorageAccountName $storageAccountName -StorageAccountKey $storageKey -Protocol 'HTTPS'
}

# Create a Temp directory if it doesn't exist for runbook files
if (-not (Test-Path -Path "C:\Temp")) {
    New-Item -ItemType Directory -Path "C:\Temp"
}

# ==========================================================================================
# 8. CREATE RUNBOOK CONTENT
# ==========================================================================================
# This snippet is the PowerShell runbook that automatically stops a VM. We write it to a temp
# file, then upload it to the Azure file share.
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

# Write the runbook content to a temporary file
Set-Content -Path $tempRunbookFilePath -Value $runbookContent

# Check if a runbook file already exists and remove it
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

# ==========================================================================================
# 9. CREATE OR CONFIRM STORED ACCESS POLICY & GENERATE SAS
# ==========================================================================================
# A stored access policy can be used to generate a SAS token for controlled read/write
# access to resources.
try {
    $policy = Get-AzStorageShareStoredAccessPolicy -Context $storageAccountContext -ShareName $fileShareName -Policy $policyName
    Write-Host "Stored access policy $policyName already exists."
} catch {
    Write-Host "Creating stored access policy $policyName..."
    $policy = New-AzStorageShareStoredAccessPolicy -Context $storageAccountContext -ShareName $fileShareName -Policy $policyName -Permission r -StartTime (Get-Date).AddMinutes(-5) -ExpiryTime (Get-Date).AddHours(1)
}

# Generate a Shared Access Signature token that references the stored access policy
$sasToken = New-AzStorageShareSASToken -Context $storageAccountContext -ShareName $fileShareName -Policy $policyName

# ==========================================================================================
# 10. CREATE OR CONFIRM VIRTUAL NETWORK & SUBNET
# ==========================================================================================
# The virtual network must already exist or will be created; then we ensure a subnet is
# present for the VM.
$virtualNetwork = Get-AzVirtualNetwork -Name $vnetName -ResourceGroupName $resourceGroup -ErrorAction SilentlyContinue
if (-not $virtualNetwork) {
    Write-Host "Creating virtual network $vnetName..."
    $vnetParams = @{
        ResourceGroupName = $resourceGroup
        Location          = $location
        Name              = $vnetName
        AddressPrefix     = "10.0.0.0/16"
    }
    New-AzVirtualNetwork @vnetParams
    $virtualNetwork = Get-AzVirtualNetwork -Name $vnetName -ResourceGroupName $resourceGroup
} else {
    Write-Host "Virtual network $vnetName already exists."
}

$existingSubnet = Get-AzVirtualNetworkSubnetConfig -VirtualNetwork $virtualNetwork -Name $subnetName -ErrorAction SilentlyContinue
if (-not $existingSubnet) {
    Write-Host "Adding subnet $subnetName to virtual network $vnetName..."
    Add-AzVirtualNetworkSubnetConfig -VirtualNetwork $virtualNetwork -AddressPrefix "10.0.1.0/24" -Name $subnetName | Out-Null
    Set-AzVirtualNetwork -VirtualNetwork $virtualNetwork | Out-Null
} else {
    Write-Host "Subnet $subnetName already exists in virtual network $vnetName."
}

# ==========================================================================================
# 11. CREATE OR CONFIRM NETWORK SECURITY GROUP
# ==========================================================================================
try {
    $nsg = Get-AzNetworkSecurityGroup -ResourceGroupName $resourceGroup -Name $nsgName
    Write-Host "Network Security Group $nsgName already exists."
} catch {
    Write-Host "Creating Network Security Group $nsgName..."
    $nsg = New-AzNetworkSecurityGroup -ResourceGroupName $resourceGroup -Location $location -Name $nsgName
}

# ==========================================================================================
# 12. CREATE OR CONFIRM NETWORK INTERFACE & PUBLIC IP
# ==========================================================================================
$nicName = "$($vmName)VMNic"   # NIC name is generally VM name + "VMNic"
try {
    $nic = Get-AzNetworkInterface -ResourceGroupName $resourceGroup -Name $nicName
    Write-Host "Network interface $nicName already exists."
} catch {
    Write-Host "Creating network interface $nicName..."
    $subnetId = (Get-AzVirtualNetwork -ResourceGroupName $resourceGroup -Name $vnetName).Subnets[0].Id
    $nic = New-AzNetworkInterface -Name $nicName -ResourceGroupName $resourceGroup -Location $location -SubnetId $subnetId
}

try {
    $publicIp = Get-AzPublicIpAddress -ResourceGroupName $resourceGroup -Name $publicIpName
    Write-Host "Public IP address $publicIpName already exists."
} catch {
    Write-Host "Creating public IP address $publicIpName..."
    $publicIp = New-AzPublicIpAddress -Name $publicIpName -ResourceGroupName $resourceGroup -Location $location -AllocationMethod Static -Sku Standard
}

# Attach the public IP to the NIC
$nic = Get-AzNetworkInterface -ResourceGroupName $resourceGroup -Name "$($vmName)VMNic"
$nic.IpConfigurations[0].PublicIpAddress = $publicIp
Set-AzNetworkInterface -NetworkInterface $nic | Out-Null

# ==========================================================================================
# 13. CREATE OR CONFIRM VM
# ==========================================================================================
# This section creates the VM if it doesn't exist and configures it with Windows, labs, etc.
try {
    $vm = Get-AzVM -ResourceGroupName $resourceGroup -Name $vmName -ErrorAction SilentlyContinue
    if ($vm) {
        Write-Host "VM $vmName already exists."
    } else {
        throw "NotFound"
    }
} catch {
    Write-Host "Creating VM $vmName..."
    # Build the VM configuration
    $vmConfig = New-AzVMConfig -VMName $vmName -VMSize "Standard_B2s"
    $vmConfig = Set-AzVMOperatingSystem -VM $vmConfig -Windows `
        -ComputerName $vmName `
        -Credential (New-Object System.Management.Automation.PSCredential($adminUsername, (ConvertTo-SecureString $adminPassword -AsPlainText -Force)))
    $vmConfig = Set-AzVMSourceImage -VM $vmConfig -PublisherName "MicrosoftWindowsServer" -Offer "WindowsServer" -Skus "2025-datacenter-core-g2" -Version "latest"
    $vmConfig = Add-AzVMNetworkInterface -VM $vmConfig -Id $nic.Id
    $vmConfig = Set-AzVMOSDisk -VM $vmConfig -Windows -Caching ReadWrite -CreateOption FromImage -DiskSizeInGB 128 -Name "$($vmName)OSDisk"

    # Enable boot diagnostics referencing the storage account
    $vmConfig = Set-AzVMBootDiagnostic -VM $vmConfig -Enable -StorageAccountName $storageAccountName -ResourceGroupName $resourceGroup

    # Enable Trusted Launch features for improved security posture
    $vmConfig.SecurityProfile = @{
        SecurityType = "TrustedLaunch"
        UefiSettings = @{
            SecureBootEnabled = $true
            VTpmEnabled       = $true
        }
    }

    # Create the VM with Azure Hybrid Benefit for licensing
    New-AzVM -ResourceGroupName $resourceGroup -Location $location -VM $vmConfig -LicenseType "Windows_Server" | Out-Null

    # Wait until the VM is fully created
    if (Wait-ForVM -ResourceGroupName $resourceGroup -VMName $vmName) {
        Write-Host "VM $vmName created successfully."
    } else {
        Write-Error "Failed to create VM $vmName within the specified time."
        exit
    }

    # Change RDP port to 10443 (example of customizing firewall ports)
    $portvalue = 10443
    $scriptBlock = {
        Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -Name "PortNumber" -Value $using:portvalue
        New-NetFirewallRule -DisplayName 'RDPPORTLatest-TCP-In' -Profile 'Public' -Direction Inbound -Action Allow -Protocol TCP -LocalPort $using:portvalue
        New-NetFirewallRule -DisplayName 'RDPPORTLatest-UDP-In' -Profile 'Public' -Direction Inbound -Action Allow -Protocol UDP -LocalPort $using:portvalue
        Restart-Service -Name TermService -Force
    }

    # Execute the script within the new VM
    Invoke-AzVMRunCommand -ResourceGroupName $resourceGroup -VMName $vmName -CommandId 'RunPowerShellScript' -ScriptString $scriptBlock

    # Create a new rule in the NSG to allow traffic on port 10443
    Add-AzNetworkSecurityRuleConfig -NetworkSecurityGroup $nsg `
        -Name "Allow_RDP_10443" `
        -Description "Allow RDP on port 10443" `
        -Access Allow `
        -Protocol Tcp `
        -Direction Inbound `
        -Priority 1001 `
        -SourceAddressPrefix * `
        -SourcePortRange * `
        -DestinationAddressPrefix * `
        -DestinationPortRange 10443 | Out-Null
    $nsg | Set-AzNetworkSecurityGroup | Out-Null
}

# ==========================================================================================
# 14. CREATE OR CONFIRM AZURE AUTOMATION ACCOUNT & RUNBOOK
# ==========================================================================================
# Azure Automation can schedule runbooks to automatically manage resource states. 
try {
    $automationAccount = Get-AzAutomationAccount -ResourceGroupName $resourceGroup -Name $automationAccountName -ErrorAction SilentlyContinue
    if ($automationAccount) {
        Write-Host "Azure Automation account $automationAccountName already exists."
    } else {
        throw "NotFound"
    }
} catch {
    Write-Host "Creating Azure Automation account $automationAccountName..."
    $automationAccount = New-AzAutomationAccount -ResourceGroupName $resourceGroup -Name $automationAccountName -Location $automationLocation
}

# Remove existing runbook if it exists, then create a new placeholder
try {
    $existingRunbook = Get-AzAutomationRunbook -AutomationAccountName $automationAccountName -Name $runbookName -ResourceGroupName $resourceGroup -ErrorAction SilentlyContinue
    if ($existingRunbook) {
        Write-Host "Runbook $runbookName already exists. Deleting it..."
        Remove-AzAutomationRunbook -AutomationAccountName $automationAccountName -Name $runbookName -ResourceGroupName $resourceGroup -Force
    } else {
        Write-Host "Runbook $runbookName does not exist. Proceeding..."
    }
} catch {
    Write-Host "Runbook $runbookName does not exist. Proceeding..."
}
New-AzAutomationRunbook -AutomationAccountName $automationAccountName -Name $runbookName -ResourceGroupName $resourceGroup -Type PowerShellWorkflow | Out-Null

# ==========================================================================================
# 15. DOWNLOAD RUNBOOK CONTENT AND IMPORT TO AUTOMATION
# ==========================================================================================
# Create a local file path and use Invoke-WebRequest with the SAS token to retrieve runbook from the file share
$downloadPath = "C:\Temp\$([System.Guid]::NewGuid().ToString()).ps1"
$sasUri = "https://$storageAccountName.file.core.windows.net/$fileShareName/$remoteFilePath?$sasToken"

try {
    Invoke-WebRequest -Uri $sasUri -OutFile $downloadPath
    Write-Host "Runbook content downloaded successfully."
} catch {
    Write-Error "Failed to download runbook content. Error: $_"
    exit
}

# Import the runbook content from the downloaded file
Import-AzAutomationRunbook -Path $downloadPath -Name $runbookName -Type PowerShellWorkflow -ResourceGroupName $resourceGroup -AutomationAccountName $automationAccountName

# Publish the runbook so it’s ready to run or schedule
Publish-AzAutomationRunbook -Name $runbookName -ResourceGroupName $resourceGroup -AutomationAccountName $automationAccountName -Force
Write-Host "Runbook $runbookName published successfully."

# ==========================================================================================
# 16. CREATE SCHEDULE & REGISTER RUNBOOK
# ==========================================================================================
# Example: create a one-time schedule that starts 5 minutes from now for demonstration.
$scheduleName = "AutoShutdownSchedule"
$startTime = (Get-Date).AddMinutes(5)

Write-Host "Creating schedule $scheduleName..."
New-AzAutomationSchedule -AutomationAccountName $automationAccountName -Name $scheduleName -StartTime $startTime -OneTime | Out-Null

Write-Host "Registering runbook $runbookName with schedule $scheduleName..."
Register-AzAutomationScheduledRunbook -AutomationAccountName $automationAccountName -Name $runbookName -ScheduleName $scheduleName -ResourceGroupName $resourceGroup
Write-Host "Auto-shutdown schedule created successfully."

# ==========================================================================================
# 17. INSTALL AND CONFIGURE AD DS ON THE VM
# ==========================================================================================
# The following script installs Windows feature AD-Domain-Services, promotes the server
# to a domain controller, and creates test users.
$script = @'
# Install AD DS role
Install-WindowsFeature -Name AD-Domain-Services -IncludeManagementTools

# Promote to Domain Controller
Import-Module ADDSDeployment
Install-ADDSForest `
    -CreateDnsDelegation:$false `
    -DatabasePath "C:\Windows\NTDS" `
    -DomainMode "WinThreshold" `
    -DomainName "JB-TEST.local" `
    -ForestMode "WinThreshold" `
    -InstallDns:$true `
    -LogPath "C:\Windows\NTDS" `
    -NoRebootOnCompletion:$false `
    -SysvolPath "C:\Windows\SYSVOL" `
    -Force:$true `
    -SafeModeAdministratorPassword (ConvertTo-SecureString 'TS=pGxB~8m^A~WH^[yB8' -AsPlainText -Force)

# Wait for AD DS installation to complete
Start-Sleep -Seconds 300

# Create 10 test users in the domain as an example
for ($i = 1; $i -le 10; $i++) {
    $username = "TestUser$i"
    $password = "TestPassword123!"
    New-ADUser -Name $username -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) `
               -PasswordNeverExpires $true -Enabled $true
}
'@

Invoke-AzVMRunCommand -ResourceGroupName $resourceGroup -Name $vmName -CommandId "RunPowerShellScript" -ScriptString $script

# ==========================================================================================
# 18. UPDATE DNS SETTINGS IN THE VNET
# ==========================================================================================
# After creating a domain, we typically set the DC as the DNS server for the VNet if needed.
$vnet = Get-AzVirtualNetwork -ResourceGroupName $resourceGroup -Name $vnetName
if ($vnet) {
    Write-Host "Updating $vnetName DNS servers to point to Domain Controller (10.0.1.4 by default)."
    $vnet.DhcpOptions.DnsServers.Add("10.0.1.4")
    $vnet | Set-AzVirtualNetwork
} else {
    Write-Error "VNet $vnetName not found; cannot update DNS servers."
}

Write-Host "Domain Controller setup complete. The VM will restart to finish AD DS installation and create test users."