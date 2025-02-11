# =============================================================================
# Script: Create-TestDomainController.ps1
# Created: 2025-02-10 22:50:04 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-11 02:06:55 UTC
# Updated By: jdyer-nuvodia
# Version: 2.0
# Additional Info: Consolidated all pre-deployment validations into validation phase
# =============================================================================

<# 
.SYNOPSIS
    Creates a test domain controller as a Trusted Launch VM in Azure.
.DESCRIPTION
    This script provisions a domain controller VM configured as a Trusted Launch VM in Azure.
    It first validates all components and dependencies, then proceeds with the actual deployment
    only if all validations pass. The script uses a two-phase approach:
    
    Phase 1: Validation
    - Validates all required modules
    - Checks resource name availability
    - Verifies permissions
    - Validates VM size availability
    - Checks network configuration
    - Validates existing resources
    
    Phase 2: Deployment
    - Creates or verifies resource group
    - Sets up storage account
    - Configures networking components
    - Deploys the virtual machine
    - Configures and validates DevTest Lab shutdown schedule
    
    Use -ValidateOnly to perform validation without deployment.
.PARAMETER resourceGroupName
    The name of the resource group where the VM and related resources will be created.
.PARAMETER location
    The Azure region (location) to deploy the resources.
.PARAMETER vmName
    The name of the VM to create.
.PARAMETER VMSize
    The size of the VM (e.g., 'Standard_DS2_v2').
.PARAMETER vnetName
    The virtual network name for the VM.
.PARAMETER subnetName
    The name of the subnet within the virtual network.
.PARAMETER adminUsername
    The administrator username for the VM.
.PARAMETER adminPassword
    The administrator password for the VM.
.PARAMETER ValidateOnly
    Performs validation only without deploying resources.
.EXAMPLE
    PS C:\> .\Create-TestDomainController.ps1 -ValidateOnly
    Performs validation of all components without deployment.
.EXAMPLE
    PS C:\> .\Create-TestDomainController.ps1
    Validates all components and proceeds with deployment if validation succeeds.
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param (
    [Parameter(Mandatory = $false)]
    [string]$resourceGroupName,
    [Parameter(Mandatory = $false)]
    [string]$location,
    [Parameter(Mandatory = $false)]
    [string]$vmName,
    [Parameter(Mandatory = $false)]
    [string]$VMSize,
    [Parameter(Mandatory = $false)]
    [string]$vnetName,
    [Parameter(Mandatory = $false)]
    [string]$subnetName,
    [Parameter(Mandatory = $false)]
    [string]$adminUsername,
    [Parameter(Mandatory = $false)]
    [string]$adminPassword,
    [Parameter(Mandatory = $false)]
    [switch]$ValidateOnly
)

# Start logging
$LogFile = Join-Path $PSScriptRoot "Create-TestDomainController.log"
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'VALIDATION')]
        [string]$Level = 'INFO'
    )
    $LogMessage = "[$([DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))] [$Level] $Message"
    Add-Content -Path $LogFile -Value $LogMessage
    Write-Host $LogMessage
}

# Clear log file
Set-Content -Path $LogFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Log file reset. New log starting."

# Explicitly defined default variables
$DefaultResourceGroupName    = 'JB-TEST-RG2'
$DefaultLocation            = 'westus2'
$DefaultStorageAccountName  = 'jbteststorage0'
$DefaultVnetName           = 'JB-TEST-VNET'
$DefaultSubnetName         = 'JB-TEST-SUBNET1'
$DefaultVmName             = 'JB-TEST-DC01'
$DefaultAdminUsername      = 'jbadmin'
$DefaultAdminPassword      = 'TS-pGxB~8m^A~WH^[yB8'
$DefaultDomainName         = 'JB-TEST.local'
$DefaultPublicIpName       = "$DefaultVmName-PUBIP"
$DefaultNsgName           = 'JB-TEST-NSG'
$DefaultVnetAddressSpace   = '10.0.0.0/16'
$DefaultSubnetAddressSpace = '10.0.1.0/24'

# If parameters are not provided, assign defaults
if (-not $resourceGroupName) { $resourceGroupName = $DefaultResourceGroupName }
if (-not $location)          { $location = $DefaultLocation }
if (-not $vmName)            { $vmName = $DefaultVmName }
if (-not $VMSize)            { $VMSize = 'Standard_DS2_v2' }
if (-not $vnetName)          { $vnetName = $DefaultVnetName }
if (-not $subnetName)        { $subnetName = $DefaultSubnetName }
if (-not $adminUsername)     { $adminUsername = $DefaultAdminUsername }
if (-not $adminPassword)     { $adminPassword = $DefaultAdminPassword }

# Validation function
function Test-AzureResources {
    $validationResults = @{
        Success = $true
        Messages = @()
    }

    try {
        # Validate required modules
        Write-Log "Validating required modules..." -Level VALIDATION
        $requiredModules = @('Az.Accounts', 'Az.Resources', 'Az.Network', 'Az.Storage', 'Az.Compute')
        foreach ($module in $requiredModules) {
            if (!(Get-Module -Name $module -ListAvailable)) {
                $validationResults.Messages += "Required module $module is not installed"
                $validationResults.Success = $false
            }
        }

        # Validate location
        Write-Log "Validating location '$location'..." -Level VALIDATION
        $validLocations = Get-AzLocation
        if ($location -notin $validLocations.Location) {
            $validationResults.Messages += "Invalid location: $location"
            $validationResults.Success = $false
        }

        # Validate VM size availability
        Write-Log "Validating VM size '$VMSize'..." -Level VALIDATION
        $vmSizes = Get-AzVMSize -Location $location
        if ($VMSize -notin $vmSizes.Name) {
            $validationResults.Messages += "VM size $VMSize is not available in $location"
            $validationResults.Success = $false
        }

        # Validate resource group existence and permissions
        Write-Log "Validating resource group access..." -Level VALIDATION
        $rg = Get-AzResourceGroup -Name $resourceGroupName -ErrorAction SilentlyContinue
        if ($rg) {
            try {
                $testResource = New-AzResourceGroupDeployment -ResourceGroupName $resourceGroupName `
                    -TemplateUri "https://raw.githubusercontent.com/Azure/azure-quickstart-templates/master/100-blank-template/azuredeploy.json" `
                    -WhatIf
            } catch {
                $validationResults.Messages += "Insufficient permissions on resource group: $resourceGroupName"
                $validationResults.Success = $false
            }
        }

        # Validate storage account name availability and permissions
        Write-Log "Validating storage account name and permissions..." -Level VALIDATION
        $storageAccount = Get-AzStorageAccount -ResourceGroupName $resourceGroupName `
            -Name $DefaultStorageAccountName -ErrorAction SilentlyContinue
        if (!$storageAccount) {
            $storageNameAvailable = Get-AzStorageAccountNameAvailability -Name $DefaultStorageAccountName
            if (!$storageNameAvailable.NameAvailable) {
                $validationResults.Messages += "Storage account name $DefaultStorageAccountName is not available"
                $validationResults.Success = $false
            }
        } else {
            try {
                $testAccess = Get-AzStorageAccountKey -ResourceGroupName $resourceGroupName `
                    -Name $DefaultStorageAccountName
            } catch {
                $validationResults.Messages += "Insufficient permissions on storage account: $DefaultStorageAccountName"
                $validationResults.Success = $false
            }
        }

        # Validate network configuration and conflicts
        Write-Log "Validating network configuration..." -Level VALIDATION
        $vnet = Get-AzVirtualNetwork -Name $vnetName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue
        if ($vnet) {
            if ($vnet.AddressSpace.AddressPrefixes -contains $DefaultVnetAddressSpace) {
                $validationResults.Messages += "Address space $DefaultVnetAddressSpace conflicts with existing VNet"
                $validationResults.Success = $false
            }
            
            # Validate subnet conflicts
            $existingSubnet = $vnet.Subnets | Where-Object { $_.Name -eq $subnetName }
            if ($existingSubnet -and $existingSubnet.AddressPrefix -ne $DefaultSubnetAddressSpace) {
                $validationResults.Messages += "Subnet $subnetName exists with different address space"
                $validationResults.Success = $false
            }
        }

        # Validate NSG
        Write-Log "Validating Network Security Group..." -Level VALIDATION
        $nsg = Get-AzNetworkSecurityGroup -Name $DefaultNsgName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue
        if ($nsg) {
            $rdpRule = $nsg.SecurityRules | Where-Object { $_.DestinationPortRange -eq '3389' }
            if (!$rdpRule) {
                $validationResults.Messages += "Existing NSG does not have required RDP rule"
                $validationResults.Success = $false
            }
        }

        # Validate Public IP availability
        Write-Log "Validating Public IP availability..." -Level VALIDATION
        $publicIp = Get-AzPublicIpAddress -Name $DefaultPublicIpName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue
        if ($publicIp) {
            if ($publicIp.IpConfiguration) {
                $validationResults.Messages += "Public IP $DefaultPublicIpName is already in use"
                $validationResults.Success = $false
            }
        }

        return $validationResults
    } catch {
        $validationResults.Success = $false
        $validationResults.Messages += "Validation error: $($_.Exception.Message)"
        return $validationResults
    }
}

try {
    # Phase 1: Validation
    Write-Log "Starting validation phase..." -Level VALIDATION
    $validation = Test-AzureResources

    if (!$validation.Success) {
        Write-Log "Validation failed:" -Level ERROR
        foreach ($message in $validation.Messages) {
            Write-Log $message -Level ERROR
        }
        throw "Resource validation failed. Please review the validation messages above."
    }

    Write-Log "All validations passed successfully." -Level VALIDATION

    if ($ValidateOnly) {
        Write-Log "Validation only mode - stopping before deployment." -Level INFO
        return
    }

    # Ask for confirmation before proceeding with deployment
    if (!$PSCmdlet.ShouldProcess("Azure Resources", "Deploy")) {
        Write-Log "Deployment cancelled by user." -Level INFO
        return
    }

    # Phase 2: Deployment
    Write-Log "Starting deployment phase..." -Level INFO

    # Create Resource Group if it doesn't exist
    $rg = Get-AzResourceGroup -Name $resourceGroupName -ErrorAction SilentlyContinue
    if (!$rg) {
        Write-Log "Creating resource group '$resourceGroupName'..."
        New-AzResourceGroup -Name $resourceGroupName -Location $location
    }

    # Create Storage Account if it doesn't exist
    $storageAccount = Get-AzStorageAccount -ResourceGroupName $resourceGroupName `
        -Name $DefaultStorageAccountName -ErrorAction SilentlyContinue
    if (!$storageAccount) {
        Write-Log "Creating Storage Account '$DefaultStorageAccountName'..." -Level INFO
        $storageAccountParams = @{
            ResourceGroupName = $resourceGroupName
            Name = $DefaultStorageAccountName
            Location = $location
            SkuName = 'Standard_LRS'
            Kind = 'StorageV2'
        }

        # Create storage account with timeout handling
        $timeout = New-TimeSpan -Minutes 5
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $storageAccount = New-AzStorageAccount @storageAccountParams

        # Wait for storage account to be ready
        Write-Log "Waiting for storage account provisioning to complete..." -Level INFO
        do {
            $storageAccount = Get-AzStorageAccount -ResourceGroupName $resourceGroupName `
                -Name $DefaultStorageAccountName -ErrorAction SilentlyContinue
            
            if ($stopwatch.Elapsed -gt $timeout) {
                Write-Log "Timeout waiting for storage account creation" -Level ERROR
                throw "Storage account creation timed out after 5 minutes"
            }

            if ($storageAccount.ProvisioningState -eq 'Failed') {
                Write-Log "Storage account provisioning failed" -Level ERROR
                throw "Storage account provisioning failed"
            }

            if ($storageAccount.ProvisioningState -ne 'Succeeded') {
                Write-Log "Storage account status: $($storageAccount.ProvisioningState)" -Level INFO
                Start-Sleep -Seconds 10
            }
        } while ($storageAccount.ProvisioningState -ne 'Succeeded')

        Write-Log "Storage account created successfully" -Level INFO
    } else {
        Write-Log "Using existing storage account '$DefaultStorageAccountName'" -Level INFO
    }

    # Create Network Security Group
    $nsg = Get-AzNetworkSecurityGroup -Name $DefaultNsgName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue
    if (!$nsg) {
        Write-Log "Creating Network Security Group '$DefaultNsgName'..."
        $nsgRules = @(
            @{
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
        )
        
        $nsg = New-AzNetworkSecurityGroup -ResourceGroupName $resourceGroupName `
            -Location $location `
            -Name $DefaultNsgName
        
        foreach ($rule in $nsgRules) {
            Add-AzNetworkSecurityRuleConfig @rule -NetworkSecurityGroup $nsg
        }
        $nsg | Set-AzNetworkSecurityGroup
    }

    # Create Virtual Network and Subnet
    $vnet = Get-AzVirtualNetwork -Name $vnetName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue
    if (!$vnet) {
        Write-Log "Creating Virtual Network '$vnetName'..."
        $subnetConfig = New-AzVirtualNetworkSubnetConfig -Name $subnetName `
            -AddressPrefix $DefaultSubnetAddressSpace `
            -NetworkSecurityGroup $nsg

        $vnet = New-AzVirtualNetwork -ResourceGroupName $resourceGroupName `
            -Location $location `
            -Name $vnetName `
            -AddressPrefix $DefaultVnetAddressSpace `
            -Subnet $subnetConfig
    }

    # Create Public IP
    $publicIp = Get-AzPublicIpAddress -Name $DefaultPublicIpName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue
    if (!$publicIp) {
        Write-Log "Creating Public IP '$DefaultPublicIpName'..."
        $publicIp = New-AzPublicIpAddress -ResourceGroupName $resourceGroupName `
            -Location $location `
            -Name $DefaultPublicIpName `
            -Sku Standard `
            -AllocationMethod Static
    }

    # Create NIC
    $nicName = "$vmName-NIC"
    $subnet = Get-AzVirtualNetwork -Name $vnetName -ResourceGroupName $resourceGroupName |
        Get-AzVirtualNetworkSubnetConfig -Name $subnetName
    $nic = New-AzNetworkInterface -Name $nicName -ResourceGroupName $resourceGroupName `
        -Location $location -SubnetId $subnet.Id -PublicIpAddressId $publicIp.Id

    # Create VM Configuration
    Write-Log "Creating VM configuration..." -Level INFO
    $vmConfig = New-AzVMConfig -VMName $vmName -VMSize $VMSize
    $vmConfig = Set-AzVMOperatingSystem -VM $vmConfig -Windows -ComputerName $vmName `
        -Credential (New-Object System.Management.Automation.PSCredential ($adminUsername, (ConvertTo-SecureString $adminPassword -AsPlainText -Force)))
    $vmConfig = Add-AzVMNetworkInterface -VM $vmConfig -Id $nic.Id
    $vmConfig = Set-AzVMSourceImage -VM $vmConfig -PublisherName 'MicrosoftWindowsServer' `
        -Offer 'WindowsServer' -Skus '2022-Datacenter' -Version latest
    $vmConfig = Set-AzVMBootDiagnostic -VM $vmConfig -Enable -StorageAccountName $DefaultStorageAccountName

    # Enable Trusted Launch
    $vmConfig.SecurityProfile = New-Object Microsoft.Azure.Management.Compute.Models.SecurityProfile
    $vmConfig.SecurityProfile.SecurityType = "TrustedLaunch"
    $vmConfig.SecurityProfile.UefiSettings = New-Object Microsoft.Azure.Management.Compute.Models.UefiSettings
    $vmConfig.SecurityProfile.UefiSettings.SecureBootEnabled = $true
    $vmConfig.SecurityProfile.UefiSettings.VTpmEnabled = $true

    # Create the VM
    Write-Log "Creating VM '$vmName'..." -Level INFO
    New-AzVM -ResourceGroupName $resourceGroupName -Location $location -VM $vmConfig

    # Configure auto-shutdown
    Write-Log "Configuring auto-shutdown schedule for VM '$vmName'..." -Level INFO
    $shutdownTime = "21:00" # 9:00 PM
    $timeZone = "UTC-07:00" # UTC-7
    
    $scheduledShutdownResourceId = "/subscriptions/{0}/resourceGroups/{1}/providers/microsoft.devtestlab/schedules/shutdown-computevm-{2}" -f `
        (Get-AzContext).Subscription.Id, $resourceGroupName, $vmName

    $properties = @{
        status = "Enabled"
        taskType = "ComputeVmShutdownTask"
        dailyRecurrence = @{time = $shutdownTime }
        timeZoneId = $timeZone
        targetResourceId = (Get-AzVM -ResourceGroupName $resourceGroupName -Name $vmName).Id
        notificationSettings = @{
            status = "Disabled"
        }
    }

    New-AzResource -ResourceId $scheduledShutdownResourceId `
        -Properties $properties `
        -Force

    # Validate DevTest Lab shutdown schedule
    Write-Log "Validating DevTest Lab shutdown schedule..." -Level VALIDATION
    $existingSchedule = Get-AzResource -ResourceId $scheduledShutdownResourceId -ErrorAction SilentlyContinue
    if (!$existingSchedule) {
        Write-Log "DevTest Lab shutdown schedule validation failed - schedule not found" -Level ERROR
        throw "Failed to create DevTest Lab shutdown schedule for VM $vmName"
    }
    Write-Log "DevTest Lab shutdown schedule validated successfully" -Level INFO
    Write-Log "Auto-shutdown schedule configured successfully for $vmName to shutdown at $shutdownTime $timeZone" -Level INFO

    Write-Log "Domain Controller VM creation completed successfully." -Level INFO
} catch {
    Write-Log $_.Exception.Message -Level ERROR
    throw
} finally {
    Write-Log "Script execution completed." -Level INFO
}