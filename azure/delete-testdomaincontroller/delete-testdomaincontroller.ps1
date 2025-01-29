# Variables (ensure these match the ones used in the setup script)
$resourceGroup = "JB-TEST-RG"
$vmName = "JB-TEST-DC01"
$bastionName = "$vmName-bastion"
$nsgName = "$vmName" + "NSG"
$nicName = "$vmName" + "VMNic"
$publicIpName = "$vmName" + "PublicIP"
$vnetName = "JB-TEST-VNET"

# Function to prompt for deletion
function Confirm-Deletion {
    param (
        [string]$Resource
    )
    $choice = Read-Host -Prompt "Are you sure you want to delete $Resource? (y/n)"
    if ($choice -eq 'y' -or $choice -eq 'Y') {
        return $true
    } else {
        Write-Host "Skipping $Resource deletion"
        return $false
    }
}

# Delete specific resources created by the setup script

# Delete the VM and associated resources
if (Confirm-Deletion -Resource "VM and associated resources ($vmName)") {
    Write-Host "Deleting VM and associated resources..."
    Remove-AzVM -ResourceGroupName $resourceGroup -Name $vmName -Force
}

# Delete any OS disk that contains $vmName
$osDisks = Get-AzDisk -ResourceGroupName $resourceGroup | Where-Object { $_.Name -like "*$vmName*" }
foreach ($osDisk in $osDisks) {
    if (Confirm-Deletion -Resource "OS disk ($($osDisk.Name))") {
        Write-Host "Deleting OS disk $($osDisk.Name)..."
        Remove-AzDisk -ResourceGroupName $resourceGroup -DiskName $osDisk.Name -Force
    }
}

# Delete the VM's network interface
if (Confirm-Deletion -Resource "network interface ($nicName)") {
    Write-Host "Deleting network interface $nicName..."
    Remove-AzNetworkInterface -ResourceGroupName $resourceGroup -Name $nicName -Force
}

# Delete the VM's public IP address
if (Confirm-Deletion -Resource "public IP address ($publicIpName)") {
    Write-Host "Deleting public IP address $publicIpName..."
    Remove-AzPublicIpAddress -ResourceGroupName $resourceGroup -Name $publicIpName -Force
}

# Delete the network security group
if (Confirm-Deletion -Resource "network security group ($nsgName)") {
    Write-Host "Deleting network security group $nsgName..."
    Remove-AzNetworkSecurityGroup -ResourceGroupName $resourceGroup -Name $nsgName -Force
}

# Delete the bastion host
if (Confirm-Deletion -Resource "bastion host ($bastionName)") {
    Write-Host "Deleting bastion host $bastionName..."
    Remove-AzBastion -ResourceGroupName $resourceGroup -Name $bastionName -Force
}

# Delete the virtual network
if (Confirm-Deletion -Resource "virtual network ($vnetName)") {
    Write-Host "Deleting virtual network $vnetName..."
    Remove-AzVirtualNetwork -ResourceGroupName $resourceGroup -Name $vnetName -Force
}

# Option to delete all remaining resources in the resource group
$remainingResources = Get-AzResource -ResourceGroupName $resourceGroup
if ($remainingResources.Count -gt 0) {
    Write-Host "`nThe following resources remain in the resource group:"
    foreach ($resource in $remainingResources) {
        Write-Host "- Name: $($resource.Name), Type: $($resource.ResourceType)"
    }

    if (Confirm-Deletion -Resource "all remaining resources in resource group '$resourceGroup'") {
        foreach ($resource in $remainingResources) {
            Write-Host "Deleting resource: $($resource.Name) of type: $($resource.ResourceType)..."
            Remove-AzResource -ResourceId $resource.ResourceId -Force
        }
        Write-Host "All remaining resources deleted."
    } else {
        Write-Host "Skipped deleting remaining resources."
    }
} else {
    Write-Host "`nNo remaining resources found in resource group '$resourceGroup'."
}

# Option to delete the entire resource group
if (Confirm-Deletion -Resource "the entire resource group '$resourceGroup'") {
    Write-Host "Deleting resource group: $resourceGroup..."
    Remove-AzResourceGroup -Name $resourceGroup -Force
} else {
    Write-Host "Skipped deleting resource group '$resourceGroup'."
}

Write-Host "`nCleanup process complete."
