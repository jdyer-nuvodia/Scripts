# Script: Delete-TestDomainController.ps1
# Version: 1.0
# Description: Deletes Azure resources created by Create-TestDomainController.ps1
# Author: jdyer-nuvodia
# Last Modified: 2025-02-05 22:42:09
#
# .SYNOPSIS
#   Confirms and deletes Azure resources created for test domain controller
#
# .DESCRIPTION
#   This script safely removes Azure resources that were created by Create-TestDomainController.ps1.
#   It includes confirmation prompts for each resource and deletes them in the correct order
#   to prevent dependency conflicts.
#
# .PARAMETER ResourceGroupName
#   The name of the resource group containing the test domain controller resources
#
# .PARAMETER Confirm
#   If specified, prompts for confirmation before deleting each resource
#
# .EXAMPLE
#   # Delete resources with confirmation for each
#   .\Delete-TestDomainController.ps1 -ResourceGroupName "rg-testdc" -Confirm
#
# .EXAMPLE
#   # Delete all resources with single confirmation
#   .\Delete-TestDomainController.ps1 -ResourceGroupName "rg-testdc"
#
# .NOTES
#   - Requires Azure PowerShell module (Az)
#   - Requires appropriate Azure permissions
#   - Will delete: VM, NIC, Public IP, NSG, VNet, Storage Account, Resource Group
#

[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='High')]
param(
    [Parameter(Mandatory=$true)]
    [string]$ResourceGroupName,
    
    [Parameter(Mandatory=$false)]
    [switch]$Confirm
)

# Function to prompt for confirmation
function Get-UserConfirmation {
    param(
        [string]$Resource
    )
    
    if ($Confirm) {
        $choice = Read-Host "Are you sure you want to delete $Resource? (y/n)"
        return $choice -eq 'y'
    }
    return $true
}

try {
    # Check if Azure PowerShell module is installed
    if (!(Get-Module -ListAvailable -Name Az)) {
        throw "Azure PowerShell module is not installed. Please install it using: Install-Module -Name Az"
    }

    # Check if resource group exists
    $resourceGroup = Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction SilentlyContinue
    if (!$resourceGroup) {
        throw "Resource group '$ResourceGroupName' not found."
    }

    Write-Host "`nPreparing to delete resources in resource group: $ResourceGroupName" -ForegroundColor Yellow
    Write-Host "This will delete ALL resources created by Create-TestDomainController.ps1" -ForegroundColor Red
    Write-Host "Including: VM, NIC, Public IP, NSG, VNet, Storage Account, and Resource Group`n" -ForegroundColor Red

    if ($PSCmdlet.ShouldProcess($ResourceGroupName, "Delete all resources")) {
        # Get VM details
        $vm = Get-AzVM -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
        if ($vm) {
            Write-Host "Found VM: $($vm.Name)" -ForegroundColor Cyan
            
            # Stop VM if running
            $vmStatus = Get-AzVM -ResourceGroupName $ResourceGroupName -Name $vm.Name -Status
            if ($vmStatus.Statuses.Code -contains "PowerState/running") {
                Write-Host "Stopping VM..." -ForegroundColor Yellow
                Stop-AzVM -ResourceGroupName $ResourceGroupName -Name $vm.Name -Force
            }

            # Delete VM
            if (Get-UserConfirmation -Resource "VM: $($vm.Name)") {
                Write-Host "Deleting VM..." -ForegroundColor Yellow
                Remove-AzVM -ResourceGroupName $ResourceGroupName -Name $vm.Name -Force
            }

            # Get and delete NIC
            $nic = Get-AzNetworkInterface -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
            if ($nic -and (Get-UserConfirmation -Resource "Network Interface: $($nic.Name)")) {
                Write-Host "Deleting Network Interface..." -ForegroundColor Yellow
                Remove-AzNetworkInterface -ResourceGroupName $ResourceGroupName -Name $nic.Name -Force
            }

            # Get and delete Public IP
            $pip = Get-AzPublicIpAddress -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
            if ($pip -and (Get-UserConfirmation -Resource "Public IP: $($pip.Name)")) {
                Write-Host "Deleting Public IP..." -ForegroundColor Yellow
                Remove-AzPublicIpAddress -ResourceGroupName $ResourceGroupName -Name $pip.Name -Force
            }
        }

        # Get and delete NSG
        $nsg = Get-AzNetworkSecurityGroup -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
        if ($nsg -and (Get-UserConfirmation -Resource "Network Security Group: $($nsg.Name)")) {
            Write-Host "Deleting Network Security Group..." -ForegroundColor Yellow
            Remove-AzNetworkSecurityGroup -ResourceGroupName $ResourceGroupName -Name $nsg.Name -Force
        }

        # Get and delete VNet
        $vnet = Get-AzVirtualNetwork -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
        if ($vnet -and (Get-UserConfirmation -Resource "Virtual Network: $($vnet.Name)")) {
            Write-Host "Deleting Virtual Network..." -ForegroundColor Yellow
            Remove-AzVirtualNetwork -ResourceGroupName $ResourceGroupName -Name $vnet.Name -Force
        }

        # Get and delete Storage Account
        $storageAccounts = Get-AzStorageAccount -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
        foreach ($sa in $storageAccounts) {
            if (Get-UserConfirmation -Resource "Storage Account: $($sa.StorageAccountName)") {
                Write-Host "Deleting Storage Account: $($sa.StorageAccountName)..." -ForegroundColor Yellow
                Remove-AzStorageAccount -ResourceGroupName $ResourceGroupName -Name $sa.StorageAccountName -Force
            }
        }

        # Finally, delete the Resource Group
        if (Get-UserConfirmation -Resource "Resource Group: $ResourceGroupName") {
            Write-Host "Deleting Resource Group..." -ForegroundColor Yellow
            Remove-AzResourceGroup -Name $ResourceGroupName -Force
        }

        Write-Host "`nResource deletion completed successfully." -ForegroundColor Green
    }
}
catch {
    Write-Error "Error occurred during resource deletion: $_"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    exit 1
}