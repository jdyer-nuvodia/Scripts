# =============================================================================
# Script: DC-Deployment.ps1
# Created: 2025-02-12 00:39:44 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-12 00:52:33 UTC
# Updated By: jdyer-nuvodia
# Version: 3.3
# Additional Info: Added New-DCEnvironment function for VM deployment
# =============================================================================

function New-DCEnvironment {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Config
    )

    try {
        Write-Log "Starting Domain Controller environment deployment..." -Level INFO

        # Create Resource Group if it doesn't exist
        if (-not (Get-AzResourceGroup -Name $Config.ResourceGroupName -ErrorAction SilentlyContinue)) {
            Write-Log "Creating Resource Group '$($Config.ResourceGroupName)'..." -Level INFO
            New-AzResourceGroup -Name $Config.ResourceGroupName -Location $Config.Location
        }

        # Create VM with Trusted Launch configuration
        $vmConfig = New-AzVMConfig -VMName $Config.VmName -VMSize $Config.VMSize -SecurityType "TrustedLaunch"
        
        # Configure OS disk
        $vmConfig = Set-AzVMOperatingSystem -VM $vmConfig `
            -Windows `
            -ComputerName $Config.VmName `
            -Credential (New-Object PSCredential ($Config.AdminUsername, (ConvertTo-SecureString $Config.AdminPassword -AsPlainText -Force)))

        # Add network interface
        $nicName = "$($Config.VmName)-nic"
        $subnet = Get-AzVirtualNetworkSubnetConfig -Name $Config.SubnetName -VirtualNetwork (Get-AzVirtualNetwork -Name $Config.VnetName -ResourceGroupName $Config.ResourceGroupName)
        $nic = New-AzNetworkInterface -Name $nicName `
            -ResourceGroupName $Config.ResourceGroupName `
            -Location $Config.Location `
            -SubnetId $subnet.Id

        $vmConfig = Add-AzVMNetworkInterface -VM $vmConfig -Id $nic.Id

        # Configure OS image
        $vmConfig = Set-AzVMSourceImage -VM $vmConfig `
            -PublisherName 'MicrosoftWindowsServer' `
            -Offer 'WindowsServer' `
            -Skus '2022-Datacenter' `
            -Version 'latest'

        # Create the VM
        Write-Log "Creating Virtual Machine '$($Config.VmName)'..." -Level INFO
        New-AzVM -ResourceGroupName $Config.ResourceGroupName `
            -Location $Config.Location `
            -VM $vmConfig

        # Configure auto-shutdown
        Write-Log "Configuring auto-shutdown schedule..." -Level INFO
        Set-VMAutoShutdown -ResourceGroupName $Config.ResourceGroupName `
            -VMName $Config.VmName `
            -Location $Config.Location `
            -TimeZoneId $Config.TimeZoneId

        Write-Log "Domain Controller environment deployment completed successfully." -Level INFO
    }
    catch {
        Write-Log "Failed to deploy Domain Controller environment: $_" -Level ERROR
        throw
    }
}

function Set-VMAutoShutdown {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ResourceGroupName,
        [Parameter(Mandatory = $true)]
        [string]$VMName,
        [Parameter(Mandatory = $true)]
        [string]$Location,
        [Parameter(Mandatory = false)]
        [string]$TimeZoneId = 'US Mountain Standard Time',
        [Parameter(Mandatory = false)]
        [string]$ShutdownTime = "2200"
    )
    
    try {
        # Validate timezone
        if ($TimeZoneId -ne 'US Mountain Standard Time') {
            Write-Log "WARNING: Using non-Arizona timezone. Recommended to use 'US Mountain Standard Time' for Arizona operations." -Level WARNING
        }

        # Validate shutdown time format and conversion
        $timeValidation = Test-ArizonaTimeValidation -ShutdownTime $ShutdownTime
        if (-not $timeValidation.IsValid) {
            throw $timeValidation.Message
        }

        Write-Log "Configuring auto-shutdown schedule:" -Level INFO
        Write-Log "  - Arizona Time: $($timeValidation.ArizonaTime.ToString('HH:mm'))" -Level INFO
        Write-Log "  - UTC Time: $($timeValidation.UtcTime.ToString('HH:mm'))" -Level INFO
        Write-Log "  - Timezone: $TimeZoneId" -Level INFO

        $properties = @{
            status             = "Enabled"
            taskType          = "ComputeVmShutdownTask"
            dailyRecurrence   = @{ time = $ShutdownTime }
            timeZoneId        = $TimeZoneId
            notificationSettings = @{
                status = "Disabled"
            }
            targetResourceId  = "/subscriptions/$((Get-AzContext).Subscription.Id)/resourceGroups/$ResourceGroupName/providers/Microsoft.Compute/virtualMachines/$VMName"
        }

        $scheduledShutdownResourceId = "/subscriptions/$((Get-AzContext).Subscription.Id)/resourceGroups/$ResourceGroupName/providers/microsoft.devtestlab/schedules/shutdown-computevm-$VMName"

        New-AzResource -ResourceId $scheduledShutdownResourceId `
                      -Properties $properties `
                      -Location $Location `
                      -ApiVersion "2018-10-15-preview" `
                      -Force

        Write-Log "Auto-shutdown schedule configured successfully." -Level INFO
    }
    catch {
        Write-Log "Failed to configure auto-shutdown schedule: $_" -Level ERROR
        throw
    }
}