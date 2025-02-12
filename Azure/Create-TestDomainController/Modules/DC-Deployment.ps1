# =============================================================================
# Script: DC-Deployment.ps1
# Created: 2025-02-12 00:39:44 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-12 00:39:44 UTC
# Updated By: jdyer-nuvodia
# Version: 3.2
# Additional Info: Enhanced timezone validation and conversion
# =============================================================================

function Set-VMAutoShutdown {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ResourceGroupName,
        [Parameter(Mandatory = $true)]
        [string]$VMName,
        [Parameter(Mandatory = $true)]
        [string]$Location,
        [Parameter(Mandatory = $false)]
        [string]$TimeZoneId = 'US Mountain Standard Time',
        [Parameter(Mandatory = $false)]
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