# =============================================================================
# Script: DC-Validation.ps1
# Created: 2025-02-12 00:25:18 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-12 00:25:18 UTC
# Updated By: jdyer-nuvodia
# Version: 1.4
# Additional Info: Optimized Trusted Launch validation for better performance
# =============================================================================

function Test-DCPrerequisites {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Config
    )
    $validationResults = @{
        Success = $true
        Messages = @()
    }
    try {
        # Validate required modules and versions
        Write-Log "Validating required modules..." -Level VALIDATION
        $requiredModules = @{
            'Az.Accounts'  = '2.12.1'
            'Az.Resources' = '6.6.0'
            'Az.Network'   = '5.0.0'
            'Az.Storage'   = '5.4.0'
            'Az.Compute'   = '5.7.0'
        }
        foreach ($module in $requiredModules.GetEnumerator()) {
            $installedModule = Get-Module -Name $module.Key -ListAvailable
            if (!$installedModule) {
                $validationResults.Messages += "Required module $($module.Key) is not installed"
                $validationResults.Success = $false
            } else {
                $latestVersion = $installedModule | Sort-Object Version -Descending | Select-Object -First 1
                if ($latestVersion.Version -lt [Version]$module.Value) {
                    $validationResults.Messages += "Module $($module.Key) version $($latestVersion.Version) is below required version $($module.Value)"
                    $validationResults.Success = $false
                }
            }
        }

        # Validate location
        Write-Log "Validating location '$($Config.Location)'..." -Level VALIDATION
        $validLocations = Get-AzLocation
        if ($Config.Location -notin $validLocations.Location) {
            $validationResults.Messages += "Invalid location: $($Config.Location)"
            $validationResults.Success = $false
        }

        # Validate VM size for Gen2 and minimum requirements
        Write-Log "Validating VM size compatibility..." -Level VALIDATION
        $vmSize = Get-AzVMSize -Location $Config.Location | Where-Object { $_.Name -eq $Config.VMSize }
        if (!$vmSize) {
            $validationResults.Messages += "VM size $($Config.VMSize) is not available in $($Config.Location)"
            $validationResults.Success = $false
        } else {
            # Check minimum requirements for Domain Controller
            if ($vmSize.NumberOfCores -lt 2 -or $vmSize.MemoryInMB -lt 4096) {
                $validationResults.Messages += "VM size $($Config.VMSize) does not meet minimum requirements (2 cores, 4GB RAM)"
                $validationResults.Success = $false
            }
            
            # For Trusted Launch, we'll use known compatible v4/v5 series sizes
            $trustedLaunchSeries = @(
                'Standard_D2s_v4', 'Standard_D4s_v4', 'Standard_D8s_v4',
                'Standard_D2s_v5', 'Standard_D4s_v5', 'Standard_D8s_v5',
                'Standard_E2s_v4', 'Standard_E4s_v4', 'Standard_E8s_v4',
                'Standard_E2s_v5', 'Standard_E4s_v5', 'Standard_E8s_v5'
            )
            
            if ($Config.VMSize -notin $trustedLaunchSeries) {
                $recommendedSize = $trustedLaunchSeries | Where-Object { 
                    $size = Get-AzVMSize -Location $Config.Location | Where-Object { $_.Name -eq $_ }
                    $size -and $size.NumberOfCores -ge 2 -and $size.MemoryInMB -ge 4096
                } | Select-Object -First 1
                
                $validationResults.Messages += "VM size $($Config.VMSize) may not support Trusted Launch. Recommended size: $recommendedSize"
                $validationResults.Success = $false
            }
        }

        # Validate resource name availability
        Write-Log "Validating resource names..." -Level VALIDATION
        $storageNameAvailable = Get-AzStorageAccountNameAvailability -Name $Config.StorageAccountName
        if (-not $storageNameAvailable.NameAvailable) {
            if (-not (Get-AzStorageAccount -ResourceGroupName $Config.ResourceGroupName -Name $Config.StorageAccountName -ErrorAction SilentlyContinue)) {
                $validationResults.Messages += "Storage account name $($Config.StorageAccountName) is not available"
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