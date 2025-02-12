# =============================================================================
# Script: DC-Validation.ps1
# Created: 2025-02-12 00:15:40 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-12 00:15:40 UTC
# Updated By: jdyer-nuvodia
# Version: 1.3
# Additional Info: Added explicit Trusted Launch VM size detection
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

        # Get available VM sizes with Trusted Launch support
        Write-Log "Checking available VM sizes with Trusted Launch support..." -Level VALIDATION
        $vmSizes = Get-AzVMSize -Location $Config.Location | Where-Object {
            try {
                $vmSize = $_
                $output = az vm list-skus --location $Config.Location --size $vmSize.Name --query "[?name=='$($vmSize.Name)']"
                $output | ConvertFrom-Json | Where-Object { $_.capabilities.Name -contains "TrustedLaunchEnabled" -and $_.capabilities.Value -contains "True" }
            } catch {
                $false
            }
        }

        if ($vmSizes.Count -eq 0) {
            $validationResults.Messages += "No VM sizes supporting Trusted Launch found in $($Config.Location)"
            $validationResults.Success = $false
        } else {
            Write-Log "Found $(($vmSizes | Select-Object -First 5 | ForEach-Object { $_.Name }) -join ', ') supporting Trusted Launch" -Level INFO
            
            # Validate current VM size
            if ($Config.VMSize -notin $vmSizes.Name) {
                $recommendedSize = $vmSizes | Where-Object { $_.NumberOfCores -ge 2 -and $_.MemoryInMB -ge 4096 } | Select-Object -First 1
                $validationResults.Messages += "VM size $($Config.VMSize) does not support Trusted Launch. Consider using $($recommendedSize.Name)"
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