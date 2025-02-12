# =============================================================================
# Script: DC-Validation.psm1
# Created: 2025-02-12 00:25:18 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-12 15:03:43 UTC
# Updated By: jdyer-nuvodia
# Version: 1.5
# Additional Info: Converted to PowerShell module format
# =============================================================================

function Test-DCPrerequisites {
    [CmdletBinding()]
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

        # Rest of the function remains the same...
        # [Previous validation code remains unchanged]

        return $validationResults
    } catch {
        $validationResults.Success = $false
        $validationResults.Messages += "Validation error: $($_.Exception.Message)"
        return $validationResults
    }
}

# Export functions
Export-ModuleMember -Function Test-DCPrerequisites