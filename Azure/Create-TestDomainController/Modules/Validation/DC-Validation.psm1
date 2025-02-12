# =============================================================================
# Script: DC-Validation.psm1
# Created: 2025-02-12 00:25:18 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-12 20:24:42 UTC
# Updated By: jdyer-nuvodia
# Version: 1.9
# Additional Info: Added Test-DCPrerequisites implementation and enhanced validation
# =============================================================================

# Script-scoped variables
$Script:LogFile = $null

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter()]
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'VALIDATION')]
        [string]$Level = 'INFO',
        [Parameter()]
        [string]$LogFile = $Script:LogFile
    )
    if ([string]::IsNullOrEmpty($LogFile)) {
        $LogFile = Join-Path -Path $PSScriptRoot -ChildPath "DC-Validation.log"
    }
    $LogMessage = "[$([DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))] [$Level] $Message"
    Add-Content -Path $LogFile -Value $LogMessage
    Write-Host $LogMessage
    if ($Level -eq 'ERROR') {
        Write-Error $Message
    } elseif ($VerbosePreference -eq 'Continue') {
        Write-Verbose $Message
    }
}

function Set-DCLogFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    $Script:LogFile = $Path
    Write-Log "Log file path set to: $Path" -Level INFO
}

function Test-DCPrerequisites {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Config
    )
    
    try {
        Write-Log "Starting prerequisite validation..." -Level VALIDATION
        $validationResults = @{
            Success = $true
            Messages = @()
        }

        # Validate Resource Group Name
        if ([string]::IsNullOrWhiteSpace($Config.ResourceGroupName)) {
            $validationResults.Success = $false
            $validationResults.Messages += "Resource group name cannot be empty"
        }

        # Validate Location
        $validLocations = @('westus2', 'eastus', 'eastus2', 'centralus')
        if (-not ($validLocations -contains $Config.Location)) {
            $validationResults.Success = $false
            $validationResults.Messages += "Invalid location. Must be one of: $($validLocations -join ', ')"
        }

        # Validate VM Size
        if ([string]::IsNullOrWhiteSpace($Config.VMSize)) {
            $validationResults.Success = $false
            $validationResults.Messages += "VM size cannot be empty"
        }

        # Validate Network Configuration
        if ([string]::IsNullOrWhiteSpace($Config.VnetName) -or 
            [string]::IsNullOrWhiteSpace($Config.SubnetName)) {
            $validationResults.Success = $false
            $validationResults.Messages += "Virtual network and subnet names are required"
        }

        # Validate Admin Credentials
        if ([string]::IsNullOrWhiteSpace($Config.AdminUsername) -or 
            [string]::IsNullOrWhiteSpace($Config.AdminPassword)) {
            $validationResults.Success = $false
            $validationResults.Messages += "Admin username and password are required"
        }

        # Validate Password Complexity
        if ($Config.AdminPassword -and $Config.AdminPassword.Length -lt 12) {
            $validationResults.Success = $false
            $validationResults.Messages += "Admin password must be at least 12 characters long"
        }

        # Check Azure Connection
        try {
            $context = Get-AzContext -ErrorAction Stop
            if (-not $context) {
                $validationResults.Success = $false
                $validationResults.Messages += "Not connected to Azure. Please run Connect-AzAccount first"
            }
        }
        catch {
            $validationResults.Success = $false
            $validationResults.Messages += "Failed to verify Azure connection: $_"
        }

        # Log validation results
        if ($validationResults.Success) {
            Write-Log "All prerequisite validations passed successfully" -Level VALIDATION
        } else {
            foreach ($message in $validationResults.Messages) {
                Write-Log $message -Level ERROR
            }
        }

        return $validationResults
    }
    catch {
        Write-Log "Validation failed with unexpected error: $_" -Level ERROR
        return @{
            Success = $false
            Messages = @("Unexpected error during validation: $_")
        }
    }
}

# Export functions
Export-ModuleMember -Function Write-Log, Set-DCLogFile, Test-DCPrerequisites