# =============================================================================
# Script: DC-Deployment.psm1
# Created: 2025-02-12 00:25:18 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-12 18:58:20 UTC
# Updated By: jdyer-nuvodia
# Version: 1.8
# Additional Info: Added logging functions and proper exports
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
        $LogFile = Join-Path -Path $PSScriptRoot -ChildPath "DC-Deployment.log"
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

# Original deployment functions here...

# Export functions
Export-ModuleMember -Function Write-Log, Set-DCLogFile, New-DCEnvironment