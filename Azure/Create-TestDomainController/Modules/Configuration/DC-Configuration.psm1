# =============================================================================
# Script: DC-Configuration.psm1
# Created: 2025-02-12 00:25:18 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-12 17:32:27 UTC
# Updated By: jdyer-nuvodia
# Version: 1.6
# Additional Info: Enhanced module structure and error handling
# =============================================================================

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter()]
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'VALIDATION')]
        [string]$Level = 'INFO'
    )
    
    $LogMessage = "[$([DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))] [$Level] $Message"
    Add-Content -Path $LogFile -Value $LogMessage
    Write-Host $LogMessage
    if ($Level -eq 'ERROR') {
        Write-Error $Message
    } elseif ($VerbosePreference -eq 'Continue') {
        Write-Verbose $Message
    }
}

function Initialize-DCConfiguration {
    [CmdletBinding()]
    param()
    
    try {
        $config = @{
            ResourceGroupName     = 'JB-TEST-RG2'
            Location             = 'westus2'
            StorageAccountName   = 'jbteststorage0'
            VnetName            = 'JB-TEST-VNET'
            SubnetName          = 'JB-TEST-SUBNET1'
            VmName              = 'JB-TEST-DC01'
            AdminUsername       = 'jbadmin'
            AdminPassword       = 'TS-pGxB~8m^A~WH^[yB8'
            DomainName         = 'JB-TEST.local'
            PublicIpName       = 'JB-TEST-DC01-PUBIP'
            NsgName            = 'JB-TEST-NSG'
            VnetAddressSpace   = '10.0.0.0/16'
            SubnetAddressSpace = '10.0.1.0/24'
            VMSize             = 'Standard_D2s_v4'
            ShutdownTime       = '21:00'
            TimeZone           = 'UTC-07:00'
            ImagePublisher     = 'MicrosoftWindowsServer'
            ImageOffer        = 'WindowsServer'
            ImageSku          = '2022-datacenter-g2'
            ImageVersion      = 'latest'
        }
        Write-Log "Successfully initialized DC configuration" -Level INFO
        return $config
    } catch {
        Write-Log "Failed to initialize DC configuration: $_" -Level ERROR
        throw
    }
}

# Export functions
Export-ModuleMember -Function Write-Log, Initialize-DCConfiguration