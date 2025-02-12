# =============================================================================
# Script: DC-Configuration.ps1
# Created: 2025-02-12 00:09:49 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-12 00:09:49 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2
# Additional Info: Updated VM size to Standard_D8_v5 for confirmed Trusted Launch support
# =============================================================================

function Write-Log {
    param(
        [string]$Message,
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
        VMSize             = 'Standard_D8_v5'  # Updated to confirmed Trusted Launch compatible size
        ShutdownTime       = '21:00'
        TimeZone           = 'UTC-07:00'
        ImagePublisher     = 'MicrosoftWindowsServer'
        ImageOffer         = 'WindowsServer'
        ImageSku           = '2022-datacenter-g2'
        ImageVersion       = 'latest'
    }
    return $config
}