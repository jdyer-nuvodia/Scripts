# =============================================================================
# Script: DC-Configuration.psm1
# Created: 2025-02-12 00:25:18 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-12 15:03:43 UTC
# Updated By: jdyer-nuvodia
# Version: 1.5
# Additional Info: Converted to PowerShell module format
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
    return $config
}

# Export functions
Export-ModuleMember -Function Write-Log, Initialize-DCConfiguration