# =============================================================================
# Script: Find-ADAccountBySID.ps1
# Created: 2025-03-11 15:30:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-11 15:30:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.0
# Additional Info: Initial script creation for AD SID resolution
# =============================================================================

<#
.SYNOPSIS
    Resolve Active Directory account from SID
.DESCRIPTION
    Checks AD module availability, validates SID format, and queries Active Directory
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$SID,
    
    [string]$DomainController
)

# Check for Active Directory module
if (-not (Get-Module -ListAvailable ActiveDirectory)) {
    Write-Host "ERROR: Active Directory module not found!" -ForegroundColor Red
    Write-Host "Install RSAT Tools first:" -ForegroundColor Yellow
    Write-Host "1. Press Win+X -> Settings -> Apps -> Optional Features"
    Write-Host "2. Add Feature -> RSAT: Active Directory and Lightweight Directory Services"
    Write-Host "OR Run as Admin: Install-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
    exit
}

Import-Module ActiveDirectory -ErrorAction Stop

# Validate SID format
if (-not ($SID -match '^S-\d-\d+(-\d+)+$')) {
    Write-Host "Invalid SID format! Use S-X-X-XX-XXXXXXXXXX-..." -ForegroundColor Red
    exit
}

try {
    $params = @{
        Filter     = {ObjectSID -eq $SID}
        Properties = 'Name', 'SamAccountName', 'ObjectClass', 'DistinguishedName'
        ErrorAction = 'Stop'
    }

    if ($DomainController) {
        $params['Server'] = $DomainController
    }

    $result = Get-ADObject @params

    Write-Host "`n[ MATCH FOUND ]`n" -ForegroundColor Green
    $result | Format-List Name, SamAccountName, ObjectClass, 
        DistinguishedName, ObjectSID
}
catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
    Write-Host "`n[ NO ACCOUNT FOUND ]`n" -ForegroundColor Red
    Write-Host "No AD object found with SID: $SID" -ForegroundColor Yellow
}
catch {
    Write-Host "`n[ ERROR ]`n" -ForegroundColor Red
    Write-Host "Failed to query Active Directory: $($_.Exception.Message)" -ForegroundColor Yellow
}
