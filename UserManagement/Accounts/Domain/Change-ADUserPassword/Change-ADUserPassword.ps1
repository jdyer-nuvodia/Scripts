# =============================================================================
# Script: Change-ADUserPassword.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2.1
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
# =============================================================================

<#
.SYNOPSIS
    Changes the password for an Active Directory user account.
.DESCRIPTION
    This script resets the password for a specified Active Directory user account.
    - Requires Active Directory PowerShell module
    - Must be run with appropriate AD permissions
    - Handles errors during password reset process
.PARAMETER Username
    The SAM account name of the AD user whose password needs to be changed
.PARAMETER NewPassword
    The new password to set for the user account (SecureString)
.EXAMPLE
    $SecurePass = ConvertTo-SecureString "NewP@ssw0rd123!" -AsPlainText -Force
    .\Change-ADUserPassword.ps1 -Username "jsmith" -NewPassword $SecurePass
    Changes password for user jsmith to the specified password
.NOTES
    Security Level: High
    Required Permissions: Domain Admin or delegated AD password reset rights
    Validation Requirements: Verify user can login with new password
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$Username,

    [Parameter(Mandatory = $true)]
    [SecureString]$NewPassword
)

Import-Module ActiveDirectory

Write-Output "Starting password change process for user $Username..."

try {
    Set-ADAccountPassword -Identity $Username -NewPassword $NewPassword -Reset
    Write-Output "Password changed successfully for user $Username"
} catch {
    Write-Error "Failed to change password: $($_.Exception.Message)"
}
