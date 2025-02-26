# =============================================================================
# Script: Change-ADUserPassword.ps1
# Created: 2024-02-20 17:15:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2024-02-20 17:15:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0
# Additional Info: Initial script creation with standard header
# =============================================================================

<#
.SYNOPSIS
    Changes the password for an Active Directory user account.
.DESCRIPTION
    This script resets the password for a specified Active Directory user account.
    - Requires Active Directory PowerShell module
    - Must be run with appropriate AD permissions
    - Handles errors during password reset process
.PARAMETER username
    The SAM account name of the AD user whose password needs to be changed
.PARAMETER newPassword
    The new password to set for the user account
.EXAMPLE
    .\Change-ADUserPassword.ps1
    Changes password for hard-coded username to specified password
.NOTES
    Security Level: High
    Required Permissions: Domain Admin or delegated AD password reset rights
    Validation Requirements: Verify user can login with new password
#>

Import-Module ActiveDirectory

$username = "username"
$newPassword = ConvertTo-SecureString "Password123!" -AsPlainText -Force

try {
    Set-ADAccountPassword -Identity $username -NewPassword $newPassword -Reset
    Write-Host "Password changed successfully for user $username"
} catch {
    Write-Host "Failed to change password. Error: $($_.Exception.Message)"
}
