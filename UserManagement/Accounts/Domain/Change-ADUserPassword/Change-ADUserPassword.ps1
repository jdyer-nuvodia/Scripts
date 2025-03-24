# =============================================================================
# Script: Change-ADUserPassword.ps1
# Created: 2024-02-20 17:15:00 UTC
# Author: nunya-nunya
# Last Updated: 2024-02-20 17:45:00 UTC
# Updated By: nunya-nunya
# Version: 1.1
# Additional Info: Added parameter support and color-coded output
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
    The new password to set for the user account
.EXAMPLE
    .\Change-ADUserPassword.ps1 -Username "jsmith" -NewPassword "NewP@ssw0rd123!"
    Changes password for user jsmith to the specified password
.NOTES
    Security Level: High
    Required Permissions: Domain Admin or delegated AD password reset rights
    Validation Requirements: Verify user can login with new password
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$Username,
    
    [Parameter(Mandatory=$true)]
    [string]$NewPassword
)

Import-Module ActiveDirectory

Write-Host "Starting password change process for user $Username..." -ForegroundColor Cyan

try {
    $SecurePassword = ConvertTo-SecureString $NewPassword -AsPlainText -Force
    Set-ADAccountPassword -Identity $Username -NewPassword $SecurePassword -Reset
    Write-Host "Password changed successfully for user $Username" -ForegroundColor Green
} catch {
    Write-Error "Failed to change password: $($_.Exception.Message)"
}
