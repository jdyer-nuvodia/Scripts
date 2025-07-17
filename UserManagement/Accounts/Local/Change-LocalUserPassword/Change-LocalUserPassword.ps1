# =============================================================================
# Script: Change-LocalUserPassword.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2.1
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
# =============================================================================

<#
.SYNOPSIS
    Changes the password for a local user account.
.DESCRIPTION
    This script allows changing the password of a specified local user account
    using secure string conversion for password handling.
.PARAMETER Username
    The username of the local account to modify
.PARAMETER NewPassword
    The new password to set for the account (SecureString)
.EXAMPLE
    $SecurePass = ConvertTo-SecureString "NewPass123!" -AsPlainText -Force
    .\Change-LocalUserPassword.ps1 -Username "localuser" -NewPassword $SecurePass
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$Username,

    [Parameter(Mandatory = $true)]
    [SecureString]$NewPassword
)

Set-LocalUser -Name $Username -Password $NewPassword
