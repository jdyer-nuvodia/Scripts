# =============================================================================
# Script: Change-LocalUserPassword.ps1
# Created: 2025-01-09 16:45:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-26 23:29:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1
# Additional Info: Added parameter support and documentation
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
    The new password to set for the account
.EXAMPLE
    .\Change-LocalUserPassword.ps1 -Username "localuser" -NewPassword "NewPass123!"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$Username,
    
    [Parameter(Mandatory=$true)]
    [string]$NewPassword
)

$SecurePassword = ConvertTo-SecureString $NewPassword -AsPlainText -Force
Set-LocalUser -Name $Username -Password $SecurePassword
