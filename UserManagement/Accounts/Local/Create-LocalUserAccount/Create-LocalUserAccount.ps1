# =============================================================================
# Script: Create-LocalUserAccount.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2.2
# Additional Info: Fixed header metadata for workflow validation
# =============================================================================

<#
.SYNOPSIS
    Creates a new local user account with specified parameters.
.DESCRIPTION
    Creates a local user account with the provided username, password, full name,
    and description. The account will be enabled and added to the Users group.
.PARAMETER Username
    The username for the new local account
.PARAMETER Password
    The password for the new local account (SecureString)
.PARAMETER FullName
    The full name of the user
.PARAMETER Description
    Description for the user account
.EXAMPLE
    $SecurePass = ConvertTo-SecureString "Password123!" -AsPlainText -Force
    .\Create-LocalUserAccount.ps1 -Username "jsmith" -Password $SecurePass -FullName "John Smith" -Description "Local account for John Smith - #12345"
.NOTES
    Security Level: Medium
    Required Permissions: Administrative rights
    Validation Requirements: Verify account creation and group membership
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$Username,

    [Parameter(Mandatory = $true)]
    [SecureString]$Password,

    [Parameter(Mandatory = $false)]
    [string]$FullName,

    [Parameter(Mandatory = $false)]
    [string]$Description
)

# Create user Account
Write-Output "Creating new local user account..."
New-LocalUser -Name $Username -Password $Password -FullName $FullName -Description $Description

# Force Password Change
Write-Output "Configuring password change requirement..."
Set-LocalUser -Name $Username -PasswordNeverExpires $false -UserMayChangePassword $true

# Ensure the account is active
Write-Output "Enabling user account..."
Enable-LocalUser -Name $Username

# Add user to Users group
Write-Output "Adding user to Users group..."
Add-LocalGroupMember -Group "Users" -Member $Username

Write-Output "Local user account creation completed successfully."
