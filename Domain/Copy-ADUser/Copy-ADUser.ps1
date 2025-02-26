# =============================================================================
# Script: Copy-ADUser.ps1
# Created: 2024-02-20 17:15:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2024-02-20 17:15:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0
# Additional Info: Initial script creation with standard header format
# =============================================================================

<#
.SYNOPSIS
    Copies an existing AD user's group memberships to a new user account.
.DESCRIPTION
    This script creates a new Active Directory user account and copies all group
    memberships from a specified source user. The script:
    - Creates new user with specified properties
    - Copies group memberships from source user
    - Enables the account with a specified password
    Dependencies:
    - Active Directory PowerShell module
    - Domain Admin or appropriate AD delegation rights
.PARAMETER sourceUser
    The username of the existing AD user to copy from
.PARAMETER newUserName
    The new username to be created
.PARAMETER newUserGivenName
    The given name for the new user
.PARAMETER newUserSurname
    The surname for the new user
.PARAMETER newUserPassword
    The initial password for the new user
.PARAMETER newUserDescription
    The description for the new user account
.EXAMPLE
    .\Copy-ADUser.ps1
    Creates a new user with predefined parameters and copies group memberships
.NOTES
    Security Level: High
    Required Permissions: Domain Admin or delegated AD user creation rights
    Validation Requirements: Verify source user exists, new username doesn't exist
#>

# Define the source user and the new user's details
$sourceUser = "pa-gbullock"   # Replace with the username of the user to be copied
$newUserName = "pa-jdyer"     # Replace with the new user's username
$newUserGivenName = "JB"  # Replace with the new user's given name
$newUserSurname = "Dyer"      # Replace with the new user's surname
$newUserPassword = "12ravenousgiantpandaS!" # Replace with the new user's password
$newUserDescription = "Nuvodia" # Replace with the new user's description

# Load the Active Directory module
Import-Module ActiveDirectory

# Get the source user's details
$sourceUserDetails = Get-ADUser -Identity $sourceUser -Properties *

# Create the new user with the different name properties and description
New-ADUser `
    -Name "$newUserGivenName $newUserSurname" `
    -GivenName $newUserGivenName `
    -Surname $newUserSurname `
    -SamAccountName $newUserName `
    -UserPrincipalName "$newUserName@$(($sourceUserDetails.UserPrincipalName).Split('@')[1])" `
    -Path $sourceUserDetails.DistinguishedName `
    -Enabled $true `
    -AccountPassword (ConvertTo-SecureString $newUserPassword -AsPlainText -Force) `
    -Description $newUserDescription

# Add the new user to the same groups as the source user
$sourceUserGroups = Get-ADUser -Identity $sourceUser -Properties MemberOf | Select-Object -ExpandProperty MemberOf
foreach ($group in $sourceUserGroups) {
    Add-ADGroupMember -Identity $group -Members $newUserName
    Write-Host "Added $newUserName to group $group" -ForegroundColor Cyan
}

Write-Host "New user $newUserName created successfully!" -ForegroundColor Green
Write-Host "Group memberships copied from $sourceUser" -ForegroundColor Green
