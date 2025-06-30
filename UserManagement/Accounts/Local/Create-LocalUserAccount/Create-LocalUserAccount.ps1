# =============================================================================
# Script: Create-LocalAccount.ps1
# Created: 2025-01-23 15:30:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-26 23:30:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1
# Additional Info: Added parameterization to script
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
    The password for the new local account
.PARAMETER FullName
    The full name of the user
.PARAMETER Description
    Description for the user account
.EXAMPLE
    .\Create-LocalAccount.ps1 -Username "jsmith" -Password "Password123!" -FullName "John Smith" -Description "Local account for John Smith - #12345"
.NOTES
    Security Level: Medium
    Required Permissions: Administrative rights
    Validation Requirements: Verify account creation and group membership
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$Username,

    [Parameter(Mandatory=$true)]
    [string]$Password,

    [Parameter(Mandatory=$false)]
    [string]$FullName,

    [Parameter(Mandatory=$false)]
    [string]$Description
)

# Convert password to secure string
$SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force

# Create user Account
Write-Host "Creating new local user account..." -ForegroundColor Cyan
New-LocalUser -Name $Username -Password $SecurePassword -FullName $FullName -Description $Description

# Force Password Change
Write-Host "Configuring password change requirement..." -ForegroundColor Cyan
net user $Username /logonpasswordchg:yes

# Ensure the account is active
Write-Host "Enabling user account..." -ForegroundColor Cyan
Enable-LocalUser -Name $Username

# Add user to Users group
Write-Host "Adding user to Users group..." -ForegroundColor Cyan
Add-LocalGroupMember -Group Users -Member $Username

Write-Host "Local user account creation completed successfully." -ForegroundColor Green
