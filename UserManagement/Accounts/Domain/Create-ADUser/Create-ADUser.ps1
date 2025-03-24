# =============================================================================
# Script: Create-ADUser.ps1
# Created: 2024-02-20 17:15:00 UTC
# Author: jdyer-nuvodia
# Version: 1.1
# Additional Info: Added parameters and secure password handling
# =============================================================================

<#
.SYNOPSIS
    Creates a new Active Directory user and adds them to specified groups.
.DESCRIPTION
    Creates a new AD user with specified parameters and adds them to designated AD groups.
    Handles password conversion securely during runtime.
.PARAMETER Name
    Full name of the user
.PARAMETER GivenName
    First name of the user
.PARAMETER Surname
    Last name of the user
.PARAMETER SamAccountName
    SAM account name for the user
.PARAMETER UserPrincipalName
    UPN for the user (email format)
.PARAMETER Password
    Initial password for the user
.PARAMETER OUPath
    OU path where the user will be created
.PARAMETER Groups
    Array of groups to add the user to
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$Name,
    
    [Parameter(Mandatory = $true)]
    [string]$GivenName,
    
    [Parameter(Mandatory = $true)]
    [string]$Surname,
    
    [Parameter(Mandatory = $true)]
    [string]$SamAccountName,
    
    [Parameter(Mandatory = $true)]
    [string]$UserPrincipalName,
    
    [Parameter(Mandatory = $true)]
    [string]$Password,
    
    [Parameter(Mandatory = $true)]
    [string]$OUPath,
    
    [Parameter(Mandatory = $false)]
    [string[]]$Groups = @()
)

# Convert the password to secure string at runtime
$SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force

# Create the new user
New-ADUser -Name $Name `
    -GivenName $GivenName `
    -Surname $Surname `
    -SamAccountName $SamAccountName `
    -UserPrincipalName $UserPrincipalName `
    -AccountPassword $SecurePassword `
    -Enabled $true `
    -Path $OUPath

# Add user to specified groups
foreach ($Group in $Groups) {
    Add-ADGroupMember -Identity $Group -Members $SamAccountName
}

