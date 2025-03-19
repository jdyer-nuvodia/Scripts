# =============================================================================
# Script: Add-UserListTo365Group.ps1
# Created: 2024-02-20 17:15:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2024-02-20 17:15:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0
# Additional Info: Initial script creation with proper header format
# =============================================================================

<#
.SYNOPSIS
    Adds users from a CSV file to a specified Microsoft 365 distribution group.
.DESCRIPTION
    This script reads a CSV file containing user principal names and adds each user
    to a specified Microsoft 365 distribution group. The script bypasses security
    group manager check for bulk operations.
    
    Dependencies:
    - Exchange Online PowerShell Module
    - CSV file with UserPrincipalName column
    - Appropriate permissions to modify distribution groups
.PARAMETER csvPath
    Path to the CSV file containing user principal names
.PARAMETER groupName
    Name of the Microsoft 365 distribution group
.EXAMPLE
    .\Add-UserListTo365Group.ps1
    Adds all users from the default CSV path to the specified distribution group
.NOTES
    Security Level: Medium
    Required Permissions: Exchange Online Administrator
    Validation Requirements: 
    - Verify CSV file exists and is accessible
    - Verify distribution group exists
    - Verify current user has appropriate permissions
#>

$csvPath = Join-Path $PSScriptRoot "users.csv"
$groupName = "ConfRmCal - Author"

$users = Import-Csv -Path $csvPath

foreach ($user in $users) {
    Add-DistributionGroupMember -Identity $groupName -Member $user.UserPrincipalName -BypassSecurityGroupManagerCheck
}
