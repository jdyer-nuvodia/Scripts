# =============================================================================
# Script: Add-UserListTo365Group.ps1
# Created: 2024-02-20 17:15:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-02 20:45:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1.0
# Additional Info: Enhanced documentation and parameter validation
# =============================================================================

<#
.SYNOPSIS
    Adds users from a CSV file to a specified Microsoft 365 distribution group.
.DESCRIPTION
    This script reads a CSV file containing user principal names and adds each user
    to a specified Microsoft 365 distribution group. The script bypasses security
    group manager check for bulk operations.
    
    Dependencies:
    - Exchange Online PowerShell Module (Install-Module -Name ExchangeOnlineManagement)
    - CSV file with UserPrincipalName column
    - Appropriate permissions to modify distribution groups
    
    The CSV file must contain at minimum:
    - UserPrincipalName: The email or UPN of the user to add
    
    The script will:
    1. Validate the CSV file exists and is properly formatted
    2. Check if the distribution group exists
    3. Add each user to the group
    4. Log success and failures
.PARAMETER CsvPath
    Path to the CSV file containing user principal names. Must contain a 'UserPrincipalName' column.
    Default is 'users.csv' in the script directory.
.PARAMETER GroupName
    Name of the Microsoft 365 distribution group to add users to.
    Default is 'ConfRmCal - Author'.
.EXAMPLE
    .\Add-UserListTo365Group.ps1
    Adds all users from the default users.csv to the default group 'ConfRmCal - Author'
.EXAMPLE
    .\Add-UserListTo365Group.ps1 -CsvPath "C:\Users\admin\Desktop\users.csv" -GroupName "Sales Team"
    Adds users from the specified CSV file to the "Sales Team" distribution group
.NOTES
    Security Level: Medium
    Required Permissions: Exchange Online Administrator
    Validation Requirements: 
    - Verify CSV file exists and is accessible
    - Verify distribution group exists
    - Verify current user has appropriate permissions
    Change Log:
    - v1.1.0: Added parameter validation and enhanced error handling
    - v1.0.0: Initial script creation
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [ValidateScript({
        if(-Not ($_ | Test-Path) ){
            throw "File or folder does not exist"
        }
        if(-Not ($_ | Test-Path -PathType Leaf) ){
            throw "The Path argument must be a file"
        }
        if($_ -notmatch "(\.csv)"){
            throw "The file specified must be a csv"
        }
        return $true
    })]
    [System.IO.FileInfo]$CsvPath = (Join-Path $PSScriptRoot "users.csv"),

    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$GroupName = "ConfRmCal - Author"
)

# Initialize logging
$LogPath = Join-Path $PSScriptRoot "GroupAdditions_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$ErrorActionPreference = "Stop"

function Write-Log {
    param($Message, $Level = "Information")
    
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC"
    $LogMessage = "$TimeStamp [$Level] $Message"
    Add-Content -Path $LogPath -Value $LogMessage
    
    switch ($Level) {
        "Information" { Write-Host $Message -ForegroundColor White }
        "Success"     { Write-Host $Message -ForegroundColor Green }
        "Warning"     { Write-Host $Message -ForegroundColor Yellow }
        "Error"       { Write-Host $Message -ForegroundColor Red }
    }
}

try {
    Write-Log "Starting user addition process for group: $GroupName" "Information"
    Write-Log "Using CSV file: $CsvPath" "Information"

    # Validate CSV content
    $users = Import-Csv -Path $CsvPath
    if (-not ($users | Get-Member -Name "UserPrincipalName")) {
        throw "CSV file must contain 'UserPrincipalName' column"
    }

    # Verify group exists
    try {
        $null = Get-DistributionGroup -Identity $GroupName -ErrorAction Stop
        Write-Log "Verified group '$GroupName' exists" "Success"
    }
    catch {
        throw "Distribution group '$GroupName' not found or access denied: $_"
    }

    $successCount = 0
    $failureCount = 0

    foreach ($user in $users) {
        try {
            Add-DistributionGroupMember -Identity $GroupName -Member $user.UserPrincipalName -BypassSecurityGroupManagerCheck -ErrorAction Stop
            Write-Log "Successfully added user: $($user.UserPrincipalName)" "Success"
            $successCount++
        }
        catch {
            Write-Log "Failed to add user $($user.UserPrincipalName): $_" "Error"
            $failureCount++
        }
    }

    Write-Log "Operation complete. Successfully added: $successCount users. Failed: $failureCount users" "Information"
}
catch {
    Write-Log "Script execution failed: $_" "Error"
    throw
}
finally {
    Write-Log "Script execution finished" "Information"
}
