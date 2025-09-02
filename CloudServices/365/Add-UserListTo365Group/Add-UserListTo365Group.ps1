# =============================================================================
# Script: Add-UserListTo365Group.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2.2
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
# =============================================================================

<#
.SYNOPSIS
    Adds users from a CSV file to a specified Microsoft 365 distribution group.
.DESCRIPTION
    This script reads a CSV file containing user principal names and adds each user
    to a specified Microsoft 365 distribution group. The script bypasses security
    group manager check for bulk operations.

    Supports -WhatIf parameter to preview changes without making them.

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
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $false)]
    [ValidateScript({
            if (-not ($_ | Test-Path)) {
                throw "File or folder does not exist"
            }
            if (-not ($_ | Test-Path -PathType Leaf)) {
                throw "The Path argument must be a file"
            }
            if ($_ -notmatch "(\.csv)$") {
                throw "The file specified must be a csv"
            }
            return $true
        })]
    [System.IO.FileInfo]$CsvPath = (Join-Path $PSScriptRoot "users.csv"),

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$GroupName = "ConfRmCal - Author"
)

# Initialize logging
$LogPath = Join-Path $PSScriptRoot "GroupAdditions_$($env:COMPUTERNAME)_$((Get-Date).ToUniversalTime().ToString('yyyyMMdd_HHmmss_UTC')).log"
$ErrorActionPreference = "Stop"

function Write-LogMessage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet("Information", "Success", "Warning", "Error")]
        [string]$Level = "Information"
    )

    $TimeStamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss UTC")
    $LogMessage = "$TimeStamp [$Level] $Message"
    Add-Content -Path $LogPath -Value $LogMessage

    switch ($Level) {
        "Information" {
            Write-Information -MessageData $Message -InformationAction Continue
        }
        "Success" {
            Write-Information -MessageData $Message -InformationAction Continue
        }
        "Warning" {
            Write-Warning -Message $Message
        }
        "Error" {
            Write-Error -Message $Message -ErrorAction Continue
        }
    }
}

try {
    Write-LogMessage "Starting user addition process for group: $GroupName" "Information"
    Write-LogMessage "Using CSV file: $CsvPath" "Information"

    # Validate CSV content
    $users = Import-Csv -Path $CsvPath
    if (-not ($users | Get-Member -Name "UserPrincipalName")) {
        throw "CSV file must contain 'UserPrincipalName' column"
    }

    # Verify group exists
    try {
        $null = Get-DistributionGroup -Identity $GroupName -ErrorAction Stop
        Write-LogMessage "Verified group '$GroupName' exists" "Success"
    } catch {
        throw "Distribution group '$GroupName' not found or access denied: $_"
    }

    $successCount = 0
    $failureCount = 0

    foreach ($user in $users) {
        try {
            if ($PSCmdlet.ShouldProcess($user.UserPrincipalName, "Add to distribution group '$GroupName'")) {
                Add-DistributionGroupMember -Identity $GroupName -Member $user.UserPrincipalName -BypassSecurityGroupManagerCheck -ErrorAction Stop
                Write-LogMessage "Successfully added user: $($user.UserPrincipalName)" "Success"
                $successCount++
            }
        } catch {
            Write-LogMessage "Failed to add user $($user.UserPrincipalName): $_" "Error"
            $failureCount++
        }
    }

    Write-LogMessage "Operation complete. Successfully added: $successCount users. Failed: $failureCount users" "Information"
} catch {
    Write-LogMessage "Script execution failed: $_" "Error"
    throw
} finally {
    Write-LogMessage "Script execution finished" "Information"
}
