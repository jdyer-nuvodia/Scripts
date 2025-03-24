# =============================================================================
# Script: Delete-Mailboxes.ps1
# Created: 2024-02-20 17:15:00 UTC
# Author: nunya-nunya
# Last Updated: 2024-02-21 17:15:00 UTC
# Updated By: nunya-nunya
# Version: 1.1
# Additional Info: Updated UserList.txt path to use script directory
# =============================================================================

<#
.SYNOPSIS
    Deletes multiple Exchange mailboxes from a list provided in a text file.
.DESCRIPTION
    This script automates the process of deleting multiple Exchange mailboxes by:
     - Reading a list of users from a specified text file
     - Verifying Exchange Management Shell is loaded
     - Deleting each mailbox and providing status updates
     - Generating a summary of successful and failed deletions
     
    Dependencies:
     - Exchange Management Shell
     - Text file containing list of mailboxes (one per line)
     
    Security considerations:
     - Requires Exchange administrator privileges
     - No confirmation prompt when deleting mailboxes
.PARAMETER userListPath
    Path to the text file containing the list of mailboxes to delete (default: C:\Temp\UserList.txt)
.EXAMPLE
    .\Delete-Mailboxes.ps1
    Reads C:\Temp\UserList.txt and attempts to delete all mailboxes listed in the file
.NOTES
    Security Level: High
    Required Permissions: Exchange Administrator
    Validation Requirements: Verify mailbox list before execution
#>

# Script to delete multiple mailboxes from a list in a text file

# Path to the text file containing the list of users
$userListPath = Join-Path $PSScriptRoot "UserList.txt"

# Function to check if Exchange Management Shell is loaded
function Test-ExchangeShell {
    if (!(Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
        Write-Host "Exchange Management Shell is not loaded. Please run this script in Exchange Management Shell." -ForegroundColor Red
        return $false
    }
    return $true
}

# Check if Exchange Management Shell is loaded
if (!(Test-ExchangeShell)) {
    exit
}

# Check if the file exists
if (!(Test-Path $userListPath)) {
    Write-Host "The specified file does not exist: $userListPath" -ForegroundColor Red
    exit
}

# Read the list of users from the file
$users = Get-Content $userListPath

# Counter for successful and failed deletions
$successCount = 0
$failCount = 0

# Process each user in the list
foreach ($user in $users) {
    try {
        # Attempt to remove the mailbox
        Remove-Mailbox -Identity $user -Confirm:$false -ErrorAction Stop
        Write-Host "Successfully deleted mailbox for: $user" -ForegroundColor Green
        $successCount++
    }
    catch {
        Write-Host "Failed to delete mailbox for: $user" -ForegroundColor Red
        Write-Host "Error: $_" -ForegroundColor Red
        $failCount++
    }
}

# Display summary
Write-Host "`nDeletion Summary:" -ForegroundColor Cyan
Write-Host "Successfully deleted: $successCount mailbox(es)" -ForegroundColor Green
Write-Host "Failed to delete: $failCount mailbox(es)" -ForegroundColor Red
