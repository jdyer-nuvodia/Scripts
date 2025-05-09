# =============================================================================
# Script: Delete-Mailboxes.ps1
# Created: 2025-02-20 17:15:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-08 19:30:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2.0
# Additional Info: Added SupportsShouldProcess for safer mailbox deletion
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
     
    Supports -WhatIf parameter to preview changes without making them.
     
    Dependencies:
     - Exchange Management Shell
     - Text file containing list of mailboxes (one per line)
     
    Security considerations:
     - Requires Exchange administrator privileges
     - Supports -WhatIf for previewing changes
     - Supports -Confirm for individual confirmations
.PARAMETER userListPath
    Path to the text file containing the list of mailboxes to delete (default: UserList.txt in script directory)
.EXAMPLE
    .\Delete-Mailboxes.ps1
    Reads UserList.txt and attempts to delete all mailboxes listed in the file
.EXAMPLE
    .\Delete-Mailboxes.ps1 -WhatIf
    Shows what mailboxes would be deleted without actually deleting them
.NOTES
    Security Level: High
    Required Permissions: Exchange Administrator
    Validation Requirements: Verify mailbox list before execution
#>

[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='High')]
param(
    [Parameter(Mandatory=$false)]
    [ValidateScript({
        if(-Not ($_ | Test-Path) ){
            throw "File does not exist: $_"
        }
        if(-Not ($_ | Test-Path -PathType Leaf) ){
            throw "The Path argument must be a file"
        }
        return $true
    })]
    [System.IO.FileInfo]$userListPath = (Join-Path $PSScriptRoot "UserList.txt")
)

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
        if ($PSCmdlet.ShouldProcess($user, "Delete mailbox")) {
            Remove-Mailbox -Identity $user -Confirm:$false -ErrorAction Stop
            Write-Host "Successfully deleted mailbox for: $user" -ForegroundColor Green
            $successCount++
        }
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
