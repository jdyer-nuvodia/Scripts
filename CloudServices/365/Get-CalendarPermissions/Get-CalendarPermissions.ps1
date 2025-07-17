# =============================================================================
# Script: Get-CalendarPermissions.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.2
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
# =============================================================================

<#
.SYNOPSIS
    Retrieves calendar permissions for a list of mailboxes.
.DESCRIPTION
    This script reads a list of mailboxes from a text file and retrieves the calendar
    permissions for each mailbox. It includes error handling for failed permission checks.

    Dependencies:
    - Exchange Online PowerShell module
    - Appropriate Exchange Online permissions
    - Text file containing mailbox list
.PARAMETER None
    No parameters required. Mailbox list is read from a fixed path.
.EXAMPLE
    .\Get-CalendarPermissions.ps1
    Retrieves calendar permissions for all mailboxes listed in the mailboxes.txt file.
.NOTES
    Security Level: Medium
    Required Permissions: Exchange Online View-Only Recipients role
    Validation Requirements: Verify Exchange Online connection before running
#>

# Read the list of mailboxes from a text file in the same directory as the script
$mailboxesFile = Join-Path -Path $PSScriptRoot -ChildPath "mailboxes.txt"

Write-Output "Reading mailboxes from: $mailboxesFile"

if (-not (Test-Path -Path $mailboxesFile)) {
    Write-Error "Mailboxes file not found: $mailboxesFile"
    exit
}

$mailboxes = Get-Content -Path $mailboxesFile

foreach ($mailbox in $mailboxes) {
    $identity = $mailbox.UserPrincipalName + ":\Calendar"

    try {
        Get-MailboxFolderPermission -Identity $identity -ErrorAction Stop
    } catch {
        Write-Error "Error accessing $($mailbox.UserPrincipalName)'s calendar: $_"
    }
}
