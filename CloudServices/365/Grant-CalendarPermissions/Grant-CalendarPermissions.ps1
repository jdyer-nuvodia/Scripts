# =============================================================================
# Script: Grant-CalendarPermissions.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.4.1
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
# =============================================================================

<#
.SYNOPSIS
    Grants calendar permissions to a specified user for multiple mailboxes.
.DESCRIPTION
    This script grants Editor access with delegate permissions to calendars for multiple mailboxes
    listed in a text file. It requires Exchange Online PowerShell module.

    Supports -WhatIf parameter to preview changes without making them.

    The script includes logging of all operations and proper error handling.
.PARAMETER UserName
    The email address of the user who will receive calendar access permissions.
.EXAMPLE
    .\Grant-CalendarPermissions.ps1 -UserName "john.doe@contoso.com"
.EXAMPLE
    .\Grant-CalendarPermissions.ps1 -UserName "john.doe@contoso.com" -WhatIf
    Shows what changes would be made without actually making them.
.NOTES
    Security Level: Medium
    Required Permissions: Exchange Administrator or Organization Management
    Validation Requirements: Verify access after granting permissions
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $true)]
    [string]$UserName
)

# Import the Exchange Online PowerShell module
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
Connect-ExchangeOnline

# Read the list of mailboxes from a text file
$mailboxes = Get-Content "$PSScriptRoot\mailboxes.txt"

foreach ($mailboxEmail in $mailboxes) {
    Write-Output "Processing calendar permissions for $mailboxEmail..."
    $calendarPath = "${ mailboxEmail}:\Calendar"

    try {
        # Check if the mailbox exists and store the result
        if ($null -eq (Get-Mailbox -Identity $mailboxEmail -ErrorAction Stop)) {
            Write-Warning "Mailbox $mailboxEmail not found"
            continue
        }

        # Set calendar permissions
        if ($PSCmdlet.ShouldProcess($calendarPath, "Grant Editor calendar permissions to $UserName")) {
            Set-MailboxFolderPermission -Identity $calendarPath -User $UserName -AccessRights Editor -SharingPermissionFlags Delegate, CanViewPrivateItems
            Write-Output "Successfully granted Editor access to $mailboxEmail's calendar for user $UserName"
        }
    } catch {
        Write-Error "Error processing $mailboxEmail's calendar: $_"
    }
}
