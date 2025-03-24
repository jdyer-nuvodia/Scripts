# =============================================================================
# Script: Grant-CalendarPermissions.ps1
# Created: 2024-02-20 17:15:00 UTC
# Author: nunya-nunya
# Last Updated: 2024-02-20 18:00:00 UTC
# Updated By: nunya-nunya
# Version: 1.2
# Additional Info: Added process logging and security notes
# =============================================================================

<#
.SYNOPSIS
    Grants calendar permissions to a specified user for multiple mailboxes.
.DESCRIPTION
    This script grants Editor access with delegate permissions to calendars for multiple mailboxes
    listed in a text file. It requires Exchange Online PowerShell module.
    
    The script includes logging of all operations and proper error handling.
.PARAMETER UserName
    The email address of the user who will receive calendar access permissions.
.EXAMPLE
    .\Grant-CalendarPermissions.ps1 -UserName "john.doe@contoso.com"
.NOTES
    Security Level: Medium
    Required Permissions: Exchange Administrator or Organization Management
    Validation Requirements: Verify access after granting permissions
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$UserName
)

# Import the Exchange Online PowerShell module
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
Connect-ExchangeOnline

# Read the list of mailboxes from a text file
$mailboxes = Get-Content "C:\Users\nunya\OneDrive - nunya\Documents\GitHub\Scripts\365\Grant-CalendarPermissions\mailboxes.txt"

foreach ($mailboxEmail in $mailboxes) {
    Write-Host "Processing calendar permissions for $mailboxEmail..." -ForegroundColor Cyan
    $calendarPath = "${mailboxEmail}:\Calendar"
    
    try {
        # Check if the mailbox exists and store the result
        if ($null -eq (Get-Mailbox -Identity $mailboxEmail -ErrorAction Stop)) {
            Write-Host "Mailbox $mailboxEmail not found" -ForegroundColor Yellow
            continue
        }
        
        # Set calendar permissions
        Set-MailboxFolderPermission -Identity $calendarPath -User $UserName -AccessRights Editor -SharingPermissionFlags Delegate,CanViewPrivateItems
        Write-Host "Successfully granted Editor access to $mailboxEmail's calendar for user $UserName" -ForegroundColor Green
    }
    catch {
        Write-Host "Error processing $mailboxEmail's calendar: $_" -ForegroundColor Red
    }
}
