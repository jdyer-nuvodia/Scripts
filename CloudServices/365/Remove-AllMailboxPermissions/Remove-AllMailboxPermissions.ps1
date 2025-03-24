# =============================================================================
# Script: Remove-AllMailboxPermissions.ps1
# Created: 2024-02-21 12:00:00 UTC
# Author: nunya-nunya
# Last Updated: 2024-02-21 13:00:00 UTC
# Updated By: nunya-nunya
# Version: 1.1
# Additional Info: Added parameter support and relative path for mailboxes.txt
# =============================================================================

<#
.SYNOPSIS
    Removes all mailbox permissions including FullAccess, SendAs, Send on Behalf, and Calendar permissions.
.DESCRIPTION
    This script removes various mailbox permissions for specified mailboxes:
     - Removes FullAccess permissions for all delegates
     - Removes SendAs permissions for all trustees
     - Removes Send on Behalf permissions
     - Removes Calendar permissions (except Default and Anonymous)
     Dependencies:
     - Exchange Online PowerShell module
     - Connection to Exchange Online
     - Optional: mailboxes.txt file in script directory
.PARAMETER MailboxIdentity
    Optional. Specify a single mailbox to process. If not specified, script will read from mailboxes.txt in the script directory.
.EXAMPLE
    .\Remove-AllMailboxPermissions.ps1
    Removes permissions for all mailboxes listed in mailboxes.txt
.EXAMPLE
    .\Remove-AllMailboxPermissions.ps1 -MailboxIdentity "user@domain.com"
    Removes all permissions for the specified mailbox
.NOTES
    Security Level: High
    Required Permissions: Exchange Administrator
    Validation Requirements: 
    - Verify Exchange Online PowerShell module is installed
    - Verify Exchange Online connection credentials
    - If using mailboxes.txt, verify file exists in script directory
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$MailboxIdentity
)

# Import the Exchange Online PowerShell module
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
Connect-ExchangeOnline

# Determine mailbox source
if ($MailboxIdentity) {
    $mailboxes = @($MailboxIdentity)
    Write-Host "Processing single mailbox: $MailboxIdentity" -ForegroundColor Cyan
} else {
    $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    $mailboxListPath = Join-Path $scriptPath "mailboxes.txt"
    
    if (Test-Path $mailboxListPath) {
        $mailboxes = Get-Content $mailboxListPath
        Write-Host "Processing mailboxes from: $mailboxListPath" -ForegroundColor Cyan
    } else {
        Write-Error "mailboxes.txt not found in script directory and no mailbox specified."
        exit 1
    }
}

foreach ($mailbox in $mailboxes) {
    # Get all delegates with FullAccess permissions
    $delegates = Get-MailboxPermission -Identity $mailbox | Where-Object {
        $_.IsInherited -eq $false -and 
        $_.User -ne "NT AUTHORITY\SELF" -and 
        $_.AccessRights -like "*FullAccess*"
    }

    # Remove FullAccess permissions for each delegate
    foreach ($delegate in $delegates) {
        Remove-MailboxPermission -Identity $mailbox -User $delegate.User -AccessRights FullAccess -Confirm:$false
        Write-Output "Removed FullAccess permission for $($delegate.User) on $mailbox"
    }

    # Remove Send As permissions
    $sendAsPermissions = Get-RecipientPermission -Identity $mailbox | Where-Object { $_.IsInherited -eq $false -and $_.Trustee -ne "NT AUTHORITY\SELF" }
    foreach ($permission in $sendAsPermissions) {
        Remove-RecipientPermission -Identity $mailbox -Trustee $permission.Trustee -AccessRights SendAs -Confirm:$false
        Write-Output "Removed SendAs permission for $($permission.Trustee) on $mailbox"
    }

    # Remove Send on Behalf permissions
    Set-Mailbox -Identity $mailbox -GrantSendOnBehalfTo $null
    Write-Output "Removed all Send on Behalf permissions for $mailbox"
}

foreach ($mailbox in $mailboxes) {
    # Get all calendar permissions
    $calendarPermissions = Get-MailboxFolderPermission -Identity ${$mailbox:\Calendar}

    # Remove all calendar permissions except for Default and Anonymous
    foreach ($permission in $calendarPermissions) {
        if ($permission.User.DisplayName -notin @("Default", "Anonymous")) {
            Remove-MailboxFolderPermission -Identity ${$mailbox:\Calendar} -User $permission.User.DisplayName -Confirm:$false
            Write-Output "Removed calendar permission for $($permission.User.DisplayName) on $mailbox"
        }
    }
}

