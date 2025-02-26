# =============================================================================
# Script: Remove-AllMailboxPermissions.ps1
# Created: 2024-02-21 12:00:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2024-02-21 12:00:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0
# Additional Info: Initial script creation for removing mailbox permissions
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
     - Text file containing mailbox list
.PARAMETER None
    Script uses a hardcoded path to mailboxes.txt file
.EXAMPLE
    .\Remove-AllMailboxPermissions.ps1
    Removes all permissions for mailboxes listed in the mailboxes.txt file
.NOTES
    Security Level: High
    Required Permissions: Exchange Administrator
    Validation Requirements: 
    - Verify Exchange Online PowerShell module is installed
    - Verify access to mailboxes.txt file
    - Verify Exchange Online connection credentials
#>

# Import the Exchange Online PowerShell module
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
Connect-ExchangeOnline

# Read the list of mailboxes from a text file
$mailboxes = Get-Content "C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\removeAllMailboxPermissions\mailboxes.txt"

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

