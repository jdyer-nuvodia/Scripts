# =============================================================================
# Script: Grant-RMToMailboxEditCalendarPermissions.ps1
# Created: 2024-02-20 17:30:00 UTC
# Author: nunya-nunya
# Last Updated: 2024-02-20 18:00:00 UTC
# Updated By: nunya-nunya
# Version: 1.2
# Additional Info: Added single mailbox parameter option
# =============================================================================

<#
.SYNOPSIS
    Grants mailbox and calendar permissions to a specified user
.DESCRIPTION
    Grants Full Access to mailboxes and Editor rights to calendars for a specified user.
    Can process either a single mailbox or multiple mailboxes from a text file.
.PARAMETER UserEmail
    Email address of the user to grant permissions to
.PARAMETER SingleMailbox
    Optional: Single mailbox to process instead of reading from mailboxes.txt
.EXAMPLE
    .\Grant-RMToMailboxEditCalendarPermissions.ps1 -UserEmail "john.doe@domain.com"
    Process all mailboxes from mailboxes.txt
.EXAMPLE
    .\Grant-RMToMailboxEditCalendarPermissions.ps1 -UserEmail "john.doe@domain.com" -SingleMailbox "user@domain.com"
    Process only the specified mailbox
.NOTES
    Security Level: High
    Required Permissions: Exchange Admin rights
#>

param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$')]
    [string]$UserEmail,
    
    [Parameter(Mandatory = $false)]
    [ValidatePattern('^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$')]
    [string]$SingleMailbox
)

# Import the Exchange Online PowerShell module
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
Connect-ExchangeOnline

# Determine mailbox source
$mailboxes = if ($SingleMailbox) {
    @($SingleMailbox)
} else {
    Get-Content (Join-Path $PSScriptRoot "mailboxes.txt")
}

foreach ($mailbox in $mailboxes) {    
    $identity = $mailbox.UserPrincipalName + ":\Calendar"
    
    Add-MailboxPermission -Identity $mailbox -User $UserEmail -AccessRights FullAccess -InheritanceType All -AutoMapping:$false
    Add-MailboxFolderPermission -Identity $Identity -User $UserEmail -AccessRights Editor -SharingPermissionFlags Delegate,CanViewPrivateItems
}
