# =============================================================================
# Script: Grant-RMToMailboxEditCalendarPermissions.ps1
# Created: 2024-02-20 17:30:00 UTC
# Author: jdyer-nuvodia
# Version: 1.1
# Additional Info: Added user parameter and documentation
# =============================================================================

<#
.SYNOPSIS
    Grants mailbox and calendar permissions to a specified user
.DESCRIPTION
    Grants Full Access to mailboxes and Editor rights to calendars for a specified user
    using a list of mailboxes from a text file in the same directory
.PARAMETER UserEmail
    Email address of the user to grant permissions to
.EXAMPLE
    .\Grant-RMToMailboxEditCalendarPermissions.ps1 -UserEmail "john.doe@domain.com"
.NOTES
    Security Level: High
    Required Permissions: Exchange Admin rights
#>

param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$')]
    [string]$UserEmail
)

# Import the Exchange Online PowerShell module
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
Connect-ExchangeOnline

# Read the list of mailboxes from a text file
$mailboxes = Get-Content (Join-Path $PSScriptRoot "mailboxes.txt")

foreach ($mailbox in $mailboxes) {    
    $identity = $mailbox.UserPrincipalName + ":\Calendar"
    
    Add-MailboxPermission -Identity $mailbox -User $UserEmail -AccessRights FullAccess -InheritanceType All -AutoMapping:$false
    Add-MailboxFolderPermission -Identity $Identity -User $UserEmail -AccessRights Editor -SharingPermissionFlags Delegate,CanViewPrivateItems
}