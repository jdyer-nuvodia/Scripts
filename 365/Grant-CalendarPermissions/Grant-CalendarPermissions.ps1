# Import the Exchange Online PowerShell module
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
Connect-ExchangeOnline

#Set variable for SG if adding to a group, be sure to change the $userName variable in the for loop to $securityGroupName
#$securityGroupName = "ConfRmCal-Author"

#Set variable for user to receive access
$userName = "mailbox"

# Read the list of mailboxes from a text file, set this to the folder where this script 
$mailboxes = Get-Content "C:\Users\jdyer\OneDrive - Nuvodia\Documents\GitHub\Scripts\365\Grant-CalendarPermissions\mailboxes.txt"

foreach ($mailbox in $mailboxes) {
    $identity = $mailbox.UserPrincipalName + ":\Calendar"
    
    try {
        Set-MailboxFolderPermission -Identity $identity -User $userName -AccessRights Editor -SharingPermissionFlags Delegate,CanViewPrivateItems
        Write-Host "Successfully granted $calendarPermission access to $($mailbox.UserPrincipalName)'s calendar for group $securityGroupName" -ForegroundColor Green
    }
    catch {
        Write-Host "Error granting access to $($mailbox.UserPrincipalName)'s calendar: $_" -ForegroundColor Red
    }
}