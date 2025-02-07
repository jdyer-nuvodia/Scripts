# Import the Exchange Online PowerShell module
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
Connect-ExchangeOnline

#Set variable for SG
$securityGroupName = "ConfRmCal-Author"

# Read the list of mailboxes from a text file
$mailboxes = Get-Content "C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\grantCalendarPermissions\mailboxes.txt"

foreach ($mailbox in $mailboxes) {
    $identity = $mailbox.UserPrincipalName + ":\Calendar"
    
    try {
        Set-MailboxFolderPermission -Identity $identity -User $securityGroupName -AccessRights CreateItems, ReadItems, FolderVisible -ErrorAction Stop
        Write-Host "Successfully granted $calendarPermission access to $($mailbox.UserPrincipalName)'s calendar for group $securityGroupName" -ForegroundColor Green
    }
    catch {
        Write-Host "Error granting access to $($mailbox.UserPrincipalName)'s calendar: $_" -ForegroundColor Red
    }
}