# Import the Exchange Online PowerShell module
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
Connect-ExchangeOnline

# Read the list of mailboxes from a text file
$mailboxes = $mailboxes = Get-Content "C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\grantRMToMailbox+EditCalendarPermissions\mailboxes.txt"

#Set the user to receive the permissions
$user = "chagan@workwith.com"

foreach ($mailbox in $mailboxes) {    
	$identity = $mailbox.UserPrincipalName + ":\Calendar"
	
	Add-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -InheritanceType All -AutoMapping:$false
    Add-MailboxFolderPermission -Identity $Identity -User $user -AccessRights Editor -SharingPermissionFlags Delegate,CanViewPrivateItems
}