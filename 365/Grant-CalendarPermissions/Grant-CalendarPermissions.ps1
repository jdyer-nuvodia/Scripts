# Import the Exchange Online PowerShell module
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
Connect-ExchangeOnline

# Set variable for user to receive access
$userName = "mailbox"

# Read the list of mailboxes from a text file, set this to the folder where this script is located
$mailboxes = Get-Content "C:\Users\jdyer\OneDrive - Nuvodia\Documents\GitHub\Scripts\365\Grant-CalendarPermissions\mailboxes.txt"

foreach ($mailboxEmail in $mailboxes) {
    $identity = "${mailboxEmail}:\Calendar"
    
    try {
        # Check if the mailbox exists
        $mailbox = Get-Mailbox -Identity $mailboxEmail -ErrorAction Stop
        
        # Set calendar permissions
        Set-MailboxFolderPermission -Identity $identity -User $userName -AccessRights Editor -SharingPermissionFlags Delegate,CanViewPrivateItems
        Write-Host "Successfully granted Editor access to $mailboxEmail's calendar for user $userName" -ForegroundColor Green
    }
    catch {
        Write-Host "Error granting access to $mailboxEmail's calendar: $_" -ForegroundColor Red
    }
}