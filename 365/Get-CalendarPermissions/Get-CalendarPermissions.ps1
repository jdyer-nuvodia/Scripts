# Read the list of mailboxes from a text file
$mailboxes = Get-Content "C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\getCalendarPermissions\mailboxes.txt"

foreach ($mailbox in $mailboxes) {
    $identity = $mailbox.UserPrincipalName + ":\Calendar"
    
    try {
        Get-MailboxFolderPermission -Identity $identity -ErrorAction Stop
    }
    catch {
        Write-Host "Error granting access to $($mailbox.UserPrincipalName)'s calendar: $_" -ForegroundColor Red
    }
}