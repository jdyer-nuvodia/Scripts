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

