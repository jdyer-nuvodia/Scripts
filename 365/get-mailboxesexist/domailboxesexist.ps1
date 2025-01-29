# Import the list of mailboxes to check
$mailboxList = Get-Content -Path "C:\Users\pa-jdyer\Documents\mailboxes.txt"

# Initialize arrays to store results
$existingMailboxes = @()
$nonExistingMailboxes = @()

# Loop through each mailbox in the list
foreach ($mailbox in $mailboxList) {
    if (Get-Mailbox -Identity $mailbox -ErrorAction SilentlyContinue) {
        $existingMailboxes += $mailbox
        Write-Host "Mailbox exists: $mailbox" -ForegroundColor Green
    } else {
        $nonExistingMailboxes += $mailbox
        Write-Host "Mailbox does not exist: $mailbox" -ForegroundColor Red
    }
}

# Output results
Write-Host "`nExisting Mailboxes:" -ForegroundColor Cyan
$existingMailboxes

Write-Host "`nNon-Existing Mailboxes:" -ForegroundColor Cyan
$nonExistingMailboxes