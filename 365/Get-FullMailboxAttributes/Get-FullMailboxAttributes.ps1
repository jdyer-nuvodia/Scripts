# Read the list of mailboxes from a text file
$mailboxes = Get-Content "C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\getFullMailboxAttributes\mailboxes.txt"


foreach ($mailbox in $mailboxes) {
    $attributes = Get-Mailbox -Identity $mailbox | Select-Object *
    $outputFile = "C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\getFullMailboxAttributes\$mailbox`_attributes.txt"
    $attributes | Out-File $outputFile
}