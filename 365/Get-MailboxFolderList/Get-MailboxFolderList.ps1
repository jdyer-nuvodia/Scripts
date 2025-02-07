# Specify the mailbox you want to search
$MailboxName = "bwilcox@workwith.com"

# Specify the folder name you're searching for (use * for wildcard)
$FolderNameSearch = "**"

# Get all folders in the mailbox
$Folders = Get-MailboxFolderStatistics -Identity $MailboxName | 
           Where-Object {$_.FolderPath -like $FolderNameSearch}

# Get all folders in the mailbox and export to CSV
Get-MailboxFolderStatistics -Identity $MailboxName | 
    Where-Object {$_.FolderPath -like $FolderNameSearch} |
    Select-Object FolderPath, FolderType, ItemsInFolder, FolderSize |
    Export-Csv -Path "C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\getMailboxFolderList\MailboxFolders.csv" -NoTypeInformation