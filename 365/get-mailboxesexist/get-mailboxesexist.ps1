# Import the list of mailboxes to check
$userList = Get-Content -Path "C:\Temp\names.txt"

# Loop through each mailbox in the list
foreach ($user in $userList) {
    if (Get-Mailbox -Identity "$user" -ErrorAction SilentlyContinue) {
        Write-Host (Get-mailbox -Identity "$user" | Select-Object PrimarySmtpAddress) -ForegroundColor Green
    } else {
        Write-Host "Mailbox does not exist: $user" -ForegroundColor Red
    }
}
