# Script to delete multiple mailboxes from a list in a text file

# Path to the text file containing the list of users
$userListPath = "C:\Temp\UserList.txt"

# Function to check if Exchange Management Shell is loaded
function Test-ExchangeShell {
    if (!(Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
        Write-Host "Exchange Management Shell is not loaded. Please run this script in Exchange Management Shell." -ForegroundColor Red
        return $false
    }
    return $true
}

# Check if Exchange Management Shell is loaded
if (!(Test-ExchangeShell)) {
    exit
}

# Check if the file exists
if (!(Test-Path $userListPath)) {
    Write-Host "The specified file does not exist: $userListPath" -ForegroundColor Red
    exit
}

# Read the list of users from the file
$users = Get-Content $userListPath

# Counter for successful and failed deletions
$successCount = 0
$failCount = 0

# Process each user in the list
foreach ($user in $users) {
    try {
        # Attempt to remove the mailbox
        Remove-Mailbox -Identity $user -Confirm:$false -ErrorAction Stop
        Write-Host "Successfully deleted mailbox for: $user" -ForegroundColor Green
        $successCount++
    }
    catch {
        Write-Host "Failed to delete mailbox for: $user" -ForegroundColor Red
        Write-Host "Error: $_" -ForegroundColor Red
        $failCount++
    }
}

# Display summary
Write-Host "`nDeletion Summary:" -ForegroundColor Cyan
Write-Host "Successfully deleted: $successCount mailbox(es)" -ForegroundColor Green
Write-Host "Failed to delete: $failCount mailbox(es)" -ForegroundColor Red