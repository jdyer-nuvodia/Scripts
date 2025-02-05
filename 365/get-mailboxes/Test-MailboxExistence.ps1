# Script: get-mailboxesexist.ps1
# Version: 2.0
# Description: Checks existence of mailboxes and provides detailed information
# Author: jdyer-nuvodia
# Last Modified: 2025-02-05 21:58:42

# Parameters
param(
    [Parameter(Mandatory=$true)]
    [string]$InputFile,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputFile
)

# Initialize arrays to store results
$existingMailboxes = @()
$nonExistingMailboxes = @()

# Import the list of mailboxes to check
try {
    $mailboxList = Get-Content -Path $InputFile -ErrorAction Stop
    Write-Host "Successfully loaded $($mailboxList.Count) mailboxes from $InputFile" -ForegroundColor Cyan
} catch {
    Write-Host "Error loading input file: $_" -ForegroundColor Red
    exit 1
}

# Create results array for detailed information
$results = @()

# Loop through each mailbox in the list
foreach ($mailbox in $mailboxList) {
    try {
        $mbx = Get-Mailbox -Identity $mailbox -ErrorAction SilentlyContinue
        if ($mbx) {
            $existingMailboxes += $mailbox
            $results += [PSCustomObject]@{
                Mailbox = $mailbox
                Exists = $true
                PrimarySmtpAddress = $mbx.PrimarySmtpAddress
                DisplayName = $mbx.DisplayName
                MailboxType = $mbx.RecipientTypeDetails
                Status = "Found"
            }
            Write-Host "Mailbox exists: $mailbox ($($mbx.PrimarySmtpAddress))" -ForegroundColor Green
        } else {
            $nonExistingMailboxes += $mailbox
            $results += [PSCustomObject]@{
                Mailbox = $mailbox
                Exists = $false
                PrimarySmtpAddress = $null
                DisplayName = $null
                MailboxType = $null
                Status = "Not Found"
            }
            Write-Host "Mailbox does not exist: $mailbox" -ForegroundColor Red
        }
    } catch {
        Write-Host "Error processing mailbox $mailbox : $_" -ForegroundColor Red
    }
}

# Output summary
Write-Host "`n=== Summary ===" -ForegroundColor Cyan
Write-Host "Total mailboxes checked: $($mailboxList.Count)" -ForegroundColor White
Write-Host "Existing mailboxes: $($existingMailboxes.Count)" -ForegroundColor Green
Write-Host "Non-existing mailboxes: $($nonExistingMailboxes.Count)" -ForegroundColor Red

# Export results if output file specified
if ($OutputFile) {
    try {
        $results | Export-Csv -Path $OutputFile -NoTypeInformation
        Write-Host "`nDetailed results exported to: $OutputFile" -ForegroundColor Cyan
    } catch {
        Write-Host "Error exporting results: $_" -ForegroundColor Red
    }
}

# Display detailed results
Write-Host "`n=== Existing Mailboxes ===" -ForegroundColor Green
$results | Where-Object { $_.Exists } | Format-Table Mailbox, PrimarySmtpAddress, MailboxType -AutoSize

Write-Host "`n=== Non-Existing Mailboxes ===" -ForegroundColor Red
$nonExistingMailboxes | Format-Table -AutoSize