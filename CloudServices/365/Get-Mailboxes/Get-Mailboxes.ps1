# =============================================================================
# Script: Get-Mailboxes.ps1
# Created: 2025-02-05 21:58:42 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-21 17:15:00 UTC
# Updated By: jdyer-nuvodia
# Version: 2.2
# Additional Info: Added default names.txt file handling
# =============================================================================

<#
.SYNOPSIS
    Checks existence of mailboxes in Microsoft 365 and provides detailed information
.DESCRIPTION
    This script verifies the existence of mailboxes from an input list and generates
    detailed information about each mailbox including:
     - Primary SMTP address
     - Display name
     - Mailbox type
     - Existence status
     
    Dependencies:
     - Exchange Online PowerShell Module
     - Active Exchange Online connection
     
    The script provides color-coded output for better visibility and exports
    detailed results to a CSV file if specified.
.PARAMETER InputFile
    (Optional) Path to the text file containing list of mailboxes to check (one per line)
    Default: names.txt in the same directory as the script
.PARAMETER OutputFile
    (Optional) Path to export the detailed results in CSV format
.EXAMPLE
    .\Get-Mailboxes.ps1 -InputFile ".\mailboxes.txt" -OutputFile ".\results.csv"
    Checks mailboxes listed in mailboxes.txt and exports results to results.csv
.NOTES
    Security Level: Medium
    Required Permissions: Exchange Online View-Only Recipients role
    Validation Requirements: 
     - Verify Exchange Online connection
     - Validate input file exists and is readable
     - Ensure write permissions if using OutputFile parameter
#>

# Parameters
param(
    [Parameter(Mandatory=$false)]
    [string]$InputFile = (Join-Path $PSScriptRoot "names.txt"),
    
    [Parameter(Mandatory=$false)]
    [string]$OutputFile
)

# Initialize arrays to store results
$existingMailboxes = @()
$nonExistingMailboxes = @()

# Import the list of mailboxes to check
try {
    if (-not (Test-Path -Path $InputFile)) {
        Write-Host "Warning: Input file not found at: $InputFile" -ForegroundColor Yellow
        if ($InputFile -ne (Join-Path $PSScriptRoot "names.txt")) {
            $defaultFile = Join-Path $PSScriptRoot "names.txt"
            if (Test-Path -Path $defaultFile) {
                Write-Host "Using default names.txt file instead" -ForegroundColor Cyan
                $InputFile = $defaultFile
            } else {
                Write-Error "Neither specified input file nor default names.txt exists"
                exit 1
            }
        } else {
            Write-Error "Default names.txt file not found in script directory"
            exit 1
        }
    }
    $mailboxList = Get-Content -Path $InputFile -ErrorAction Stop
    Write-Host "Successfully loaded $($mailboxList.Count) mailboxes from $InputFile" -ForegroundColor Cyan
} catch {
    Write-Error "Error loading input file: $_"
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
