# =============================================================================
# Script: Get-Mailboxes.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 2.3.1
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
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
    [Parameter(Mandatory = $false)]
    [string]$InputFile = (Join-Path $PSScriptRoot "names.txt"),

    [Parameter(Mandatory = $false)]
    [string]$OutputFile
)

# Initialize arrays to store results
$existingMailboxes = @()
$nonExistingMailboxes = @()

# Import the list of mailboxes to check
try {
    if (-not (Test-Path -Path $InputFile)) {
        Write-Warning "Input file not found at: $InputFile"
        if ($InputFile -ne (Join-Path $PSScriptRoot "names.txt")) {
            $defaultFile = Join-Path $PSScriptRoot "names.txt"
            if (Test-Path -Path $defaultFile) {
                Write-Output "Using default names.txt file instead"
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
    Write-Output "Successfully loaded $($mailboxList.Count) mailboxes from $InputFile"
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
            Write-Output "Mailbox exists: $mailbox ($($mbx.PrimarySmtpAddress))"
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
            Write-Output "Mailbox does not exist: $mailbox"
        }
    } catch {
        Write-Error "Error processing mailbox $mailbox : $_"
    }
}

# Output summary
Write-Output "`n=== Summary ==="
Write-Output "Total mailboxes checked: $($mailboxList.Count)"
Write-Output "Existing mailboxes: $($existingMailboxes.Count)"
Write-Output "Non-existing mailboxes: $($nonExistingMailboxes.Count)"

# Export results if output file specified
if ($OutputFile) {
    try {
        $results | Export-Csv -Path $OutputFile -NoTypeInformation
        Write-Output "`nDetailed results exported to: $OutputFile"
    } catch {
        Write-Error "Error exporting results: $_"
    }
}

# Display detailed results
Write-Output "`n=== Existing Mailboxes ==="
$results | Where-Object { $_.Exists } | Format-Table Mailbox, PrimarySmtpAddress, MailboxType -AutoSize

Write-Output "`n=== Non-Existing Mailboxes ==="
$nonExistingMailboxes | Format-Table -AutoSize
