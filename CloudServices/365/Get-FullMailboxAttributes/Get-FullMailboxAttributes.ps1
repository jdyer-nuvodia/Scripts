# =============================================================================
# Script: Get-FullMailboxAttributes.ps1
# Created: 2024-02-20 17:15:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2024-02-20 17:20:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1
# Additional Info: Updated to use relative paths and added error handling
# =============================================================================

<#
.SYNOPSIS
    Retrieves all attributes for specified mailboxes and exports them to individual text files.
.DESCRIPTION
    This script reads a list of mailboxes from a text file and retrieves all available
    attributes for each mailbox using Get-Mailbox cmdlet. The results are exported
    to individual text files named after each mailbox.
    
    Dependencies:
    - Exchange Online PowerShell module
    - Active Exchange Online connection
    - Appropriate permissions to view mailbox properties
.PARAMETER None
    No parameters required. Mailbox list is read from a fixed path.
.EXAMPLE
    .\Get-FullMailboxAttributes.ps1
    Processes all mailboxes listed in the mailboxes.txt file and creates individual attribute files.
.NOTES
    Security Level: Medium
    Required Permissions: Exchange View-Only Recipients role or higher
    Validation Requirements: Verify Exchange Online connection before running
#>

# Get the script's directory and setup paths
$scriptDir = $PSScriptRoot
$inputFile = Join-Path $scriptDir "mailboxes.txt"

# Verify input file exists
if (-not (Test-Path $inputFile)) {
    Write-Error "mailboxes.txt not found in script directory: $scriptDir"
    exit 1
}

# Read the list of mailboxes from a text file
$mailboxes = Get-Content $inputFile

foreach ($mailbox in $mailboxes) {
    Write-Host "Processing mailbox: $mailbox" -ForegroundColor Cyan
    $attributes = Get-Mailbox -Identity $mailbox | Select-Object *
    $outputFile = Join-Path $scriptDir "$mailbox`_attributes.txt"
    $attributes | Out-File $outputFile
    Write-Host "Created attribute file: $($mailbox)_attributes.txt" -ForegroundColor Green
}
