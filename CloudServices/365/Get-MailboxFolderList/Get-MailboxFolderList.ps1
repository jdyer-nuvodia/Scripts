# =============================================================================
# Script: Get-MailboxFolderList.ps1
# Created: 2024-02-20 17:15:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2024-03-24 20:18:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1.1
# Additional Info: Updated CSV export path to script directory
# =============================================================================

<#
.SYNOPSIS
    Gets a list of mailbox folders and exports them to CSV.
.DESCRIPTION
    This script retrieves all folders from a specified mailbox and exports the 
    folder statistics to a CSV file.
.PARAMETER MailboxName
    The email address of the mailbox to analyze
.EXAMPLE
    .\Get-MailboxFolderList.ps1 -MailboxName "user@domain.com"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$MailboxName
)

# Specify the folder name you're searching for (use * for wildcard)
$FolderNameSearch = "**"

Write-Host "Starting mailbox folder analysis for: $MailboxName" -ForegroundColor Cyan

try {
    # Get all folders in the mailbox
    $Folders = Get-MailboxFolderStatistics -Identity $MailboxName | 
               Where-Object {$_.StartPath -like $FolderNameSearch}

    # Get script directory for CSV export
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $ExportPath = Join-Path $ScriptDir "MailboxFolders.csv"

    # Get all folders in the mailbox and export to CSV
    $Folders | 
        Select-Object StartPath, FolderType, ItemsInFolder, FolderSize |
        Export-Csv -Path $ExportPath -NoTypeInformation

    Write-Host "Export completed successfully" -ForegroundColor Green
} catch {
    Write-Error "Failed to process mailbox: $_"
}
