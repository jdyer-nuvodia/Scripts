# =============================================================================
# Script: Get-MailboxFolderList.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.3.1
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
# =============================================================================

<#
.SYNOPSIS
    Gets a list of mailbox folders and exports them to CSV with detailed statistics.
.DESCRIPTION
    This script retrieves all folders from a specified Microsoft 365 mailbox and exports
    detailed folder statistics to a CSV file. It provides information about:
    - Folder paths and hierarchy
    - Folder types (Calendar, Mail, Contacts, etc.)
    - Item counts per folder
    - Folder sizes

    Dependencies:
    - Exchange Online PowerShell Module (ExchangeOnlineManagement)
    - Active Exchange Online connection
    - Appropriate mailbox permissions

    The script performs the following actions:
    1. Validates Exchange Online connectivity
    2. Retrieves mailbox folder statistics
    3. Exports data to a timestamped CSV file
    4. Provides progress feedback and error handling
.PARAMETER MailboxName
    The email address or identity of the mailbox to analyze.
    Accepts either email format (user@domain.com) or distinguished name.
.PARAMETER FolderFilter
    Optional. Search pattern for filtering folders.
    Default is "**" for all folders.
    Use patterns like "*Inbox*" or "*Calendar*" to filter specific folders.
.PARAMETER ExportPath
    Optional. Path where the CSV file will be saved.
    Defaults to script directory if not specified.
.EXAMPLE
    .\Get-MailboxFolderList.ps1 -MailboxName "user@domain.com"
    Retrieves all folders from the specified mailbox and exports to CSV
.EXAMPLE
    .\Get-MailboxFolderList.ps1 -MailboxName "user@domain.com" -FolderFilter "*Calendar*"
    Retrieves only calendar-related folders from the mailbox
.EXAMPLE
    .\Get-MailboxFolderList.ps1 -MailboxName "user@domain.com" -ExportPath "C:\Reports"
    Exports the folder list to a specific directory
.NOTES
    Security Level: Medium
    Required Permissions: Exchange Online View-Only Recipients role or higher
    Validation Requirements:
    - Verify Exchange Online connectivity
    - Verify mailbox exists
    - Verify export path is accessible
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[^@]+@[^@]+\.[^@]+$|^[a-zA-Z]+/[^/]+$')]
    [string]$MailboxName,

    [Parameter(Mandatory = $false)]
    [string]$FolderFilter = "**",

    [Parameter(Mandatory = $false)]
    [ValidateScript({
            if ($_ -and !(Test-Path $_)) {
                New-Item -Path $_ -ItemType Directory -Force | Out-Null
            }
            return $true
        })]
    [string]$ExportPath = $PSScriptRoot
)

# Initialize logging
$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile = Join-Path $ExportPath "MailboxFolderList_$($MailboxName.Split('@')[0])_${ TimeStamp}.log"

function Write-ScriptLog {
    param($Message, $Level = "Information")

    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC"
    $LogMessage = "$TimeStamp [$Level] $Message"
    Add-Content -Path $LogFile -Value $LogMessage

    switch ($Level) {
        "Information" { Write-Output $Message }
        "Success" { Write-Output $Message }
        "Warning" { Write-Warning $Message }
        "Error" { Write-Error $Message }
        "Process" { Write-Output $Message }
    }
}

try {
    Write-ScriptLog "Starting mailbox folder analysis for: $MailboxName" "Process"

    # Verify Exchange Online connection
    try {
        $null = Get-EXOMailbox -Identity $MailboxName -ErrorAction Stop
        Write-ScriptLog "Successfully connected to Exchange Online" "Success"
    } catch {
        if ($_.Exception.Message -like "*Connect-ExchangeOnline*") {
            Write-ScriptLog "Not connected to Exchange Online. Please run Connect-ExchangeOnline first" "Error"
            throw "Exchange Online connection required"
        }
        if ($_.Exception.Message -like "*couldn't be found*") {
            Write-ScriptLog "Mailbox $MailboxName not found" "Error"
            throw "Mailbox not found"
        }
        throw
    }

    # Create export filename with timestamp
    $ExportFile = Join-Path $ExportPath "MailboxFolders_$($MailboxName.Split('@')[0])_${ TimeStamp}.csv"
    Write-ScriptLog "Export will be saved to: $ExportFile" "Information"

    # Get folder statistics
    Write-ScriptLog "Retrieving folder statistics..." "Process"
    $Folders = Get-MailboxFolderStatistics -Identity $MailboxName |
        Where-Object { $_.StartPath -like $FolderFilter }

    Write-ScriptLog "Found $($Folders.Count) folders matching filter" "Success"

    # Export to CSV with enhanced information
    $Folders |
        Select-Object @{ N = 'Mailbox'; E = { $MailboxName } },
        StartPath,
        FolderType,
        ItemsInFolder,
        @{ N = 'FolderSizeInMB'; E = { [math]::Round($_.FolderSize.ToMB(), 2) } },
        LastModifiedTime,
        ContentMailboxGuid,
        ContentMailboxServerName |
        Export-Csv -Path $ExportFile -NoTypeInformation

    Write-ScriptLog "Export completed successfully to: $ExportFile" "Success"
    Write-ScriptLog "Total folders exported: $($Folders.Count)" "Success"
} catch {
    Write-ScriptLog "Error processing mailbox: $_" "Error"
    Write-ScriptLog "Stack Trace: $($_.ScriptStackTrace)" "Error"
    throw
} finally {
    Write-ScriptLog "Script execution finished" "Information"
}
