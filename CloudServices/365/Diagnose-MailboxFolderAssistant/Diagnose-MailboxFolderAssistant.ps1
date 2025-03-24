# =============================================================================
# Script: Diagnose-MailboxFolderAssistant.ps1
# Created: 2024-02-20 17:15:00 UTC
# Author: nunya-nunya
# Last Updated: 2024-02-20 17:30:00 UTC
# Updated By: nunya-nunya
# Version: 1.1
# Additional Info: Added error handling, parameter validation, and formatted output
# =============================================================================

<#
.SYNOPSIS
    Diagnoses Managed Folder Assistant settings and logs for a mailbox.
.DESCRIPTION
    This script exports and analyzes diagnostic logs related to the Managed Folder Assistant
    for a specified mailbox. It focuses on ELC (Enterprise Lifecycle) properties and MRM
    (Messaging Records Management) components.
    
    Key actions:
    - Exports mailbox diagnostic logs with extended properties
    - Filters for ELC-related properties
    - Exports MRM component specific logs
    
    Dependencies:
    - Exchange Online PowerShell Module
    - Appropriate Exchange admin permissions
.PARAMETER mailbox
    The email address of the mailbox to diagnose
.EXAMPLE
    .\Diagnose-MailboxFolderAssistant.ps1 -Mailbox "leadership@leadershipspokane.org"
    Analyzes the folder assistant settings for the specified mailbox
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Mailbox
)

# Function to test Exchange Online connectivity
function Test-ExchangeOnlineConnection {
    try {
        $null = Get-ConnectionInformation -ErrorAction Stop
        return $true
    }
    catch {
        Write-Error "Not connected to Exchange Online. Please run Connect-ExchangeOnline first."
        return $false
    }
}

# Main script execution
try {
    Write-Host "Starting mailbox folder assistant diagnostics..." -ForegroundColor Cyan

    if (-not (Test-ExchangeOnlineConnection)) {
        exit 1
    }

    Write-Host "Analyzing mailbox: $Mailbox" -ForegroundColor Cyan

    # Export and analyze diagnostic logs
    [xml]$diag = (Export-MailboxDiagnosticLogs $Mailbox -ExtendedProperties -ErrorAction Stop).MailboxLog
    
    Write-Host "`nELC Properties:" -ForegroundColor Cyan
    $elcProperties = $diag.Properties.MailboxTable.Property | Where-Object {$_.Name -like "ELC*"} | 
        Select-Object @{N='Property';E={$_.Name}}, @{N='Value';E={$_.Value}}
    $elcProperties | Format-Table -AutoSize

    Write-Host "`nExporting MRM diagnostic logs..." -ForegroundColor Cyan
    $mrmLogs = Export-MailboxDiagnosticLogs $Mailbox -ComponentName MRM -ErrorAction Stop
    $mrmLogs | Format-List

    Write-Host "`nDiagnostic analysis completed successfully." -ForegroundColor Green
}
catch {
    Write-Error "An error occurred during diagnostic analysis: $_"
    exit 1
}
