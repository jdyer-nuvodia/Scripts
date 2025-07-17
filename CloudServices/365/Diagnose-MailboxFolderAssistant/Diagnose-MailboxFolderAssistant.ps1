# =============================================================================
# Script: Diagnose-MailboxFolderAssistant.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.3.1
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
# =============================================================================

<#
.SYNOPSIS
    Diagnoses Managed Folder Assistant settings and logs for a mailbox.
.DESCRIPTION
    This script exports and analyzes diagnostic logs related to the Managed Folder Assistant
    for a specified mailbox. It performs comprehensive analysis of:
    - ELC (Enterprise Lifecycle) properties
    - MRM (Messaging Records Management) components
    - Retention policy settings
    - Folder assistant processing status

    Key features:
    - Exports mailbox diagnostic logs with extended properties
    - Filters for ELC-related properties
    - Exports MRM component specific logs
    - Validates Exchange Online connectivity
    - Provides formatted, color-coded output
    - Handles errors gracefully

    Dependencies:
    - Exchange Online PowerShell Module (ExchangeOnlineManagement)
    - Active Exchange Online connection
    - Exchange Administrator or Global Reader role

    The script creates detailed logs that help diagnose issues with:
    - Retention policies not being applied
    - Folder assistant processing delays
    - MRM configuration problems
    - Policy inheritance issues
.PARAMETER Mailbox
    The email address of the mailbox to diagnose. Must be a valid Exchange Online mailbox.
    Accepts either email format (user@domain.com) or distinguished name.
.EXAMPLE
    .\Diagnose-MailboxFolderAssistant.ps1 -Mailbox "user@contoso.com"
    Analyzes the folder assistant settings for the specified mailbox
.EXAMPLE
    .\Diagnose-MailboxFolderAssistant.ps1 -Mailbox "CN = John Smith, OU = Users, DC = contoso, DC = com"
    Analyzes the folder assistant using distinguished name format
.NOTES
    Security Level: Medium
    Required Permissions: Exchange Administrator or Global Reader role
    Validation Requirements:
    - Verify Exchange Online connectivity
    - Verify mailbox exists
    - Verify appropriate permissions
    - Verify ExchangeOnlineManagement module is installed
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true,
        Position = 0,
        HelpMessage = "Enter the email address or distinguished name of the mailbox")]
    [ValidateNotNullOrEmpty()]
    [string]$Mailbox
)

# Function to test Exchange Online connectivity
function Test-ExchangeOnlineConnection {
    try {
        $null = Get-ConnectionInformation -ErrorAction Stop
        return $true
    } catch {
        Write-Error "Not connected to Exchange Online. Please run Connect-ExchangeOnline first."
        return $false
    }
}

# Main script execution
try {
    Write-Information "Starting mailbox folder assistant diagnostics..." -InformationAction Continue

    if (-not (Test-ExchangeOnlineConnection)) {
        exit 1
    }

    Write-Information "Analyzing mailbox: $Mailbox" -InformationAction Continue

    # Export and analyze diagnostic logs
    [xml]$diag = (Export-MailboxDiagnosticLogs $Mailbox -ExtendedProperties -ErrorAction Stop).MailboxLog

    Write-Information "`nELC Properties:" -InformationAction Continue
    $elcProperties = $diag.Properties.MailboxTable.Property | Where-Object { $_.Name -like "ELC*" } |
        Select-Object @{ N = 'Property'; E = { $_.Name } }, @{ N = 'Value'; E = { $_.Value } }
    $elcProperties | Format-Table -AutoSize

    Write-Information "`nExporting MRM diagnostic logs..." -InformationAction Continue
    $mrmLogs = Export-MailboxDiagnosticLogs $Mailbox -ComponentName MRM -ErrorAction Stop
    $mrmLogs | Format-List

    Write-Information "`nDiagnostic analysis completed successfully." -InformationAction Continue
} catch {
    Write-Error "An error occurred during diagnostic analysis: $_"
    exit 1
}
