# =============================================================================
# Script: Diagnose-MailboxFolderAssistant.ps1
# Created: 2024-02-20 17:15:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2024-02-20 17:15:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0
# Additional Info: Initial script creation with diagnostic capabilities
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
    .\Diagnose-MailboxFolderAssistant.ps1
    Analyzes the folder assistant settings for the specified mailbox
#>

$mailbox = "leadership@leadershipspokane.org"

[xml]$diag = (Export-MailboxDiagnosticLogs $mailbox -ExtendedProperties).MailboxLog
$diag.Properties.MailboxTable.Property | Where-Object {$_.Name -like "ELC*"}

Export-MailboxDiagnosticLogs $mailbox -ComponentName MRM
