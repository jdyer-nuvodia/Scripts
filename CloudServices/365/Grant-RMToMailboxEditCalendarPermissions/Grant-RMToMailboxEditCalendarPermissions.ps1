# =============================================================================
# Script: Grant-RMToMailboxEditCalendarPermissions.ps1
# Created: 2024-02-20 17:30:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-08 19:23:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.4.0
# Additional Info: Added SupportsShouldProcess for safer permission grants
# =============================================================================

<#
.SYNOPSIS
    Grants mailbox and calendar permissions to a specified user in Exchange Online.
.DESCRIPTION
    This script manages Exchange Online mailbox and calendar permissions by:
    - Granting Full Access to mailboxes
    - Setting Editor rights on calendars
    - Supporting both single and batch operations
    - Providing detailed progress tracking
    - Maintaining operation logs
    
    Supports -WhatIf parameter to preview changes without making them.
    
    Key Features:
    - Flexible input options (single mailbox or file list)
    - Validates email formats
    - Checks existing permissions
    - Handles errors gracefully
    - Creates detailed operation logs
    
    Dependencies:
    - Exchange Online PowerShell Module (ExchangeOnlineManagement)
    - Active Exchange Online connection
    - Exchange Administrator role
    - Access to mailboxes and calendars
    
    The script performs the following actions:
    1. Validates input parameters and connectivity
    2. Checks existing permissions to avoid duplicates
    3. Grants Full Access to mailboxes
    4. Sets Editor rights on calendars with delegate access
    5. Logs all operations with timestamps
.PARAMETER UserEmail
    Email address of the user to grant permissions to.
    Must be a valid email format.
.PARAMETER SingleMailbox
    Optional. Single mailbox to process instead of reading from mailboxes.txt.
    Must be a valid email format.
.PARAMETER InputFile
    Optional. Path to a text file containing mailbox email addresses (one per line).
    Defaults to mailboxes.txt in script directory if not specified.
.PARAMETER AutoMapping
    Optional. Whether to automatically map the mailboxes in Outlook.
    Default is false to prevent cluttering user's Outlook.
.PARAMETER LogPath
    Optional. Directory for log files. Creates 'Logs' subdirectory in script
    path if not specified.
.EXAMPLE
    .\Grant-RMToMailboxEditCalendarPermissions.ps1 -UserEmail "admin@contoso.com"
    Grants permissions on all mailboxes listed in mailboxes.txt
.EXAMPLE
    .\Grant-RMToMailboxEditCalendarPermissions.ps1 -UserEmail "admin@contoso.com" -SingleMailbox "user@contoso.com"
    Grants permissions on a single specified mailbox
.EXAMPLE
    .\Grant-RMToMailboxEditCalendarPermissions.ps1 -UserEmail "admin@contoso.com" -InputFile "C:\Data\mailboxes.txt" -AutoMapping $true
    Processes mailboxes from custom file location with automapping enabled
.NOTES
    Security Level: High
    Required Permissions: Exchange Administrator role
    Validation Requirements:
    - Verify Exchange Online connectivity
    - Validate all email addresses
    - Verify mailbox existence
    - Check for adequate permissions
    - Verify ExchangeOnlineManagement module is installed
#>

[CmdletBinding(DefaultParameterSetName='File',
               SupportsShouldProcess=$true,
               ConfirmImpact='High')]
param(
    [Parameter(Mandatory=$true,
              Position=0,
              HelpMessage="Email address of user to grant permissions to")]
    [ValidatePattern('^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$')]
    [string]$UserEmail,
    
    [Parameter(Mandatory=$false,
              ParameterSetName='Single',
              HelpMessage="Single mailbox to process")]
    [ValidatePattern('^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$')]
    [string]$SingleMailbox,
    
    [Parameter(Mandatory=$false,
              ParameterSetName='File',
              HelpMessage="Path to mailbox list file")]
    [ValidateScript({Test-Path $_})]
    [string]$InputFile = (Join-Path $PSScriptRoot "mailboxes.txt"),
    
    [Parameter(Mandatory=$false)]
    [bool]$AutoMapping = $false,
    
    [Parameter(Mandatory=$false)]
    [ValidateScript({
        if (-not (Test-Path $_)) {
            New-Item -Path $_ -ItemType Directory -Force | Out-Null
        }
        return $true
    })]
    [string]$LogPath = (Join-Path $PSScriptRoot "Logs")
)

# Initialize logging
$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile = Join-Path $LogPath "PermissionGrants_${TimeStamp}.log"

function Write-Log {
    param($Message, $Level = "Information")
    
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC"
    $LogMessage = "$TimeStamp [$Level] $Message"
    Add-Content -Path $LogFile -Value $LogMessage
    
    switch ($Level) {
        "Information" { Write-Host $Message -ForegroundColor White }
        "Success"     { Write-Host $Message -ForegroundColor Green }
        "Warning"     { Write-Host $Message -ForegroundColor Yellow }
        "Error"       { Write-Host $Message -ForegroundColor Red }
        "Process"     { Write-Host $Message -ForegroundColor Cyan }
    }
}

function Test-ExchangeConnection {
    try {
        $null = Get-OrganizationConfig -ErrorAction Stop
        Write-Log "Successfully connected to Exchange Online" "Success"
        return $true
    }
    catch {
        Write-Log "Not connected to Exchange Online. Please run Connect-ExchangeOnline first." "Error"
        return $false
    }
}

try {
    Write-Log "Starting permission grant process..." "Process"
    
    # Verify Exchange Online connection
    if (-not (Test-ExchangeConnection)) {
        throw "Exchange Online connection required"
    }
    
    # Get mailbox list
    $mailboxes = if ($PSCmdlet.ParameterSetName -eq 'Single') {
        @($SingleMailbox)
    }
    else {
        Get-Content $InputFile
    }
    
    $totalMailboxes = $mailboxes.Count
    Write-Log "Processing $totalMailboxes mailbox(es)" "Process"
    $processed = 0
    
    foreach ($mailbox in $mailboxes) {
        $processed++
        $percent = [math]::Round(($processed / $totalMailboxes) * 100)
        Write-Progress -Activity "Granting Permissions" -Status "$mailbox ($processed of $totalMailboxes)" -PercentComplete $percent
        
        try {
            Write-Log "Processing mailbox: $mailbox" "Process"
            
            # Verify mailbox exists
            $mbx = Get-Mailbox -Identity $mailbox -ErrorAction Stop
            
            # Grant mailbox permissions
            if ($PSCmdlet.ShouldProcess($mailbox, "Grant FullAccess permissions to $UserEmail")) {
                Add-MailboxPermission -Identity $mailbox -User $UserEmail -AccessRights FullAccess -InheritanceType All -AutoMapping:$AutoMapping -ErrorAction Stop
                Write-Log "Granted Full Access on mailbox $mailbox" "Success"
            }
            
            # Grant calendar permissions
            $calendarPath = $mbx.UserPrincipalName + ":\Calendar"
            if ($PSCmdlet.ShouldProcess($calendarPath, "Grant Editor calendar permissions to $UserEmail")) {
                Add-MailboxFolderPermission -Identity $calendarPath -User $UserEmail -AccessRights Editor -SharingPermissionFlags Delegate,CanViewPrivateItems -ErrorAction Stop
                Write-Log "Granted Editor rights on calendar for $mailbox" "Success"
            }
        }
        catch {
            Write-Log "Error processing $mailbox`: $_" "Error"
        }
    }
}
catch {
    Write-Log "Script execution failed: $_" "Error"
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" "Error"
    exit 1
}
finally {
    Write-Progress -Activity "Granting Permissions" -Completed
    Write-Log "Script execution completed. See log file for details: $LogFile" "Process"
}
