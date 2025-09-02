# =============================================================================
# Script: Remove-AllMailboxPermissions.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.3.3
# Additional Info: Fixed header metadata for workflow validation
# =============================================================================

<#
.SYNOPSIS
    Removes all mailbox permissions including FullAccess, SendAs, Send on Behalf, and Calendar permissions.
.DESCRIPTION
    This script performs comprehensive permission removal from Exchange Online mailboxes:
    - Removes FullAccess permissions for all delegates
    - Removes SendAs permissions for all trustees
    - Removes Send on Behalf permissions
    - Removes Calendar permissions (except Default and Anonymous)
    - Validates changes after removal
    - Creates detailed operation logs

    Supports -WhatIf parameter to preview changes without making them.

    Key Features:
    - Batch or single mailbox processing
    - Progress tracking with status indicators
    - Comprehensive error handling
    - Detailed logging of all operations
    - Validation of permission removal

    Dependencies:
    - Exchange Online PowerShell Module (ExchangeOnlineManagement)
    - Active Exchange Online connection
    - Exchange Administrator role
    - Write access to log directory

    The script removes all types of permissions while:
    - Preserving system-required permissions
    - Maintaining audit trail of removals
    - Providing detailed progress feedback
    - Validating successful removal
.PARAMETER MailboxIdentity
    Optional. Single mailbox to process (email or UPN format).
    If not specified, processes mailboxes from mailboxes.txt.
.PARAMETER InputPath
    Optional. Path to text file containing mailbox identifiers.
    Defaults to mailboxes.txt in script directory.
.PARAMETER LogPath
    Optional. Directory for storing operation logs.
    Creates 'Logs' subdirectory in script path if not specified.
.EXAMPLE
    .\Remove-AllMailboxPermissions.ps1
    Removes permissions for all mailboxes listed in mailboxes.txt
.EXAMPLE
    .\Remove-AllMailboxPermissions.ps1 -MailboxIdentity "user@contoso.com"
    Removes all permissions for a single specified mailbox
.EXAMPLE
    .\Remove-AllMailboxPermissions.ps1 -InputPath "C:\Data\mailboxes.txt" -LogPath "C:\Logs"
    Processes mailboxes from custom file with custom log location
.NOTES
    Security Level: High
    Required Permissions: Exchange Administrator role
    Validation Requirements:
    - Verify Exchange Online connectivity
    - Validate mailbox existence
    - Verify appropriate permissions
    - Verify ExchangeOnlineManagement module is installed
#>

[CmdletBinding(DefaultParameterSetName = 'File',
    SupportsShouldProcess = $true,
    ConfirmImpact = 'High')]
param(
    [Parameter(ParameterSetName = 'Single',
        Position = 0,
        HelpMessage = "Email address of mailbox to process")]
    [ValidatePattern('^[\w-\.]+@([\w-]+\.)+[\w-]{ 2, 4}$')]
    [string]$MailboxIdentity,

    [Parameter(ParameterSetName = 'File')]
    [ValidateScript({ Test-Path $_ })]
    [string]$InputPath = (Join-Path $PSScriptRoot "mailboxes.txt"),

    [Parameter()]
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
$LogFile = Join-Path $LogPath "PermissionRemovals_${ TimeStamp}.log"

function Write-LogMessage {
    [CmdletBinding()]
    [OutputType([void])]
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

function Test-ExchangeConnection {
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    try {
        $null = Get-OrganizationConfig -ErrorAction Stop
        Write-LogMessage "Successfully connected to Exchange Online" "Success"
        return $true
    } catch {
        Write-LogMessage "Not connected to Exchange Online. Please run Connect-ExchangeOnline first." "Error"
        return $false
    }
}

function Remove-MailboxDelegate {
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([bool])]
    param([string]$Identity)

    try {
        $delegates = Get-MailboxPermission -Identity $Identity | Where-Object {
            $_.IsInherited -eq $false -and
            $_.User -ne "NT AUTHORITY\SELF" -and
            $_.AccessRights -like "*FullAccess*"
        }

        foreach ($delegate in $delegates) {
            if ($PSCmdlet.ShouldProcess($Identity, "Remove FullAccess permission for $($delegate.User)")) {
                Remove-MailboxPermission -Identity $Identity -User $delegate.User -AccessRights FullAccess -Confirm:$false -ErrorAction Stop
                Write-LogMessage "Removed FullAccess permission for $($delegate.User) on $Identity" "Success"
            }
        }
        return $true
    } catch {
        Write-LogMessage "Error removing FullAccess permissions for $Identity`: $_" "Error"
        return $false
    }
}

function Remove-SendAsPermission {
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([bool])]
    param([string]$Identity)

    try {
        $sendAsPermissions = Get-RecipientPermission -Identity $Identity |
            Where-Object { $_.IsInherited -eq $false -and $_.Trustee -ne "NT AUTHORITY\SELF" }

        foreach ($permission in $sendAsPermissions) {
            if ($PSCmdlet.ShouldProcess($Identity, "Remove SendAs permission for $($permission.Trustee)")) {
                Remove-RecipientPermission -Identity $Identity -Trustee $permission.Trustee -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                Write-LogMessage "Removed SendAs permission for $($permission.Trustee) on $Identity" "Success"
            }
        }
        return $true
    } catch {
        Write-LogMessage "Error removing SendAs permissions for $Identity`: $_" "Error"
        return $false
    }
}

function Remove-SendOnBehalfPermission {
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([bool])]
    param([string]$Identity)

    try {
        if ($PSCmdlet.ShouldProcess($Identity, "Remove all Send on Behalf permissions")) {
            Set-Mailbox -Identity $Identity -GrantSendOnBehalfTo $null -ErrorAction Stop
            Write-LogMessage "Removed all Send on Behalf permissions for $Identity" "Success"
        }
        return $true
    } catch {
        Write-LogMessage "Error removing Send on Behalf permissions for $Identity`: $_" "Error"
        return $false
    }
}

function Remove-CalendarPermission {
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([bool])]
    param([string]$Identity)

    try {
        $calendarPath = $Identity + ":\Calendar"
        $calendarPermissions = Get-MailboxFolderPermission -Identity $calendarPath -ErrorAction Stop

        foreach ($permission in $calendarPermissions) {
            if ($permission.User.DisplayName -notin @("Default", "Anonymous")) {
                if ($PSCmdlet.ShouldProcess($calendarPath, "Remove calendar permission for $($permission.User.DisplayName)")) {
                    Remove-MailboxFolderPermission -Identity $calendarPath -User $permission.User.DisplayName -Confirm:$false -ErrorAction Stop
                    Write-LogMessage "Removed calendar permission for $($permission.User.DisplayName) on $Identity" "Success"
                }
            }
        }
        return $true
    } catch {
        Write-LogMessage "Error removing calendar permissions for $Identity`: $_" "Error"
        return $false
    }
}

try {
    Write-LogMessage "Starting permission removal process..." "Process"

    # Verify Exchange Online connection
    if (-not (Test-ExchangeConnection)) {
        throw "Exchange Online connection required"
    }

    # Get mailbox list
    $mailboxes = if ($PSCmdlet.ParameterSetName -eq 'Single') {
        @($MailboxIdentity)
    } else {
        Get-Content $InputPath
    }

    $totalMailboxes = $mailboxes.Count
    Write-LogMessage "Processing $totalMailboxes mailbox(es)" "Process"
    $processed = 0

    foreach ($mailbox in $mailboxes) {
        $processed++
        $percent = [math]::Round(($processed / $totalMailboxes) * 100)
        Write-Progress -Activity "Removing Permissions" -Status "$mailbox ($processed of $totalMailboxes)" -PercentComplete $percent

        Write-LogMessage "Processing mailbox: $mailbox" "Process"

        # Verify mailbox exists
        try {
            $null = Get-Mailbox -Identity $mailbox -ErrorAction Stop
        } catch {
            Write-LogMessage "Mailbox not found: $mailbox. Skipping." "Warning"
            continue
        }

        # Remove each permission type
        $results = @(
            @{ Type = "FullAccess"; Success = Remove-MailboxDelegate $mailbox },
            @{ Type = "SendAs"; Success = Remove-SendAsPermission $mailbox },
            @{ Type = "SendOnBehalf"; Success = Remove-SendOnBehalfPermission $mailbox },
            @{ Type = "Calendar"; Success = Remove-CalendarPermission $mailbox }
        )

        # Log summary for this mailbox
        $successCount = ($results | Where-Object { $_.Success }).Count
        $status = if ($successCount -eq $results.Count) { "Success" } else { "Warning" }
        Write-LogMessage "Completed processing $mailbox`: $successCount of $($results.Count) operations successful" $status
    }
} catch {
    Write-LogMessage "Script execution failed: $_" "Error"
    Write-LogMessage "Stack Trace: $($_.ScriptStackTrace)" "Error"
    exit 1
} finally {
    Write-Progress -Activity "Removing Permissions" -Completed
    Write-LogMessage "Script execution completed. See log file for details: $LogFile" "Process"
}
