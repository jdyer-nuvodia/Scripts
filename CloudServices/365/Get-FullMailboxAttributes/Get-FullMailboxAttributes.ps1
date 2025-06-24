# =============================================================================
# Script: Get-FullMailboxAttributes.ps1
# Created: 2024-02-20 17:15:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-06-24 20:56:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2.1
# Additional Info: Implemented Write-Information for PSScriptAnalyzer compliance
# =============================================================================

<#
.SYNOPSIS
    Retrieves all attributes for specified mailboxes and exports them to individual text files.
.DESCRIPTION
    This script performs comprehensive mailbox attribute collection from Exchange Online:
    - Retrieves all available mailbox properties
    - Formats output in readable format
    - Creates individual files for each mailbox
    - Tracks progress with status indicators
    - Validates input and Exchange connection

    Key Features:
    - Flexible input options (file or direct mailbox list)
    - Customizable output location
    - Progress tracking and logging
    - Error handling and validation
    - Color-coded status output

    Dependencies:
    - Exchange Online PowerShell Module (ExchangeOnlineManagement)
    - Active Exchange Online connection
    - Exchange View-Only Recipients role or higher
    - Access to specified output directory

    The script creates detailed attribute files that include:
    - Basic mailbox properties
    - Custom attributes
    - Forwarding settings
    - Resource configurations
    - Retention settings
    - Security properties
.PARAMETER InputPath
    Optional. Path to a text file containing mailbox identifiers (one per line).
    If not specified, reads from 'mailboxes.txt' in script directory.
.PARAMETER OutputPath
    Optional. Directory where attribute files will be created.
    Defaults to script directory if not specified.
.PARAMETER Mailboxes
    Optional. Array of mailbox identifiers to process.
    Takes precedence over InputPath if both are specified.
.EXAMPLE
    .\Get-FullMailboxAttributes.ps1
    Processes mailboxes listed in mailboxes.txt in script directory
.EXAMPLE
    .\Get-FullMailboxAttributes.ps1 -InputPath "C:\Data\mailboxes.txt" -OutputPath "C:\Reports"
    Processes mailboxes from specified file and saves reports to custom location
.EXAMPLE
    .\Get-FullMailboxAttributes.ps1 -Mailboxes "user1@domain.com","user2@domain.com"
    Processes specified mailboxes directly without input file
.NOTES
    Security Level: Medium
    Required Permissions: Exchange View-Only Recipients role or higher
    Validation Requirements:
    - Verify Exchange Online connectivity
    - Verify input file exists (if specified)
    - Verify write access to output directory
    - Validate mailbox existence before processing
    - Verify ExchangeOnlineManagement module is installed
#>

[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(ParameterSetName='File')]
    [ValidateScript({
        if ($_) { Test-Path -Path $_ }
        else { $true }
    })]
    [string]$InputPath = (Join-Path -Path $PSScriptRoot -ChildPath "mailboxes.txt"),

    [Parameter()]
    [ValidateScript({
        if (-not (Test-Path $_)) {
            New-Item -Path $_ -ItemType Directory -Force | Out-Null
        }
        return $true
    })]
    [string]$OutputPath = $PSScriptRoot,

    [Parameter(ParameterSetName='Direct')]
    [string[]]$Mailboxes
)

# Initialize logging
$LogFile = Join-Path -Path $OutputPath -ChildPath "MailboxAttributes_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param($Message, $Level = "Information")

    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC"
    $LogMessage = "$TimeStamp [$Level] $Message"
    Add-Content -Path $LogFile -Value $LogMessage

    switch ($Level) {
        "Information" { Write-Information -MessageData $Message -InformationAction Continue }
        "Success" { Write-Information -MessageData $Message -InformationAction Continue }
        "Warning" { Write-Warning -Message $Message }
        "Error" { Write-Error -Message $Message }
        "Process" { Write-Information -MessageData $Message -InformationAction Continue }
    }
}

function Test-ExchangeConnection {
    try {
        $null = Get-OrganizationConfig -ErrorAction Stop
        Write-Log -Message "Successfully connected to Exchange Online" -Level "Success"
        return $true
    }
    catch {
        Write-Log -Message "Not connected to Exchange Online. Please run Connect-ExchangeOnline first." -Level "Error"
        return $false
    }
}

try {
    Write-Log -Message "Starting mailbox attribute collection..." -Level "Process"

    # Verify Exchange Online connection
    if (-not (Test-ExchangeConnection)) {
        throw "Exchange Online connection required"
    }

    # Get mailbox list
    if ($PSCmdlet.ParameterSetName -eq 'Direct') {
        $processMailboxes = $Mailboxes
    }
    else {
        if (-not (Test-Path -Path $InputPath)) {
            throw "Input file not found: $InputPath"
        }
        $processMailboxes = Get-Content -Path $InputPath
    }

    $totalMailboxes = $processMailboxes.Count
    Write-Log -Message "Found $totalMailboxes mailboxes to process" -Level "Process"
    $processed = 0

    foreach ($mailbox in $processMailboxes) {
        $processed++
        $percent = [math]::Round(($processed / $totalMailboxes) * 100)
        Write-Progress -Activity "Processing Mailboxes" -Status "$mailbox ($processed of $totalMailboxes)" -PercentComplete $percent

        try {
            Write-Log -Message "Processing mailbox: $mailbox" -Level "Process"
            $attributes = Get-Mailbox -Identity $mailbox -ErrorAction Stop | Select-Object *
            $outputFile = Join-Path -Path $OutputPath -ChildPath "$($mailbox -replace '[@\\/:*?"<>|]', '_')_attributes.txt"
            $attributes | Out-File -FilePath $outputFile -Force
            Write-Log -Message "Created attribute file: $(Split-Path -Path $outputFile -Leaf)" -Level "Success"
        }
        catch {
            Write-Log -Message "Error processing $mailbox`: $_" -Level "Error"
        }
    }
}
catch {
    Write-Log -Message "Script execution failed: $_" -Level "Error"
    exit 1
}
finally {
    Write-Progress -Activity "Processing Mailboxes" -Completed
    Write-Log -Message "Script execution completed. See log file for details: $LogFile" -Level "Process"
}
