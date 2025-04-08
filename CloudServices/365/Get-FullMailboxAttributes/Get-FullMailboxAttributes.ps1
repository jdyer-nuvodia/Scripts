# =============================================================================
# Script: Get-FullMailboxAttributes.ps1
# Created: 2024-02-20 17:15:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-02 21:00:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2.0
# Additional Info: Enhanced documentation, added parameter support and progress tracking
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
        if ($_) { Test-Path $_ }
        else { $true }
    })]
    [string]$InputPath = (Join-Path $PSScriptRoot "mailboxes.txt"),

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
$LogFile = Join-Path $OutputPath "MailboxAttributes_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param($Message, $Level = "Information")
    
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC"
    $LogMessage = "$TimeStamp [$Level] $Message"
    Add-Content -Path $LogFile -Value $LogMessage
    
    switch ($Level) {
        "Information" { Write-Host $Message -ForegroundColor White }
        "Success" { Write-Host $Message -ForegroundColor Green }
        "Warning" { Write-Host $Message -ForegroundColor Yellow }
        "Error" { Write-Host $Message -ForegroundColor Red }
        "Process" { Write-Host $Message -ForegroundColor Cyan }
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
    Write-Log "Starting mailbox attribute collection..." "Process"
    
    # Verify Exchange Online connection
    if (-not (Test-ExchangeConnection)) {
        throw "Exchange Online connection required"
    }
    
    # Get mailbox list
    if ($PSCmdlet.ParameterSetName -eq 'Direct') {
        $processMailboxes = $Mailboxes
    }
    else {
        if (-not (Test-Path $InputPath)) {
            throw "Input file not found: $InputPath"
        }
        $processMailboxes = Get-Content $InputPath
    }
    
    $totalMailboxes = $processMailboxes.Count
    Write-Log "Found $totalMailboxes mailboxes to process" "Process"
    $processed = 0
    
    foreach ($mailbox in $processMailboxes) {
        $processed++
        $percent = [math]::Round(($processed / $totalMailboxes) * 100)
        Write-Progress -Activity "Processing Mailboxes" -Status "$mailbox ($processed of $totalMailboxes)" -PercentComplete $percent
        
        try {
            Write-Log "Processing mailbox: $mailbox" "Process"
            $attributes = Get-Mailbox -Identity $mailbox -ErrorAction Stop | Select-Object *
            $outputFile = Join-Path $OutputPath "$($mailbox -replace '[@\\/:*?"<>|]', '_')_attributes.txt"
            $attributes | Out-File $outputFile -Force
            Write-Log "Created attribute file: $(Split-Path $outputFile -Leaf)" "Success"
        }
        catch {
            Write-Log "Error processing $mailbox`: $_" "Error"
        }
    }
}
catch {
    Write-Log "Script execution failed: $_" "Error"
    exit 1
}
finally {
    Write-Progress -Activity "Processing Mailboxes" -Completed
    Write-Log "Script execution completed. See log file for details: $LogFile" "Process"
}
