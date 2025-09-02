# =============================================================================
# Script: Get-FullMailboxAttributes.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-16 21:16:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2.4
# Additional Info: Fixed PSScriptAnalyzer CI/CD issues and implemented Write-ColorOutput function
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
    .\Get-FullMailboxAttributes.ps1 -Mailboxes "user1@domain.com", "user2@domain.com"
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

[CmdletBinding(DefaultParameterSetName = 'File')]
param(
    [Parameter(ParameterSetName = 'File')]
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

    [Parameter(ParameterSetName = 'Direct')]
    [string[]]$Mailboxes
)

# Initialize color support and logging
$Script:UseAnsiColors = $PSVersionTable.PSVersion.Major -ge 7
$Script:Colors = if ($Script:UseAnsiColors) {
    @{
        Red       = "`e[31m"
        Green     = "`e[32m"
        Yellow    = "`e[33m"
        Cyan      = "`e[36m"
        White     = "`e[37m"
        Magenta   = "`e[35m"
        DarkGray  = "`e[90m"
        Reset     = "`e[0m"
    }
} else {
    @{
        Red       = "Red"
        Green     = "Green"
        Yellow    = "Yellow"
        Cyan      = "Cyan"
        White     = "White"
        Magenta   = "Magenta"
        DarkGray  = "DarkGray"
        Reset     = ""
    }
}

function Write-ColorOutput {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$Color = "White"
    )

    if ($Script:UseAnsiColors) {
        # PowerShell 7+ with ANSI escape codes
        $colorCode = $Script:Colors[$Color]
        $resetCode  = $Script:Colors.Reset
        Write-Output "${colorCode}${Message}${resetCode}"
    } else {
        # PowerShell 5.1 - Change console color, write output, then reset
        $originalColor = $Host.UI.RawUI.ForegroundColor
        try {
            if ($Script:Colors[$Color] -and $Script:Colors[$Color] -ne "") {
                $Host.UI.RawUI.ForegroundColor = $Script:Colors[$Color]
            }
            Write-Output $Message
        } finally {
            $Host.UI.RawUI.ForegroundColor = $originalColor
        }
    }
}

# Initialize logging
$LogFile = Join-Path -Path $OutputPath -ChildPath "MailboxAttributes_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-ScriptLog {
    param($Message, $Level = "Information")

    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC"
    $LogMessage = "$TimeStamp [$Level] $Message"
    Add-Content -Path $LogFile -Value $LogMessage

    switch ($Level) {
        "Information" { Write-ColorOutput -Message $Message -Color "White" }
        "Success" { Write-ColorOutput -Message $Message -Color "Green" }
        "Warning" { Write-ColorOutput -Message $Message -Color "Yellow" }
        "Error" { Write-ColorOutput -Message $Message -Color "Red" }
        "Process" { Write-ColorOutput -Message $Message -Color "Cyan" }
    }
}

function Test-ExchangeConnection {
    try {
        $null = Get-OrganizationConfig -ErrorAction Stop
        Write-ScriptLog -Message "Successfully connected to Exchange Online" -Level "Success"
        return $true
    } catch {
        Write-ScriptLog -Message "Not connected to Exchange Online. Please run Connect-ExchangeOnline first." -Level "Error"
        return $false
    }
}

try {
    Write-ScriptLog -Message "Starting mailbox attribute collection..." -Level "Process"

    # Verify Exchange Online connection
    if (-not (Test-ExchangeConnection)) {
        throw "Exchange Online connection required"
    }

    # Get mailbox list
    if ($PSCmdlet.ParameterSetName -eq 'Direct') {
        $processMailboxes = $Mailboxes
    } else {
        if (-not (Test-Path -Path $InputPath)) {
            throw "Input file not found: $InputPath"
        }
        $processMailboxes = Get-Content -Path $InputPath
    }

    $totalMailboxes = $processMailboxes.Count
    Write-ScriptLog -Message "Found $totalMailboxes mailboxes to process" -Level "Process"
    $processed = 0

    foreach ($mailbox in $processMailboxes) {
        $processed++
        $percent = [math]::Round(($processed / $totalMailboxes) * 100)
        Write-Progress -Activity "Processing Mailboxes" -Status "$mailbox ($processed of $totalMailboxes)" -PercentComplete $percent

        try {
            Write-ScriptLog -Message "Processing mailbox: $mailbox" -Level "Process"
            $attributes = Get-Mailbox -Identity $mailbox -ErrorAction Stop | Select-Object *
            $outputFile = Join-Path -Path $OutputPath -ChildPath "$($mailbox -replace '[@\\/:*?"<>|]', '_')_attributes.txt"
            $attributes | Out-File -FilePath $outputFile -Force
            Write-ScriptLog -Message "Created attribute file: $(Split-Path -Path $outputFile -Leaf)" -Level "Success"
        } catch {
            Write-ScriptLog -Message "Error processing $mailbox`: $_" -Level "Error"
        }
    }
} catch {
    Write-ScriptLog -Message "Script execution failed: $_" -Level "Error"
    exit 1
} finally {
    Write-Progress -Activity "Processing Mailboxes" -Completed
    Write-ScriptLog -Message "Script execution completed. See log file for details: $LogFile" -Level "Process"
}
