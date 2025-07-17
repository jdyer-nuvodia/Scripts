# =============================================================================
# Script: Manage-RDPSessions.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1.4
# Additional Info: Fixed header metadata for workflow validation
# =============================================================================

<#
.SYNOPSIS
    Lists and manages Remote Desktop (RDP) sessions on a server.

.DESCRIPTION
    This script provides functionality to list and optionally terminate inactive or disconnected
    Remote Desktop sessions on a specified server. It uses qwinsta (query session) and rwinsta
    (reset session) commands to perform these operations.

    The script can be run in report-only mode or termination mode with confirmation options.
    It includes functionality to filter sessions based on state and idle time.

.PARAMETER ComputerName
    The name of the remote server to manage RDP sessions on. Default is the local computer.

.PARAMETER State
    The session state to filter by. Valid options are: "Active", "Disconnected", "All". Default is "All".

.PARAMETER IdleTimeThreshold
    The threshold in minutes to consider a session as inactive. Sessions with idle time greater than this
    value will be candidates for termination. Default is 60 minutes.

.PARAMETER TerminateInactiveSessions
    Switch to enable termination of inactive or disconnected sessions.
    When specified, the script will terminate eligible sessions based on other parameters.

.PARAMETER Force
    Switch to bypass confirmation prompts when terminating sessions.
    Only applies when TerminateInactiveSessions is also specified.

.PARAMETER WhatIf
    Shows what would happen if the script runs. Does not perform any actions.

.EXAMPLE
    .\Manage-RDPSessions.ps1
    Lists all RDP sessions on the local machine.

.EXAMPLE
    .\Manage-RDPSessions.ps1 -ComputerName "SERVER01"
    Lists all RDP sessions on SERVER01.

.EXAMPLE
    .\Manage-RDPSessions.ps1 -State "Disconnected" -TerminateInactiveSessions
    Lists all disconnected sessions on the local machine and prompts to terminate them.

.EXAMPLE
    .\Manage-RDPSessions.ps1 -ComputerName "SERVER01" -State "Disconnected" -IdleTimeThreshold 120 -TerminateInactiveSessions -Force
    Terminates all disconnected sessions on SERVER01 that have been idle for more than 120 minutes without prompting.

.EXAMPLE
    .\Manage-RDPSessions.ps1 -ComputerName "SERVER01" -TerminateInactiveSessions -WhatIf
    Shows what sessions would be terminated without actually terminating them.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $false)]
    [string]$ComputerName = $env:COMPUTERNAME,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Active", "Disconnected", "All")]
    [string]$State = "All",

    [Parameter(Mandatory = $false)]
    [int]$IdleTimeThreshold = 60,

    [Parameter(Mandatory = $false)]
    [switch]$TerminateInactiveSessions,

    [Parameter(Mandatory = $false)]
    [switch]$Force
)

# Set up logging
$LogFolder = Join-Path -Path $PSScriptRoot -ChildPath "Logs"
if (-not (Test-Path -Path $LogFolder)) {
    New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
}

$LogFileName        = "RDPSessions_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$LogPath            = Join-Path -Path $LogFolder -ChildPath $LogFileName

function Write-LogMessage {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR", "DEBUG")]
        [string]$Level = "INFO"
    )

    $TimeStamp      = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage     = "[$TimeStamp] [$Level] $Message"

    # Write to log file
    Add-Content -Path $LogPath -Value $LogMessage

    # Write to console with appropriate color
    $ConsoleColor = switch ($Level) {
        "INFO" { "White" }
        "SUCCESS" { "Green" }
        "WARNING" { "Yellow" }
        "ERROR" { "Red" }
        "DEBUG" { "Magenta" }
        default { "White" }
    }

    Write-Output $LogMessage -ForegroundColor $ConsoleColor
}

function Get-RDPSession {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName
    )

    Write-LogMessage -Message "Retrieving RDP sessions from $ComputerName" -Level "INFO"

    try {
        # Run qwinsta to get session information
        $Output = qwinsta /server:$ComputerName 2>&1

        if ($LASTEXITCODE -ne 0) {
            throw "Error executing qwinsta: $Output"
            # Parse the output into objects
        }
        $Sessions = @()
        # Skip the header line (index 0) and start processing from line 1

        for ($i = 1; $i -lt $Output.Count; $i++) {
            $Line           = $Output[$i] -replace '\s+', ' ' -replace '^\s', ''
            $Values         = $Line -split ' '

            # Check if session ID is numeric (skip console session if needed)
            if ($Values[1] -match '^\d+$') {
                $SessionId = $Values[1]
                $Username = $Values[0]
                $SessionName = $Values[2]
                $SessionState = $Values[3]
                $IdleTime = $Values[4]

                # If username is numeric, there was likely a shift in the columns
                if ($Username -match '^\d+$') {
                    $SessionId = $Username
                    $Username = "N/A"
                    $SessionName = $Values[1]
                    $SessionState = $Values[2]
                    $IdleTime = $Values[3]
                }

                # Convert idle time to minutes
                $IdleMinutes = 0
                if ($IdleTime -match '(\d+)\+(\d+):(\d+)') {
                    # Days + hours + minutes format
                    $Days = [int]$Matches[1]
                    $Hours = [int]$Matches[2]
                    $Minutes = [int]$Matches[3]
                    $IdleMinutes = ($Days * 24 * 60) + ($Hours * 60) + $Minutes
                } elseif ($IdleTime -match '(\d+):(\d+)') {
                    # Hours + minutes format
                    $Hours = [int]$Matches[1]
                    $Minutes = [int]$Matches[2]
                    $IdleMinutes = ($Hours * 60) + $Minutes
                } elseif ($IdleTime -match '^\d+$') {
                    # Just minutes
                    $IdleMinutes = [int]$IdleTime
                } elseif ($IdleTime -eq '.') {
                    # Active session with no idle time
                    $IdleMinutes = 0
                }

                $Session = [PSCustomObject]@{
                    ComputerName = $ComputerName
                    SessionId = $SessionId
                    Username = $Username
                    SessionName = $SessionName
                    State = $SessionState
                    IdleTime = $IdleTime
                    IdleMinutes = $IdleMinutes
                }

                $Sessions += $Session
            }
        }

        return $Sessions
    } catch {
        Write-LogMessage -Message "Error retrieving RDP sessions: $_" -Level "ERROR"
        throw
    }
}

function Remove-RDPSession {
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,

        [Parameter(Mandatory = $true)]
        [string]$SessionId
    )

    try {
        $SessionInfo = "Session ID $SessionId on $ComputerName"

        if ($PSCmdlet.ShouldProcess($SessionInfo, "Terminate RDP session")) {
            Write-LogMessage -Message "Terminating $SessionInfo" -Level "INFO"
            $Output = rwinsta $SessionId /server:$ComputerName 2>&1

            if ($LASTEXITCODE -ne 0) {
                throw "Error executing rwinsta: $Output"
            }

            Write-LogMessage -Message "Successfully terminated $SessionInfo" -Level "SUCCESS"
            return $true
        } else {
            Write-LogMessage -Message "WhatIf: Would terminate $SessionInfo" -Level "DEBUG"
            return $true
        }
    } catch {
        Write-LogMessage -Message "Error terminating session: $_" -Level "ERROR"
        return $false
    }
}

# Main script execution
try {
    Write-LogMessage -Message "Script started. Target server: $ComputerName" -Level "INFO"

    # Get all RDP sessions
    $AllSessions = Get-RDPSession -ComputerName $ComputerName

    # Filter sessions based on state
    $FilteredSessions = switch ($State) {
        "Active" { $AllSessions | Where-Object { $_.State -eq "Active" } }
        "Disconnected" { $AllSessions | Where-Object { $_.State -eq "Disc" } }
        "All" { $AllSessions }
    }

    # Display all sessions
    Write-LogMessage -Message "RDP Sessions on ${ComputerName}:" -Level "INFO"
    $FilteredSessions | Format-Table -AutoSize | Out-String | ForEach-Object { Write-LogMessage -Message $_ -Level "INFO" }

    # Handle termination if requested
    if ($TerminateInactiveSessions) {
        # Identify sessions to terminate (disconnected or idle beyond threshold)
        $SessionsToTerminate = $FilteredSessions | Where-Object {
            ($_.State -eq "Disc") -or ($_.IdleMinutes -ge $IdleTimeThreshold)
        }

        if ($SessionsToTerminate.Count -eq 0) {
            Write-LogMessage -Message "No eligible sessions to terminate based on current criteria." -Level "WARNING"
        } else {
            Write-LogMessage -Message "Sessions eligible for termination:" -Level "WARNING"
            $SessionsToTerminate | Format-Table -AutoSize | Out-String | ForEach-Object { Write-LogMessage -Message $_ -Level "WARNING" }

            $ConfirmAll = $Force

            foreach ($Session in $SessionsToTerminate) {
                $SessionInfo = "Session ID: $($Session.SessionId), User: $($Session.Username), State: $($Session.State), Idle: $($Session.IdleTime)"

                $ShouldTerminate = $false

                if (-not $ConfirmAll) {
                    $Confirmation = Read-Host "Terminate $SessionInfo? (Y/N/A/Q) (Yes/No/All/Quit)"
                    switch ($Confirmation.ToUpper()) {
                        "Y" { $ShouldTerminate = $true }
                        "A" { $ConfirmAll = $true; $ShouldTerminate = $true }
                        "Q" { break }
                    }
                } else {
                    $ShouldTerminate = $true
                }

                if ($ShouldTerminate) {
                    $Result = Remove-RDPSession -ComputerName $ComputerName -SessionId $Session.SessionId
                    if ($Result) {
                        Write-LogMessage -Message "Action taken on $SessionInfo" -Level "SUCCESS"
                    }
                }
            }
        }
    }

    Write-LogMessage -Message "Script completed successfully" -Level "SUCCESS"
} catch {
    Write-LogMessage -Message "Script execution failed: $_" -Level "ERROR"
    exit 1
}
