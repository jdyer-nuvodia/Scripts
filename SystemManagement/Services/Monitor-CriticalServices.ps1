# =============================================================================
# Script: Monitor-CriticalServices.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2.7
# Additional Info: Fixed header metadata for workflow validation
# =============================================================================

<#
.SYNOPSIS
    Monitors critical Windows services with interactive continuation and detailed logging.
.DESCRIPTION
    This script provides continuous monitoring of essential Windows services with:
    - Real-time status updates
    - Color-coded console output
    - Detailed logging with timestamps
    - Interactive continuation prompts
    - Customizable service list

    The script monitors services that are crucial for:
    - System operation (RPC, Workstation)
    - Security (Windows Defender, Event Log)
    - Networking (DNS Client, Server service)
    - System updates (Windows Update, BITS)

    Dependencies:
    - Windows PowerShell 5.1 or higher
    - Administrator rights for service status access

    The script creates a log file in the script directory with detailed
    monitoring information including timestamps and service states.
.PARAMETER Services
    Array of service names to monitor. If not specified, monitors a default
    set of critical services. Service names should be short names (not display names).
.PARAMETER LogPath
    Custom path for log file storage. If not specified, uses script directory.
.PARAMETER RefreshInterval
    Time in seconds between service checks. Default is 0 (manual refresh).
    Set to a value greater than 0 for automatic refresh.
.EXAMPLE
    .\Monitor-CriticalServices.ps1
    Monitors default critical services with manual continuation prompt
.EXAMPLE
    .\Monitor-CriticalServices.ps1 -Services "wuauserv", "WinDefend", "spooler"
    Monitors specific services only
.EXAMPLE
    .\Monitor-CriticalServices.ps1 -RefreshInterval 60 -LogPath "C:\Logs"
    Monitors services with automatic 60-second refresh, saving logs to C:\Logs
.NOTES
    Security Level: Medium
    Required Permissions: Local Administrator or Service Query rights
    Validation Requirements:
    - Verify service names exist
    - Verify log path is accessible
    - Verify sufficient permissions
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string[]]$Services = @(
        # Windows Update
        "wuauserv",
        # Windows Defender
        "WinDefend",
        # Windows Event Log
        "EventLog",
        # DNS Client
        "Dnscache",
        # Background Intelligent Transfer Service
        "BITS",
        # Server
        "LanmanServer",
        # Workstation
        "LanmanWorkstation",
        # Remote Procedure Call
        "RpcSs"
    ),

    [Parameter(Mandatory = $false)]
    [ValidateScript({
            if ($_ -and !(Test-Path $_)) {
                New-Item -Path $_ -ItemType Directory -Force | Out-Null
            }
            return $true
        })]
    [string]$LogPath = $PSScriptRoot,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 3600)]
    [int]$RefreshInterval = 0
)

# Color support variables and Write-ColorOutput function
$Script:UseAnsiColors = $PSVersionTable.PSVersion.Major -ge 7
$Script:Colors = if ($Script:UseAnsiColors) {
    @{
        'White'    = "`e[37m"
        'Cyan'     = "`e[36m"
        'Green'    = "`e[32m"
        'Yellow'   = "`e[33m"
        'Red'      = "`e[31m"
        'Magenta'  = "`e[35m"
        'DarkGray' = "`e[90m"
        'Reset'    = "`e[0m"
    }
} else {
    @{
        'White'    = [ConsoleColor]::White
        'Cyan'     = [ConsoleColor]::Cyan
        'Green'    = [ConsoleColor]::Green
        'Yellow'   = [ConsoleColor]::Yellow
        'Red'      = [ConsoleColor]::Red
        'Magenta'  = [ConsoleColor]::Magenta
        'DarkGray' = [ConsoleColor]::DarkGray
        'Reset'    = ''
    }
}

function Write-ColorOutput {
    <#
    .SYNOPSIS
    Outputs colored text in a way that's compatible with PSScriptAnalyzer requirements.

    .DESCRIPTION
    This function provides colored output while maintaining compatibility with PSScriptAnalyzer
    by using only Write-Output and standard PowerShell cmdlets.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$Color = "White"
    )

    # Always use Write-Output to satisfy PSScriptAnalyzer
    # For PowerShell 7+, include ANSI color codes in the output
    if ($Script:UseAnsiColors) {
        $colorCode      = $Script:Colors[$Color]
        $resetCode      = $Script:Colors.Reset
        Write-Output "${colorCode}${Message}${resetCode}"
    } else {
        # For PowerShell 5.1, just output the message
        # Color formatting will be handled by the terminal/host if supported
        Write-Output $Message
    }
}



function Write-ServiceStatus {
    <#
    .SYNOPSIS
        Writes service status with color-coded output.
    .DESCRIPTION
        Displays service information with consistent color coding:
        - Running: Green
        - Stopped: Red
        - Other states: Yellow
        Service name is always in white for readability.
    .PARAMETER ServiceName
        The service's system name
    .PARAMETER Status
        Current status of the service
    .PARAMETER DisplayName
        User-friendly display name of the service
    .EXAMPLE
        Write-ServiceStatus -ServiceName "wuauserv" -Status "Running" -DisplayName "Windows Update"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ServiceName,

        [Parameter(Mandatory = $true)]
        [string]$Status,

        [Parameter(Mandatory = $true)]
        [string]$DisplayName
    )

    # Validate that ServiceName is not empty
    if ([string]::IsNullOrWhiteSpace($ServiceName)) {
        throw "ServiceName cannot be null or empty"
    }

    switch ($Status) {
        "Running" {
            Write-ColorOutput -Message "$DisplayName : " -Color "White"
            Write-ColorOutput -Message $Status -Color 'Green'
        }
        "Stopped" {
            Write-ColorOutput -Message "$DisplayName : " -Color "White"
            Write-ColorOutput -Message $Status -Color 'Red'
        }
        default {
            Write-ColorOutput -Message "$DisplayName : " -Color "White"
            Write-ColorOutput -Message $Status -Color 'Yellow'
        }
    }
}

function Watch-Service {
    <#
    .SYNOPSIS
        Monitors services and logs their status.
    .DESCRIPTION
        Main monitoring function that:
        1. Creates and maintains log file
        2. Continuously monitors specified services
        3. Displays real-time status
        4. Handles user interaction for continuation
        5. Logs all events with timestamps
    .PARAMETER Services
        Array of service names to monitor
    .PARAMETER LogPath
        Path for log file storage
    .PARAMETER RefreshInterval
        Time in seconds between service checks
    .EXAMPLE
        Watch-Service -Services $Services -LogPath $LogPath -RefreshInterval $RefreshInterval
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Services,

        [Parameter(Mandatory = $true)]
        [string]$LogPath,

        [Parameter(Mandatory = $true)]
        [int]$RefreshInterval
    )

    # Create log file with system name and timestamp
    $systemName     = $env:COMPUTERNAME
    $dateStamp      = Get-Date -Format "yyyyMMdd"
    $logFile        = Join-Path $LogPath "ServiceMonitor_${systemName}_${dateStamp}.log"

    try {
        do {
            Clear-Host
            $currentTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC"
            Write-ColorOutput -Message "=== Critical Services Monitor ===" -Color 'Cyan'
            Write-ColorOutput -Message "System: $systemName" -Color 'DarkGray'
            Write-ColorOutput -Message "Timestamp: $currentTimestamp" -Color 'DarkGray'
            Write-ColorOutput -Message "================================`n" -Color 'Cyan'

            # Log start of monitoring session
            "[$currentTimestamp] Starting service monitoring session on $systemName" | Out-File -FilePath $logFile -Append

            foreach ($service in $Services) {
                try {
                    $svc = Get-Service -Name $service -ErrorAction Stop
                    Write-ServiceStatus -ServiceName $svc.Name -Status $svc.Status -DisplayName $svc.DisplayName
                    # Log each service status
                    "[$currentTimestamp] Service: $($svc.DisplayName) - Status: $($svc.Status)" | Out-File -FilePath $logFile -Append
                } catch {
                    Write-ColorOutput -Message "Service $service not found!" -Color 'Red'
                    # Log missing service
                    "[$currentTimestamp] ERROR: Service $service not found - $($_.Exception.Message)" | Out-File -FilePath $logFile -Append
                }
            }

            if ($RefreshInterval -gt 0) {
                Write-ColorOutput -Message "`nAuto-refresh every $RefreshInterval seconds. Press Ctrl+C to exit..." -Color 'Cyan'
                Start-Sleep -Seconds $RefreshInterval
                $continue = New-Object System.Management.Automation.Host.KeyInfo
                $continue.Character = 'y'
            } else {
                Write-ColorOutput -Message "`nPress 'Y' to continue monitoring, any other key to exit..." -Color 'Cyan'
                $continue = $host.UI.RawUI.ReadKey("NoEcho, IncludeKeyDown")
            }

        } while ($continue.Character -eq 'y' -or $continue.Character -eq 'Y')
    } catch {
        "[$currentTimestamp] ERROR: Monitoring failed - $($_.Exception.Message)" | Out-File -FilePath $logFile -Append
        throw
    } finally {
        # Log end of monitoring session
        "[$currentTimestamp] Service monitoring session completed on $systemName" | Out-File -FilePath $logFile -Append
    }
}

# Start monitoring
Watch-Service -Services $Services -LogPath $LogPath -RefreshInterval $RefreshInterval
