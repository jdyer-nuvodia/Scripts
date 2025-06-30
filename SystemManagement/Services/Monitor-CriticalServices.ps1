# =============================================================================
# Script: Monitor-CriticalServices.ps1
# Created: 2025-04-02 17:18:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-02 21:00:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2.0
# Additional Info: Enhanced documentation, added parameter support and function-level help
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
    .\Monitor-CriticalServices.ps1 -Services "wuauserv","WinDefend","spooler"
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
    [Parameter(Mandatory=$false)]
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

    [Parameter(Mandatory=$false)]
    [ValidateScript({
        if ($_ -and !(Test-Path $_)) {
            New-Item -Path $_ -ItemType Directory -Force | Out-Null
        }
        return $true
    })]
    [string]$LogPath = $PSScriptRoot,

    [Parameter(Mandatory=$false)]
    [ValidateRange(0, 3600)]
    [int]$RefreshInterval = 0
)

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
function Write-ServiceStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$ServiceName,

        [Parameter(Mandatory=$true)]
        [string]$Status,

        [Parameter(Mandatory=$true)]
        [string]$DisplayName
    )

    switch ($Status) {
        "Running" {
            Write-Host "$DisplayName : " -NoNewline -ForegroundColor White
            Write-Host $Status -ForegroundColor Green
        }
        "Stopped" {
            Write-Host "$DisplayName : " -NoNewline -ForegroundColor White
            Write-Host $Status -ForegroundColor Red
        }
        default {
            Write-Host "$DisplayName : " -NoNewline -ForegroundColor White
            Write-Host $Status -ForegroundColor Yellow
        }
    }
}

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
.EXAMPLE
    Watch-Services
#>
function Watch-Services {
    [CmdletBinding()]
    param()

    # Create log file with system name and timestamp
    $systemName = $env:COMPUTERNAME
    $dateStamp = Get-Date -Format "yyyyMMdd"
    $logFile = Join-Path $LogPath "ServiceMonitor_${systemName}_${dateStamp}.log"

    try {
        do {
            Clear-Host
            $currentTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC"
            Write-Host "=== Critical Services Monitor ===" -ForegroundColor Cyan
            Write-Host "System: $systemName" -ForegroundColor DarkGray
            Write-Host "Timestamp: $currentTimestamp" -ForegroundColor DarkGray
            Write-Host "================================`n" -ForegroundColor Cyan

            # Log start of monitoring session
            "[$currentTimestamp] Starting service monitoring session on $systemName" | Out-File -FilePath $logFile -Append

            foreach ($service in $Services) {
                try {
                    $svc = Get-Service -Name $service -ErrorAction Stop
                    Write-ServiceStatus -ServiceName $svc.Name -Status $svc.Status -DisplayName $svc.DisplayName
                    # Log each service status
                    "[$currentTimestamp] Service: $($svc.DisplayName) - Status: $($svc.Status)" | Out-File -FilePath $logFile -Append
                }
                catch {
                    Write-Host "Service $service not found!" -ForegroundColor Red
                    # Log missing service
                    "[$currentTimestamp] ERROR: Service $service not found - $($_.Exception.Message)" | Out-File -FilePath $logFile -Append
                }
            }

            if ($RefreshInterval -gt 0) {
                Write-Host "`nAuto-refresh every $RefreshInterval seconds. Press Ctrl+C to exit..." -ForegroundColor Cyan
                Start-Sleep -Seconds $RefreshInterval
                $continue = New-Object System.Management.Automation.Host.KeyInfo
                $continue.Character = 'y'
            }
            else {
                Write-Host "`nPress 'Y' to continue monitoring, any other key to exit..." -ForegroundColor Cyan
                $continue = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }

        } while ($continue.Character -eq 'y' -or $continue.Character -eq 'Y')
    }
    catch {
        "[$currentTimestamp] ERROR: Monitoring failed - $($_.Exception.Message)" | Out-File -FilePath $logFile -Append
        throw
    }
    finally {
        # Log end of monitoring session
        "[$currentTimestamp] Service monitoring session completed on $systemName" | Out-File -FilePath $logFile -Append
    }
}

# Start monitoring
Watch-Services
