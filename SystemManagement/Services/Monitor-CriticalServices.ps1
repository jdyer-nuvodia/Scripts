# =============================================================================
# Script: Monitor-CriticalServices.ps1
# Created: 2025-04-02 17:18:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-02 17:27:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.1
# Additional Info: Fixed function name to use approved PowerShell verb
# =============================================================================

<#
.SYNOPSIS
Monitors critical Windows services and provides an interactive way to continue monitoring.

.DESCRIPTION
This script continuously monitors essential Windows services, displays their status using color-coded output,
and prompts the user to continue monitoring after each iteration. It focuses on services that are crucial
for system operation and security.

.EXAMPLE
.\Monitor-CriticalServices.ps1
Monitors critical services and asks for continuation after each check.
#>

# Define critical services to monitor
$criticalServices = @(
    "wuauserv",        # Windows Update
    "WinDefend",       # Windows Defender
    "EventLog",        # Windows Event Log
    "Dnscache",        # DNS Client
    "BITS",           # Background Intelligent Transfer Service
    "LanmanServer",   # Server
    "LanmanWorkstation", # Workstation
    "RpcSs"           # Remote Procedure Call
)

function Write-ServiceStatus {
    param(
        [string]$ServiceName,
        [string]$Status,
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

function Watch-Services {
    do {
        Clear-Host
        Write-Host "=== Critical Services Monitor ===" -ForegroundColor Cyan
        Write-Host "Timestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor DarkGray
        Write-Host "================================`n" -ForegroundColor Cyan

        foreach ($service in $criticalServices) {
            $svc = Get-Service -Name $service -ErrorAction SilentlyContinue
            if ($svc) {
                Write-ServiceStatus -ServiceName $svc.Name -Status $svc.Status -DisplayName $svc.DisplayName
            } else {
                Write-Host "Service $service not found!" -ForegroundColor Red
            }
        }

        Write-Host "`nPress 'Y' to continue monitoring, any other key to exit..." -ForegroundColor Cyan
        $continue = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        
    } while ($continue.Character -eq 'y' -or $continue.Character -eq 'Y')
}

# Start monitoring
Watch-Services

# Create log entry
$logPath = ".\ServiceMonitor.log"
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC"
"[$timestamp] Service monitoring session completed" | Out-File -FilePath $logPath -Append