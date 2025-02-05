# Script: Test-PingExtended.ps1
# Version: 2.2
# Description: Extended ping test with network configuration logging and continuous mode
# Author: jdyer-nuvodia
# Created: 2025-02-05 23:37:02
#
[CmdletBinding()]
param(
    [Parameter(Position=0)]
    [string]$Target = "8.8.8.8",
    
    [Parameter(Position=1)]
    [int]$Count = 0,  # 0 means continuous
    
    [Parameter()]
    [string]$OutputPath = "C:\PingLogs"  # Changed default path
)

# Create output directory if it doesn't exist
if (!(Test-Path -Path $OutputPath)) {
    try {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        Write-Host "Created output directory: $OutputPath" -ForegroundColor Yellow
    }
    catch {
        Write-Error "Failed to create output directory: $_"
        exit 1
    }
}

# Rest of the script remains the same until the finally block...

finally {
    if ($logFile) {
        # Log final statistics
        $packetLoss = if ($sent -gt 0) { 100 - ($received / $sent * 100) } else { 0 }
        $avgTime = if ($received -gt 0) { $totalTime / $received } else { 0 }
        
        $finalStats = @"

========================================
Final Statistics:
========================================
Test Duration: $((Get-Date) - (Get-Item $logFile).CreationTime)
Packets: Sent = $sent, Received = $received, Lost = $($sent - $received) ($($packetLoss.ToString('N2'))% loss)
Round Trip Times: Min = $($minTime)ms, Max = $($maxTime)ms, Avg = $($avgTime.ToString('N2'))ms
========================================
Test completed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Log file size: $(Get-FormattedSize (Get-Item $logFile).Length)
========================================
"@
        Add-Content -Path $logFile -Value $finalStats
        Write-Host $finalStats -ForegroundColor Cyan

        # Add clear message about log file location
        Write-Host "`n==================================================" -ForegroundColor Green
        Write-Host "Log file has been created:" -ForegroundColor Green
        Write-Host "Name: $(Split-Path $logFile -Leaf)" -ForegroundColor Yellow
        Write-Host "Location: $(Split-Path $logFile)" -ForegroundColor Yellow
        Write-Host "Full Path: $logFile" -ForegroundColor Yellow
        Write-Host "Size: $(Get-FormattedSize (Get-Item $logFile).Length)" -ForegroundColor Yellow
        Write-Host "==================================================" -ForegroundColor Green
    }
}