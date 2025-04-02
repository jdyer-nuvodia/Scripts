# =============================================================================
# Script: Monitor-SystemResources.ps1
# Created: 2025-04-02 15:18:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-02 15:18:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.0
# Additional Info: Initial creation - System resource monitoring script
# =============================================================================

<#
.SYNOPSIS
Monitors system resources and logs performance metrics.

.DESCRIPTION
Continuously monitors CPU usage, memory consumption, disk space, and network statistics.
Logs the data to a file and provides real-time console output with color-coding.
Supports configurable monitoring intervals and thresholds.

.PARAMETER IntervalSeconds
The interval between measurements in seconds. Default is 60.

.PARAMETER LogPath
The path where log files will be stored. Default is "./logs"

.PARAMETER ThresholdCPU
CPU usage percentage that triggers a warning. Default is 80.

.PARAMETER ThresholdMemory
Memory usage percentage that triggers a warning. Default is 85.

.PARAMETER ThresholdDisk
Free disk space percentage that triggers a warning. Default is 15.

.EXAMPLE
.\Monitor-SystemResources.ps1
Monitors system resources with default settings

.EXAMPLE
.\Monitor-SystemResources.ps1 -IntervalSeconds 30 -ThresholdCPU 90
Monitors with 30-second intervals and 90% CPU threshold
#>

[CmdletBinding()]
param(
    [int]$IntervalSeconds = 60,
    [string]$LogPath = ".\logs",
    [int]$ThresholdCPU = 80,
    [int]$ThresholdMemory = 85,
    [int]$ThresholdDisk = 15
)

# Create log directory if it does not exist
if (-not (Test-Path -Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath | Out-Null
}

$logFile = Join-Path $LogPath "SystemMonitor_$(Get-Date -Format 'yyyyMMdd').log"

function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
    Add-Content -Path $logFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC'): $Message"
}

function Get-SystemMetrics {
    # Get CPU Usage
    $cpu = (Get-Counter '\Processor(_Total)\% Processor Time').CounterSamples.CookedValue

    # Get Memory Usage
    $os = Get-Ciminstance Win32_OperatingSystem
    $memoryUsed = [math]::Round(($os.TotalVisibleMemorySize - $os.FreePhysicalMemory) / $os.TotalVisibleMemorySize * 100, 2)

    # Get Disk Space
    $disks = Get-WmiObject Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }
    $diskMetrics = $disks | ForEach-Object {
        @{
            Drive = $_.DeviceID
            FreeSpace = [math]::Round($_.FreeSpace / $_.Size * 100, 2)
        }
    }

    return @{
        CPU = [math]::Round($cpu, 2)
        Memory = $memoryUsed
        Disks = $diskMetrics
    }
}

Write-ColorOutput "Starting System Resource Monitor..." -Color Cyan
Write-ColorOutput "Press Ctrl+C to stop monitoring." -Color Cyan

while ($true) {
    $metrics = Get-SystemMetrics

    # CPU Status
    $cpuColor = if ($metrics.CPU -ge $ThresholdCPU) { "Red" } elseif ($metrics.CPU -ge ($ThresholdCPU * 0.8)) { "Yellow" } else { "Green" }
    Write-ColorOutput "CPU Usage: $($metrics.CPU)%" -Color $cpuColor

    # Memory Status
    $memColor = if ($metrics.Memory -ge $ThresholdMemory) { "Red" } elseif ($metrics.Memory -ge ($ThresholdMemory * 0.8)) { "Yellow" } else { "Green" }
    Write-ColorOutput "Memory Usage: $($metrics.Memory)%" -Color $memColor

    # Disk Status
    foreach ($disk in $metrics.Disks) {
        $diskColor = if ($disk.FreeSpace -le $ThresholdDisk) { "Red" } elseif ($disk.FreeSpace -le ($ThresholdDisk * 2)) { "Yellow" } else { "Green" }
        Write-ColorOutput "Drive $($disk.Drive) Free Space: $($disk.FreeSpace)%" -Color $diskColor
    }

    Write-ColorOutput "----------------------------------------" -Color DarkGray
    Start-Sleep -Seconds $IntervalSeconds
}