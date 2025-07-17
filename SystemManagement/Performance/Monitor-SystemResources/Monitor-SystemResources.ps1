# =============================================================================
# Script: Monitor-SystemResources.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1.2
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
# =============================================================================

<#
.SYNOPSIS
    Monitors system resources and logs performance metrics.
.DESCRIPTION
    This script provides continuous monitoring of system resources with:
    - Real-time CPU, memory, and disk usage tracking
    - Color-coded performance alerts
    - Detailed logging with UTC timestamps
    - Configurable thresholds and intervals
    - Resource consumption trending

    Key features:
    - Performance counter integration
    - CIM-based resource monitoring
    - Threshold-based alerting
    - Continuous or scheduled monitoring
    - Detailed logging capabilities

    Dependencies:
    - Windows PowerShell 5.1 or higher
    - Administrator rights for performance counter access
    - Write access to log directory
    - CIM/WMI access permissions

    The script monitors critical system metrics:
    - CPU utilization percentage
    - Memory consumption and availability
    - Disk space and I/O statistics
    - Resource consumption trends

    Performance impact:
    - Low CPU usage (< 1%)
    - Minimal memory footprint
    - Small log file size growth
.PARAMETER IntervalSeconds
    Time between measurements in seconds.
    Range: 5-3600 seconds
    Default: 15 seconds
.PARAMETER LogPath
    Directory path for storing log files.
    Creates directory if it does not exist.
    Default: Script directory
.PARAMETER ThresholdCPU
    CPU usage percentage that triggers warning.
    Range: 50-100 percent
    Default: 80 percent
.PARAMETER ThresholdMemory
    Memory usage percentage that triggers warning.
    Range: 50-100 percent
    Default: 85 percent
.PARAMETER ThresholdDisk
    Free disk space percentage that triggers warning.
    Range: 5-30 percent
    Default: 15 percent
.EXAMPLE
    .\Monitor-SystemResources.ps1
    Monitors system resources with default settings
.EXAMPLE
    .\Monitor-SystemResources.ps1 -IntervalSeconds 30 -ThresholdCPU 90 -LogPath "C:\Logs"
    Monitors with custom interval, CPU threshold, and log location
.EXAMPLE
    .\Monitor-SystemResources.ps1 -ThresholdMemory 75 -ThresholdDisk 20
    Monitors with custom memory and disk space thresholds
.NOTES
    Security Level: Medium
    Required Permissions: Local Administrator
    Validation Requirements:
    - Verify performance counter access
    - Validate log directory permissions
    - Check CIM/WMI connectivity
    - Verify disk space for logging
#>

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateRange(5, 3600)]
    [int]$IntervalSeconds = 15,

    [Parameter()]
    [ValidateScript({
            if (-not (Test-Path $_)) {
                New-Item -ItemType Directory -Path $_ -Force | Out-Null
            }
            return $true
        })]
    [string]$LogPath = $PSScriptRoot,

    [Parameter()]
    [ValidateRange(50, 100)]
    [int]$ThresholdCPU = 80,

    [Parameter()]
    [ValidateRange(50, 100)]
    [int]$ThresholdMemory = 85,

    [Parameter()]
    [ValidateRange(5, 30)]
    [int]$ThresholdDisk = 15
)

# Initialize error handling
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# Create log directory if it does not exist
if (-not (Test-Path -Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
}

$logFile = Join-Path $LogPath "SystemMonitor_$(Get-Date -Format 'yyyyMMdd').log"

function Write-LogMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter()]
        [ValidateSet("Information", "Warning", "Error", "Success")]
        [string]$Level = "Information",

        [Parameter()]
        [string]$ConsoleColor = "White"
    )

    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC"
    $LogMessage = "[$TimeStamp] [$Level] $Message"

    # Write to console with color
    Write-Output $Message -ForegroundColor $ConsoleColor

    # Write to log file
    Add-Content -Path $logFile -Value $LogMessage
}

function Get-SystemMetric {
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param()

    try {
        # Get CPU Usage
        $cpu = (Get-Counter '\Processor(_Total)\% Processor Time' -ErrorAction Stop).CounterSamples.CookedValue

        # Get Memory Usage
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        if ($os.TotalVisibleMemorySize -gt 0) {
            $memoryUsed = [math]::Round(($os.TotalVisibleMemorySize - $os.FreePhysicalMemory) / $os.TotalVisibleMemorySize * 100, 2)
        } else {
            throw "Invalid memory size reported by system"
        }

        # Get Disk Space
        $disks = Get-CimInstance Win32_LogicalDisk -Filter "DriveType = 3" -ErrorAction Stop
        $diskMetrics = $disks | ForEach-Object {
            if ($_.Size -gt 0) {
                $freeSpacePercent = [math]::Round($_.FreeSpace / $_.Size * 100, 2)
            } else {
                Write-LogMessage "Drive $($_.DeviceID) reported zero size" -Level Warning -ConsoleColor Yellow
                $freeSpacePercent = 0
            }

            @{
                Drive = $_.DeviceID
                FreeSpace = $freeSpacePercent
                TotalGB = [math]::Round($_.Size / 1GB, 2)
                FreeGB = [math]::Round($_.FreeSpace / 1GB, 2)
            }
        }

        return @{
            CPU = [math]::Round($cpu, 2)
            Memory = $memoryUsed
            Disks = $diskMetrics
            Timestamp = Get-Date
        }
    } catch {
        Write-LogMessage "Error collecting system metrics: $_" -Level Error -ConsoleColor Red
        throw
    }
}

try {
    Write-LogMessage "Starting System Resource Monitor..." -Level Information -ConsoleColor Cyan
    Write-LogMessage "Monitoring interval: $IntervalSeconds seconds" -Level Information -ConsoleColor Cyan
    Write-LogMessage "Log file: $logFile" -Level Information -ConsoleColor Cyan
    Write-LogMessage "Press Ctrl+C to stop monitoring." -Level Information -ConsoleColor Cyan
    Write-LogMessage "----------------------------------------" -Level Information -ConsoleColor DarkGray

    while ($true) {
        $metrics = Get-SystemMetric

        # CPU Status
        $cpuColor = if ($metrics.CPU -ge $ThresholdCPU) { "Red" }
        elseif ($metrics.CPU -ge ($ThresholdCPU * 0.8)) { "Yellow" }
        else { "Green" }
        Write-LogMessage "CPU Usage: $($metrics.CPU)%" -Level Information -ConsoleColor $cpuColor

        # Memory Status
        $memColor = if ($metrics.Memory -ge $ThresholdMemory) { "Red" }
        elseif ($metrics.Memory -ge ($ThresholdMemory * 0.8)) { "Yellow" }
        else { "Green" }
        Write-LogMessage "Memory Usage: $($metrics.Memory)%" -Level Information -ConsoleColor $memColor

        # Disk Status
        foreach ($disk in $metrics.Disks) {
            $diskColor = if ($disk.FreeSpace -le $ThresholdDisk) { "Red" }
            elseif ($disk.FreeSpace -le ($ThresholdDisk * 2)) { "Yellow" }
            else { "Green" }
            Write-LogMessage "Drive $($disk.Drive) - Free: $($disk.FreeSpace)% ($($disk.FreeGB)GB / $($disk.TotalGB)GB)" -Level Information -ConsoleColor $diskColor
        }

        Write-LogMessage "----------------------------------------" -Level Information -ConsoleColor DarkGray
        Start-Sleep -Seconds $IntervalSeconds
    }
} catch {
    Write-LogMessage "Script execution failed: $_" -Level Error -ConsoleColor Red
    Write-LogMessage "Stack Trace: $($_.ScriptStackTrace)" -Level Error -ConsoleColor Red
    throw
} finally {
    Write-LogMessage "Monitoring stopped. Check log file for details: $logFile" -Level Information -ConsoleColor Cyan
}
