# =============================================================================
# Script: Monitor-DiskSpace.ps1
# Created: 2025-04-02 17:20:00 UTC
# Author: GitHub-Copilot
# Last Updated: 2025-04-02 19:32:00 UTC
# Updated By: GitHub-Copilot
# Version: 1.1.1
# Additional Info: Changed to use $PSScriptRoot for script directory path
# =============================================================================

<#
.SYNOPSIS
    Monitors disk space usage and alerts on configurable thresholds.
.DESCRIPTION
    This script monitors disk space across all drives, logs the results,
    and provides configurable alerts based on space thresholds. It uses
    native PowerShell commands for compatibility and follows organizational
    color-coding standards for output.
.PARAMETER WarningThreshold
    Percentage at which to trigger warning alerts. Default is 80.
.PARAMETER CriticalThreshold
    Percentage at which to trigger critical alerts. Default is 90.
.EXAMPLE
    .\Monitor-DiskSpace.ps1
    Monitors all drives with default thresholds
.EXAMPLE
    .\Monitor-DiskSpace.ps1 -WarningThreshold 75 -CriticalThreshold 85
    Monitors all drives with custom thresholds
#>

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateRange(1,99)]
    [int]$WarningThreshold = 80,
    
    [Parameter()]
    [ValidateRange(1,99)]
    [int]$CriticalThreshold = 90
)

# Initialize log file path with system name and timestamp
$SystemName = $env:COMPUTERNAME
$Timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$LogPath = Join-Path $PSScriptRoot "DiskSpace_${SystemName}_${Timestamp}.log"

function Write-LogMessage {
    param(
        [string]$Message,
        [string]$Level = "Info"
    )
    
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC"
    $LogMessage = "[$TimeStamp] [$Level] $Message"
    Add-Content -Path $LogPath -Value $LogMessage
    
    # Output with appropriate color based on level
    switch ($Level) {
        "Info"      { Write-Host $LogMessage -ForegroundColor White }
        "Process"   { Write-Host $LogMessage -ForegroundColor Cyan }
        "Success"   { Write-Host $LogMessage -ForegroundColor Green }
        "Warning"   { Write-Host $LogMessage -ForegroundColor Yellow }
        "Error"     { Write-Host $LogMessage -ForegroundColor Red }
        "Debug"     { Write-Host $LogMessage -ForegroundColor Magenta }
        Default     { Write-Host $LogMessage -ForegroundColor DarkGray }
    }
}

function Get-DiskSpaceStatus {
    Write-LogMessage "Starting disk space analysis..." -Level "Process"
    
    try {
        $Drives = Get-WmiObject Win32_LogicalDisk -Filter "DriveType=3"
        foreach ($Drive in $Drives) {
            $FreeSpacePercent = [math]::Round(($Drive.FreeSpace / $Drive.Size) * 100, 2)
            $UsedSpacePercent = 100 - $FreeSpacePercent
            $FreeSpaceGB = [math]::Round($Drive.FreeSpace / 1GB, 2)
            $TotalSpaceGB = [math]::Round($Drive.Size / 1GB, 2)
            
            $Message = "Drive $($Drive.DeviceID) - Free: $FreeSpaceGB GB of $TotalSpaceGB GB ($FreeSpacePercent% free)"
            
            # Determine alert level based on thresholds
            if ($UsedSpacePercent -ge $CriticalThreshold) {
                Write-LogMessage $Message -Level "Error"
            }
            elseif ($UsedSpacePercent -ge $WarningThreshold) {
                Write-LogMessage $Message -Level "Warning"
            }
            else {
                Write-LogMessage $Message -Level "Success"
            }
        }
    }
    catch {
        Write-LogMessage "Error analyzing disk space: $_" -Level "Error"
    }
    
    Write-LogMessage "Disk space analysis complete." -Level "Process"
}

# Script execution
Write-LogMessage "=== Disk Space Monitoring Started ===" -Level "Process"
Get-DiskSpaceStatus
Write-LogMessage "=== Disk Space Monitoring Completed ===" -Level "Process"