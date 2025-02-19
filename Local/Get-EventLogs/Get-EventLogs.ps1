# =============================================================================
# Script: Get-EventLogs.ps1
# Created: 2024-02-12 18:30:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2024-02-12 18:30:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0
# Additional Info: Initial script creation for event log collection
# =============================================================================

<#
.SYNOPSIS
    Collects Windows Event Logs for a specified time period and exports to file.
.DESCRIPTION
    This script retrieves Windows Event Logs from specified log names for a defined
    time period (default last hour) and exports them to a file in C:\Temp as .evtx format.
.PARAMETER LogNames
    Array of event log names to collect. Defaults to Application and System.
.PARAMETER Hours
    Number of hours to look back for events. Defaults to 1 hour.
.EXAMPLE
    .\Get-EventLogs.ps1
    Collects last hour of Application and System logs
.EXAMPLE
    .\Get-EventLogs.ps1 -LogNames "Application","System","Security" -Hours 24
    Collects last 24 hours of specified logs
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string[]]$LogNames = @("Application", "System"),
    
    [Parameter()]
    [int]$Hours = 1
)

# Function to ensure output directory exists
function Initialize-OutputDirectory {
    try {
        if (-not (Test-Path "C:\Temp")) {
            New-Item -ItemType Directory -Path "C:\Temp" -Force | Out-Null
        }
    } catch {
        throw "Failed to create output directory: $_"
    }
}

# Function to get timestamp for filename
function Get-TimeStamp {
    return (Get-Date -Format "yyyyMMdd_HHmmss")
}

# Main script execution
try {
    # Initialize variables
    $startTime = (Get-Date).AddHours(-$Hours)
    
    # Ensure output directory exists
    Initialize-OutputDirectory
    
    foreach ($logName in $LogNames) {
        Write-Verbose "Exporting events from $logName"
        $outputFile = "C:\Temp\${logName}_$(Get-TimeStamp).evtx"
        
        # Create query string for time filter
        $timeQuery = "*[System[TimeCreated[@SystemTime>='$(Get-Date $startTime -Format o)']]"
        
        # Export events using wevtutil
        $result = wevtutil.exe export-log $logName $outputFile /q:$timeQuery
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "Events from $logName exported to: $outputFile"
        } else {
            Write-Warning "Failed to export events from $logName"
        }
    }
} catch {
    Write-Error "Error collecting event logs: $_"
    exit 1
}
