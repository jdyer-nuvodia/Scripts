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
    time period (default last hour) and exports them to a file in C:\Temp.
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
    $outputFile = "C:\Temp\EventLogs_$(Get-TimeStamp).csv"
    
    # Ensure output directory exists
    Initialize-OutputDirectory
    
    # Collect events from specified logs
    $events = foreach ($logName in $LogNames) {
        Write-Verbose "Collecting events from $logName"
        Get-WinEvent -FilterHashtable @{
            LogName = $logName
            StartTime = $startTime
        } -ErrorAction SilentlyContinue
    }
    
    # Export events to CSV
    if ($events) {
        $events | Select-Object TimeCreated, LogName, Id, LevelDisplayName, Message |
            Export-Csv -Path $outputFile -NoTypeInformation
        Write-Host "Events exported to: $outputFile"
    } else {
        Write-Warning "No events found for the specified time period"
    }
} catch {
    Write-Error "Error collecting event logs: $_"
    exit 1
}
