# =============================================================================
# Script: Get-EventLogs.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-11-10 22:03:52 UTC
# Updated By: GitHub Copilot
# Version: 1.2.2
# Additional Info: Fixed StartDate/EndDate regex to ^\d{14}$ for proper validation
# =============================================================================

<#
.SYNOPSIS
    Collects Windows Event Logs for a specified time period and exports to file.
.DESCRIPTION
    This script retrieves Windows Event Logs from specified log names for a defined
    time period (either by hours or specific dates) and exports them to a file in C:\Temp as .evtx format.
    The exported files include the system name and UTC timestamp in the filename.
.PARAMETER LogNames
    Array of event log names to collect. Defaults to Application and System.
.PARAMETER Hours
    Number of hours to look back for events. Defaults to 1 hour.
.PARAMETER StartDate
    Start date in format YYYYMMDDHHMMSS. If specified, overrides Hours parameter.
.PARAMETER EndDate
    End date in format YYYYMMDDHHMMSS. If not specified when using StartDate, defaults to current time.
.EXAMPLE
    .\Get-EventLogs.ps1
    Collects last hour of Application and System logs
.EXAMPLE
    .\Get-EventLogs.ps1 -StartDate "20240220000000" -EndDate "20240220235959"
    Collects logs for entire day of February 20, 2024
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string[]]$LogNames = @("Application", "System"),

    [Parameter()]
    [int]$Hours = 1,

    [Parameter()]
    [ValidateScript({
            if ($_ -match '^\d{14}$') {
                $true
            } else {
                throw "Date must be exactly 14 digits in format YYYYMMDDHHMMSS. Example: 20240221235959"
            }
        })]
    [string]$StartDate,

    [Parameter()]
    [ValidateScript({
            if ($_ -match '^\d{14}$') {
                $true
            } else {
                throw "Date must be exactly 14 digits in format YYYYMMDDHHMMSS. Example: 20240221235959"
            }
        })]
    [string]$EndDate
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
    # Get system name
    $systemName = $env:COMPUTERNAME

    # Initialize variables
    if ($StartDate) {
        try {
            $startTime = [datetime]::ParseExact($StartDate, "yyyyMMddHHmmss", $null)
            $endTime = if ($EndDate) {
                [datetime]::ParseExact($EndDate, "yyyyMMddHHmmss", $null)
            } else {
                Get-Date
            }
        } catch {
            throw "Invalid date format. Use YYYYMMDDHHMMSS format."
        }
    } else {
        $startTime = (Get-Date).AddHours(-$Hours)
        $endTime = Get-Date
    }

    $startTimeFormatted = $startTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.000Z")
    $endTimeFormatted = $endTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.000Z")

    # Ensure output directory exists
    Initialize-OutputDirectory

    foreach ($logName in $LogNames) {
        Write-Verbose "Exporting events from $logName"
        $outputFile = "C:\Temp\${systemName}_${ logName}_$(Get-TimeStamp).evtx"

        # Create query string for time filter
        $timeQuery = "*[System[TimeCreated[@SystemTime>=`'$startTimeFormatted`'] and TimeCreated[@SystemTime<=`'$endTimeFormatted`']]]"

        Write-Verbose "Using query: $timeQuery"

        try {
            # Export events using wevtutil
            $result = wevtutil.exe export-log $logName $outputFile "/q:$timeQuery" 2>&1

            if ($LASTEXITCODE -eq 0) {
                Write-Output "Events from $logName exported to: $outputFile"
            } else {
                Write-Warning "Failed to export events from $logName. Error: $result"
            }
        } catch {
            Write-Warning "Error processing $logName : $_"
        }
    }
} catch {
    Write-Error "Error collecting event logs: $_"
    exit 1
}
