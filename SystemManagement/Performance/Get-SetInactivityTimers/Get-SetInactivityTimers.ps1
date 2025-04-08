# =============================================================================
# Script: Get-SetInactivityTimers.ps1
# Created: 2025-04-08 21:45:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-08 23:14:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.1
# Additional Info: Fixed PowerShell analyzer warnings and improved ShouldProcess implementation
# =============================================================================

<#
.SYNOPSIS
Gets and optionally sets Windows system inactivity timers.

.DESCRIPTION
This script retrieves all available system inactivity settings including screen timeout,
sleep settings, and power management configurations. It displays these settings to the user
and provides the option to modify them. All changes support -WhatIf functionality for safety.

.EXAMPLE
.\Get-SetInactivityTimers.ps1
Displays current inactivity settings and prompts for changes

.EXAMPLE
.\Get-SetInactivityTimers.ps1 -WhatIf
Shows what changes would be made without actually making them
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param()

# Function to format minutes into a readable string
function Format-Minutes {
    param([int]$Minutes)
    if ($Minutes -eq 0) { return "Never" }
    if ($Minutes -ge 1440) { return "$($Minutes / 1440) hours" }
    if ($Minutes -ge 60) { return "$($Minutes / 60) hours" }
    return "$Minutes minutes"
}

function Get-PowerSettings {
    Write-Host "Retrieving current power settings..." -ForegroundColor Cyan
      # Get current power scheme GUID
    $schemeGuid = (powercfg /getactivescheme) -split " " | Where-Object { $_ -match '^[{]?[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}[}]?$' }
    if (-not $schemeGuid) {
        Write-Warning "Could not determine active power scheme GUID"
        return $null
    }
    # Get monitor timeout settings using active scheme
    $monitorTimeoutAC = powercfg /query $schemeGuid SUB_VIDEO VIDEOIDLE | Select-String "AC Power Setting Index: ([0-9a-fx]+)" | ForEach-Object { $_.Matches.Groups[1].Value }
    $monitorTimeoutDC = powercfg /query $schemeGuid SUB_VIDEO VIDEOIDLE | Select-String "DC Power Setting Index: ([0-9a-fx]+)" | ForEach-Object { $_.Matches.Groups[1].Value }
    
    # Get sleep settings using active scheme
    $sleepTimeoutAC = powercfg /query $schemeGuid SUB_SLEEP STANDBYIDLE | Select-String "AC Power Setting Index: ([0-9a-fx]+)" | ForEach-Object { $_.Matches.Groups[1].Value }
    $sleepTimeoutDC = powercfg /query $schemeGuid SUB_SLEEP STANDBYIDLE | Select-String "DC Power Setting Index: ([0-9a-fx]+)" | ForEach-Object { $_.Matches.Groups[1].Value }
    
    # Screen saver settings from registry
    $screenSaverTimeout = Get-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "ScreenSaveTimeout" -ErrorAction SilentlyContinue
    
    return @{
        MonitorAC = $monitorTimeoutAC
        MonitorDC = $monitorTimeoutDC
        SleepAC = $sleepTimeoutAC
        SleepDC = $sleepTimeoutDC
        ScreenSaver = $screenSaverTimeout
    }
}

function Set-PowerTimeout {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$true)]
        [int]$TimeoutMinutes
    )
    
    if ($PSCmdlet.ShouldProcess("Power Settings", "Set timeout to $TimeoutMinutes minutes")) {
        try {            # Get current power scheme GUID
            $schemeGuid = (powercfg /getactivescheme) -split " " | Where-Object { $_ -match '^[{]?[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}[}]?$' }
            if (-not $schemeGuid) {
                Write-Warning "Could not determine active power scheme GUID"
                return $false
            }
            
            # Convert minutes to seconds
            $timeoutSeconds = $TimeoutMinutes * 60
            
            # Set monitor timeout (AC and DC)
            Write-Host "Setting monitor timeout..." -ForegroundColor Cyan
            powercfg /change monitor-timeout-ac $TimeoutMinutes
            powercfg /change monitor-timeout-dc $TimeoutMinutes
            
            # Set sleep timeout (AC and DC)
            Write-Host "Setting sleep timeout..." -ForegroundColor Cyan
            powercfg /change standby-timeout-ac $TimeoutMinutes
            powercfg /change standby-timeout-dc $TimeoutMinutes
            
            # Set screen saver timeout
            Write-Host "Setting screen saver timeout..." -ForegroundColor Cyan
            Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "ScreenSaveTimeout" -Value $timeoutSeconds
            
            Write-Host "All inactivity timers have been updated successfully!" -ForegroundColor Green
        }
        catch {
            Write-Host "Error setting power settings: $_" -ForegroundColor Red
            return $false
        }
    }
    return $true
}

# Main script execution
try {
    # Get current settings
    $currentSettings = Get-PowerSettings
    
    # Display current settings
    Write-Host "`nCurrent Inactivity Settings:" -ForegroundColor White
    Write-Host "------------------------" -ForegroundColor White
    Write-Host "Monitor Timeout (AC): $(Format-Minutes $currentSettings.MonitorAC)"
    Write-Host "Monitor Timeout (Battery): $(Format-Minutes $currentSettings.MonitorDC)"
    Write-Host "Sleep Timeout (AC): $(Format-Minutes $currentSettings.SleepAC)"
    Write-Host "Sleep Timeout (Battery): $(Format-Minutes $currentSettings.SleepDC)"
    Write-Host "Screen Saver Timeout: $(Format-Minutes $currentSettings.ScreenSaver)"
    
    # Ask if user wants to change settings
    $response = Read-Host "`nWould you like to change these settings? (Y/N)"
    
    if ($response -eq "Y") {
        $newTimeout = Read-Host "Enter new timeout in minutes (0 for never)"
        if ($newTimeout -match '^\d+$') {
            if (Set-PowerTimeout -TimeoutMinutes ([int]$newTimeout)) {
                Write-Host "Settings updated successfully!" -ForegroundColor Green
            }
        }
        else {
            Write-Host "Invalid input. Please enter a number." -ForegroundColor Red
        }
    }
}
catch {
    Write-Host "An error occurred: $_" -ForegroundColor Red
    exit 1
}
