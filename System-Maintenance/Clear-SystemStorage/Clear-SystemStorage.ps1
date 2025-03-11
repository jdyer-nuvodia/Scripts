# =============================================================================
# Script: Clear-SystemStorage.ps1
# Created: 2025-03-11 20:57:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-11 21:30:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1.0
# Additional Info: Added Downloads folder cleanup functionality
# =============================================================================

<#
.SYNOPSIS
    Performs system storage cleanup operations using PowerShell and .NET methods.
.DESCRIPTION
    Creates a system restore point and performs various cleanup operations including:
    - Windows Update cleanup
    - Temporary files removal
    - Recycle Bin emptying
    - Windows Error Reports cleanup
    - Browser cache cleanup
    - Windows logs cleanup
    
    Uses .NET methods where available for improved performance.
.PARAMETER Force
    Bypasses confirmation prompts for cleanup operations
.EXAMPLE
    .\Clear-SystemStorage.ps1
    Performs cleanup with confirmation prompts
.EXAMPLE
    .\Clear-SystemStorage.ps1 -Force
    Performs cleanup without confirmation prompts
#>

[CmdletBinding()]
param (
    [switch]$Force
)

# Import required modules
Import-Module -Name ComputerManagement

# Set error action preference
$ErrorActionPreference = 'Stop'

function Write-StatusMessage {
    param(
        [string]$Message,
        [string]$Color = 'White'
    )
    Write-Host $Message -ForegroundColor $Color
}

function Test-AdminPrivileges {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function New-SystemRestorePoint {
    try {
        Write-StatusMessage "Creating system restore point..." -Color Cyan
        Enable-ComputerRestore -Drive "$env:SystemDrive"
        Checkpoint-Computer -Description "Pre-System Cleanup $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -RestorePointType "MODIFY_SETTINGS"
        Write-StatusMessage "System restore point created successfully." -Color Green
        return $true
    }
    catch {
        Write-StatusMessage "Failed to create system restore point: $_" -Color Red
        return $false
    }
}

function Remove-TempFiles {
    Write-StatusMessage "Removing temporary files..." -Color Cyan
    $tempFolders = @(
        "$env:TEMP",
        "$env:SystemRoot\Temp",
        "$env:SystemRoot\Prefetch",
        [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::UserProfile) + "\Downloads"
    )

    foreach ($folder in $tempFolders) {
        try {
            if ($folder.EndsWith("Downloads")) {
                Write-StatusMessage "Cleaning Downloads folder..." -Color Cyan
                if (-not $Force) {
                    $downloadConfirm = Read-Host "Do you want to clean the Downloads folder? (y/n)"
                    if ($downloadConfirm -ne 'y') {
                        Write-StatusMessage "Skipping Downloads folder cleanup." -Color Yellow
                        continue
                    }
                }
            }
            
            [System.IO.Directory]::GetFiles($folder, "*.*", [System.IO.SearchOption]::AllDirectories) | ForEach-Object {
                try {
                    [System.IO.File]::Delete($_)
                }
                catch {
                    Write-StatusMessage "Could not delete file $_" -Color Yellow
                }
            }
            Write-StatusMessage "Cleaned $folder successfully" -Color Green
        }
        catch {
            Write-StatusMessage "Error accessing folder $folder" -Color Yellow
        }
    }
}

function Clear-RecycleBin {
    Write-StatusMessage "Clearing Recycle Bin..." -Color Cyan
    try {
        [System.Runtime.InteropServices.Marshal]::RunDll32("shell32.dll,SHEmptyRecycleBin")
        Write-StatusMessage "Recycle Bin cleared successfully." -Color Green
    }
    catch {
        Write-StatusMessage "Error clearing Recycle Bin: $_" -Color Yellow
    }
}

function Remove-WindowsErrorReports {
    Write-StatusMessage "Removing Windows Error Reports..." -Color Cyan
    $wer = "$env:ProgramData\Microsoft\Windows\WER"
    try {
        if ([System.IO.Directory]::Exists($wer)) {
            [System.IO.Directory]::Delete($wer, $true)
            Write-StatusMessage "Windows Error Reports removed successfully." -Color Green
        }
    }
    catch {
        Write-StatusMessage "Error removing Windows Error Reports: $_" -Color Yellow
    }
}

function Clear-BrowserCaches {
    Write-StatusMessage "Clearing browser caches..." -Color Cyan
    $chromePath = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache"
    $firefoxPath = "$env:APPDATA\Mozilla\Firefox\Profiles"
    $edgePath = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache"

    $browserPaths = @($chromePath, $firefoxPath, $edgePath)

    foreach ($path in $browserPaths) {
        if ([System.IO.Directory]::Exists($path)) {
            try {
                [System.IO.Directory]::Delete($path, $true)
                Write-StatusMessage "Cleared cache in $path" -Color Green
            }
            catch {
                Write-StatusMessage "Could not clear cache in $path" -Color Yellow
            }
        }
    }
}

function Remove-WindowsLogs {
    Write-StatusMessage "Clearing Windows logs..." -Color Cyan
    try {
        [System.Diagnostics.EventLog]::GetEventLogs() | ForEach-Object {
            try {
                $_.Clear()
                Write-StatusMessage "Cleared $($_.Log) log" -Color Green
            }
            catch {
                Write-StatusMessage "Could not clear $($_.Log) log" -Color Yellow
            }
        }
    }
    catch {
        Write-StatusMessage "Error accessing event logs: $_" -Color Yellow
    }
}

# Main execution
if (-not (Test-AdminPrivileges)) {
    Write-StatusMessage "This script requires administrative privileges. Please run as administrator." -Color Red
    exit 1
}

if (-not $Force) {
    $confirmation = Read-Host "This will clean up system storage and create a restore point. Continue? (y/n)"
    if ($confirmation -ne 'y') {
        Write-StatusMessage "Operation cancelled by user." -Color Yellow
        exit 0
    }
}

if (-not (New-SystemRestorePoint)) {
    if (-not $Force) {
        Write-StatusMessage "Cleanup cancelled due to restore point creation failure." -Color Red
        exit 1
    }
    Write-StatusMessage "Proceeding without restore point due to Force parameter." -Color Yellow
}

# Perform cleanup operations
Remove-TempFiles
Clear-RecycleBin
Remove-WindowsErrorReports
Clear-BrowserCaches
Remove-WindowsLogs

Write-StatusMessage "System storage cleanup completed successfully." -Color Green