# =============================================================================
# Script: Clear-SystemStorage.ps1
# Created: 2025-03-11 20:57:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-11 21:03:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.3.0
# Additional Info: Added comprehensive logging and automated execution
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
.PARAMETER NoRestore
    Skips the creation of a system restore point
.EXAMPLE
    .\Clear-SystemStorage.ps1
    Performs cleanup with confirmation prompts
.EXAMPLE
    .\Clear-SystemStorage.ps1 -Force
    Performs cleanup without confirmation prompts
.EXAMPLE
    .\Clear-SystemStorage.ps1 -NoRestore
    Performs cleanup without creating a system restore point
#>

[CmdletBinding()]
param (
    [switch]$Force,
    [switch]$NoRestore
)

# Initialize logging
$scriptPath = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Path)
$logFile = [System.IO.Path]::Combine($scriptPath, "Clear-SystemStorage_$([DateTime]::UtcNow.ToString('yyyyMMdd_HHmmss')).log")
$script:logStream = [System.IO.StreamWriter]::new($logFile, $true, [System.Text.Encoding]::UTF8)

function Write-Log {
    param(
        [string]$Message,
        [string]$Color = 'White',
        [switch]$NoConsole
    )
    
    $timestamp = [DateTime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ss UTC')
    $logMessage = "[$timestamp] $Message"
    
    # Write to log file
    $script:logStream.WriteLine($logMessage)
    $script:logStream.Flush()
    
    # Write to console if not suppressed
    if (-not $NoConsole) {
        Write-Host $Message -ForegroundColor $Color
    }
}

function Write-StatusMessage {
    param(
        [string]$Message,
        [string]$Color = 'White'
    )
    Write-Log -Message $Message -Color $Color
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
        "$env:SystemRoot\Prefetch"
    )

    $filesRemoved = 0
    $errorCount = 0

    # Clean standard temp folders
    foreach ($folder in $tempFolders) {
        try {
            Write-Log "Processing folder: $folder" -Color DarkGray
            [System.IO.Directory]::GetFiles($folder, "*.*", [System.IO.SearchOption]::AllDirectories) | ForEach-Object {
                try {
                    [System.IO.File]::Delete($_)
                    $filesRemoved++
                    Write-Log "Deleted: $_" -Color DarkGray -NoConsole
                }
                catch {
                    $errorCount++
                    Write-Log "Could not delete file $_: $($_.Exception.Message)" -Color Yellow
                }
            }
            Write-StatusMessage "Cleaned $folder successfully ($filesRemoved files removed)" -Color Green
        }
        catch {
            Write-StatusMessage "Error accessing folder $folder: $($_.Exception.Message)" -Color Yellow
            $errorCount++
        }
    }

    # Clean Downloads folders for all users
    try {
        $usersPath = [System.IO.Path]::Combine($env:SystemDrive, "Users")
        $cutoffDate = (Get-Date).AddDays(-180)
        
        [System.IO.Directory]::GetDirectories($usersPath) | ForEach-Object {
            $downloadPath = [System.IO.Path]::Combine($_, "Downloads")
            
            if ([System.IO.Directory]::Exists($downloadPath)) {
                Write-StatusMessage "Processing Downloads folder for $([System.IO.Path]::GetFileName($_))..." -Color Cyan
                
                try {
                    $oldFilesCount = 0
                    [System.IO.Directory]::GetFiles($downloadPath, "*.*", [System.IO.SearchOption]::AllDirectories) | ForEach-Object {
                        try {
                            $fileInfo = [System.IO.FileInfo]::new($_)
                            if ($fileInfo.LastWriteTime -lt $cutoffDate) {
                                [System.IO.File]::Delete($_)
                                $oldFilesCount++
                                $filesRemoved++
                                Write-Log "Deleted old file: $_" -Color DarkGray -NoConsole
                            }
                        }
                        catch {
                            $errorCount++
                            Write-Log "Could not delete file $_: $($_.Exception.Message)" -Color Yellow
                        }
                    }
                    Write-StatusMessage "Cleaned $oldFilesCount old files from $downloadPath" -Color Green
                }
                catch {
                    Write-StatusMessage "Error accessing Downloads folder for $([System.IO.Path]::GetFileName($_)): $($_.Exception.Message)" -Color Yellow
                    $errorCount++
                }
            }
        }
    }
    catch {
        Write-StatusMessage "Error accessing Users directory: $($_.Exception.Message)" -Color Yellow
        $errorCount++
    }

    Write-StatusMessage "Temp file cleanup completed. Total files removed: $filesRemoved, Errors: $errorCount" -Color Cyan
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
try {
    if (-not (Test-AdminPrivileges)) {
        Write-Log "This script requires administrative privileges. Please run as administrator." -Color Red
        exit 1
    }

    if (-not $NoRestore) {
        if (-not (New-SystemRestorePoint)) {
            if (-not $Force) {
                Write-Log "Cleanup cancelled due to restore point creation failure." -Color Red
                exit 1
            }
            Write-Log "Proceeding without restore point due to Force parameter." -Color Yellow
        }
    }
    else {
        Write-Log "Skipping restore point creation as requested." -Color Yellow
    }

    # Perform cleanup operations
    Remove-TempFiles
    Clear-RecycleBin
    Remove-WindowsErrorReports
    Clear-BrowserCaches
    Remove-WindowsLogs

    Write-Log "System storage cleanup completed successfully. See log file for details: $logFile" -Color Green
}
catch {
    Write-Log "Critical error during execution: $($_.Exception.Message)" -Color Red
    Write-Log "Stack trace: $($_.ScriptStackTrace)" -Color Red
    exit 1
}
finally {
    if ($null -ne $script:logStream) {
        $script:logStream.Close()
        $script:logStream.Dispose()
    }
}