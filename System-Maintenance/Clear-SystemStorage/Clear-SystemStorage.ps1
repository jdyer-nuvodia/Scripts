# =============================================================================
# Script: Clear-SystemStorage.ps1
# Created: 2025-03-11 20:57:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-11 21:37:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.4.2
# Additional Info: Modified Windows Event Log clearing to only clear logs older than 30 days
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
$systemName = [System.Environment]::MachineName
$logFile = [System.IO.Path]::Combine($scriptPath, "Clear-SystemStorage_${systemName}_$([DateTime]::UtcNow.ToString('yyyyMMdd_HHmmss')).log")
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
    $lockedFiles = @{
        Count = 0
        TotalSize = 0
    }
    $removedFiles = @{
        Count = 0
        TotalSize = 0
    }

    $standardTempFolders = @(
        [System.IO.Path]::GetTempPath(),
        "$env:SystemRoot\Temp",
        "$env:SystemRoot\Prefetch"
    )

    # Clean standard temp folders
    foreach ($folder in $standardTempFolders) {
        Write-Log "Processing folder: $folder" -Color Cyan
        try {
            if ([System.IO.Directory]::Exists($folder)) {
                Get-ChildItem -Path $folder -File -Force -ErrorAction SilentlyContinue | ForEach-Object {
                    try {
                        $fileSize = $_.Length
                        [System.IO.File]::Delete($_.FullName)
                        $removedFiles.Count++
                        $removedFiles.TotalSize += $fileSize
                        Write-Log "Deleted: $_" -Color DarkGray -NoConsole
                    }
                    catch {
                        $lockedFiles.Count++
                        $lockedFiles.TotalSize += $fileSize
                    }
                }
                Write-StatusMessage "Processed $folder" -Color Green
            }
        }
        catch {
            Write-Log "Error accessing folder ${folder}: $($_.Exception.Message)" -Color Yellow
        }
    }

    # Clean Downloads folders for all users
    try {
        $usersPath = [System.IO.Path]::Combine($env:SystemDrive, "Users")
        $cutoffDate = (Get-Date).AddDays(-180)
        
        [System.IO.Directory]::GetDirectories($usersPath) | ForEach-Object {
            $downloadPath = [System.IO.Path]::Combine($_, "Downloads")
            
            if ([System.IO.Directory]::Exists($downloadPath)) {
                Write-Log "Processing Downloads folder for $([System.IO.Path]::GetFileName($_))..." -Color Cyan
                
                try {
                    Get-ChildItem -Path $downloadPath -File -Force -ErrorAction SilentlyContinue | 
                        Where-Object { $_.LastWriteTime -lt $cutoffDate } | 
                        ForEach-Object {
                            try {
                                $fileSize = $_.Length
                                [System.IO.File]::Delete($_.FullName)
                                $removedFiles.Count++
                                $removedFiles.TotalSize += $fileSize
                                Write-Log "Deleted old file: $_" -Color DarkGray -NoConsole
                            }
                            catch {
                                $lockedFiles.Count++
                                $lockedFiles.TotalSize += $fileSize
                            }
                    }
                    Write-StatusMessage "Processed $downloadPath" -Color Green
                }
                catch {
                    Write-Log "Error accessing Downloads folder for $([System.IO.Path]::GetFileName($_)): $($_.Exception.Message)" -Color Yellow
                }
            }
        }
    }
    catch {
        Write-Log "Error accessing Users directory: $($_.Exception.Message)" -Color Yellow
    }

    # Format sizes for display
    $removedSizeGB = [math]::Round($removedFiles.TotalSize / 1GB, 2)
    $lockedSizeGB = [math]::Round($lockedFiles.TotalSize / 1GB, 2)

    Write-StatusMessage "Temp file cleanup completed:" -Color Cyan
    Write-Log "- Files removed: $($removedFiles.Count) ($($removedSizeGB) GB)" -Color Green
    Write-Log "- Files locked: $($lockedFiles.Count) ($($lockedSizeGB) GB)" -Color Yellow

    return $removedFiles.Count
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

function Clear-ShadowCopies {
    Write-StatusMessage "Managing Volume Shadow Copies..." -Color Cyan
    try {
        # Get current shadow copies using vssadmin
        $tempFile = [System.IO.Path]::GetTempFileName()
        $process = Start-Process -FilePath "vssadmin" -ArgumentList "list shadows" -NoNewWindow -Wait -RedirectStandardOutput $tempFile -PassThru
        if ($process.ExitCode -ne 0) {
            throw "vssadmin failed with exit code $($process.ExitCode)"
        }
        $shadowList = [System.IO.File]::ReadAllText($tempFile)
        [System.IO.File]::Delete($tempFile)
        
        # Parse shadow copies
        $shadowCopies = @($shadowList | Select-String -Pattern "Shadow Copy ID: {(.*?)}" -AllMatches | 
            ForEach-Object { $_.Matches.Groups[1].Value })
            
        $totalCopies = $shadowCopies.Count
        
        if ($totalCopies -gt 1) {
            Write-StatusMessage "Found $totalCopies shadow copies. Keeping most recent only." -Color Yellow
            
            # Keep the last one (most recent), delete the rest
            $shadowCopies | Select-Object -SkipLast 1 | ForEach-Object {
                try {
                    $process = Start-Process -FilePath "vssadmin" -ArgumentList "delete shadows /Shadow={$_} /Quiet" -NoNewWindow -Wait -PassThru
                    if ($process.ExitCode -eq 0) {
                        Write-Log "Deleted shadow copy: $_" -Color DarkGray -NoConsole
                    }
                    else {
                        Write-Log "Failed to delete shadow copy $_. Exit code: $($process.ExitCode)" -Color Yellow
                    }
                }
                catch {
                    Write-Log "Error deleting shadow copy ${_}: $($_.Exception.Message)" -Color Yellow
                }
            }
            Write-StatusMessage "Shadow copy cleanup completed. Kept most recent copy." -Color Green
        }
        else {
            Write-StatusMessage "No excess shadow copies found (Current count: $totalCopies)." -Color Green
        }
    }
    catch {
        Write-StatusMessage "Error managing shadow copies: $($_.Exception.Message)" -Color Yellow
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
                # Skip if log name is empty or null
                if ([string]::IsNullOrWhiteSpace($_.Log)) {
                    Write-Log "Skipped empty log name" -Color Yellow -NoConsole
                    return
                }
                
                # Only clear logs older than 30 days
                $cutoffDate = (Get-Date).AddDays(-30)
                $oldEntries = $_.Entries | Where-Object { $_.TimeGenerated -lt $cutoffDate }
                
                if ($oldEntries.Count -gt 0) {
                    $_.Clear()
                    Write-StatusMessage "Cleared old entries from $($_.Log) log" -Color Green
                } else {
                    Write-StatusMessage "No entries older than 30 days in $($_.Log) log" -Color DarkGray
                }
            }
            catch {
                Write-StatusMessage "Could not process $($_.Log) log: $_" -Color Yellow
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
        Clear-ShadowCopies
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