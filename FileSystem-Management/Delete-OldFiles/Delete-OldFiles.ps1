# =============================================================================
# Script: Delete-OldFiles.ps1
# Created: 2024-02-20 17:15:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-13 18:15:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.5.0
# Additional Info: Added progress bar to show file deletion progress
# =============================================================================

<#
.SYNOPSIS
    Deletes files older than specified number of days from a target folder.
.DESCRIPTION
    This script performs the following actions:
    - Takes a specified folder path and number of days as input
    - Calculates cutoff date based on current date minus specified days
    - Finds all files older than cutoff date
    - Deletes found files and displays volume information
    - Shows drive space comparison before and after deletion
    - Optional recursive deletion of files and empty directories
    - Silent operation with error suppression
.PARAMETER StartPath
    The path to the folder containing files to be cleaned up
.PARAMETER daysOld
    Number of days old the files must be to be deleted
.PARAMETER Recurse
    Optional switch to enable recursive deletion of files and empty directories
.EXAMPLE
    .\Delete-OldFiles.ps1
    Deletes files older than 30 days from C:\windows\System32\winevt\logs
.EXAMPLE
    .\Delete-OldFiles.ps1 -StartPath "D:\Backups" -daysOld 90 -Recurse
    Recursively deletes files older than 90 days from D:\Backups and its subdirectories
.NOTES
    Security Level: Medium
    Required Permissions: Administrator rights on target folder
    Validation Requirements: Verify folder path exists before execution
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$StartPath = "C:\windows\System32\winevt\logs",
    
    [Parameter(Mandatory=$false)]
    [int]$daysOld = 30,

    [Parameter(Mandatory=$false)]
    [switch]$Recurse
)

function Show-DriveInfo {
    param (
        [Parameter(Mandatory=$true)]
        [object]$Volume
    )
    
    Write-Host "`nDrive Volume Details:" -ForegroundColor Green
    Write-Host "------------------------" -ForegroundColor Green
    Write-Host "Drive Letter: $($Volume.DriveLetter)" -ForegroundColor Cyan
    Write-Host "Drive Label: $($Volume.FileSystemLabel)" -ForegroundColor Cyan
    Write-Host "File System: $($Volume.FileSystem)" -ForegroundColor Cyan
    Write-Host "Drive Type: $($Volume.DriveType)" -ForegroundColor Cyan
    Write-Host "Size: $([math]::Round($Volume.Size/1GB, 2)) GB" -ForegroundColor Cyan
    Write-Host "Free Space: $([math]::Round($Volume.SizeRemaining/1GB, 2)) GB" -ForegroundColor Cyan
    Write-Host "Health Status: $($Volume.HealthStatus)" -ForegroundColor Cyan
    Write-Host ""
}

try {
    # Setup logging
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $logFile = Join-Path $PSScriptRoot "OldFilesDeleted_$($env:COMPUTERNAME)_$timestamp.log"
    Start-Transcript -Path $logFile -Force

    # Get the drive letter from the folder path
    $driveLetter = $StartPath.Substring(0, 1)
    
    # Get volume information before deletion
    Write-Host "Getting drive information before deletion..." -ForegroundColor Cyan
    $volumeBefore = Get-Volume -DriveLetter $driveLetter -ErrorAction Stop
    Write-Host "Drive information before file deletion:" -ForegroundColor Yellow
    Show-DriveInfo -Volume $volumeBefore

    # Get the current date
    $currentDate = Get-Date

    # Calculate the cutoff date
    $cutoffDate = $currentDate.AddDays(-$daysOld)

    Write-Host "`nDeleting files$(if($Recurse) { ' and directories' }) older than $daysOld days..." -ForegroundColor Cyan

    # Get files to delete based on recursion setting
    $oldFiles = if ($Recurse) {
        Get-ChildItem -Path $StartPath -File -Recurse | Where-Object { $_.LastWriteTime -lt $cutoffDate }
    } else {
        Get-ChildItem -Path $StartPath -File | Where-Object { $_.LastWriteTime -lt $cutoffDate }
    }

    # Initialize progress counter
    $totalFiles = $oldFiles.Count
    $currentFile = 0

    foreach ($file in $oldFiles) {
        $currentFile++
        $percentComplete = [math]::Round(($currentFile / $totalFiles) * 100, 2)
        
        Write-Progress -Activity "Deleting Old Files" `
                      -Status "Processing $currentFile of $totalFiles files ($percentComplete%)" `
                      -PercentComplete $percentComplete `
                      -CurrentOperation $file.Name
        
        Remove-Item $file.FullName -Force -ErrorAction SilentlyContinue
    }

    # Clear the progress bar
    Write-Progress -Activity "Deleting Old Files" -Completed

    # Only process directories if -Recurse is specified
    if ($Recurse) {
        $oldDirs = Get-ChildItem -Path $StartPath -Directory -Recurse | 
                   Where-Object { $_.LastWriteTime -lt $cutoffDate } |
                   Sort-Object FullName -Descending

        $totalDirs = $oldDirs.Count
        $currentDir = 0

        foreach ($dir in $oldDirs) {
            $currentDir++
            $percentComplete = [math]::Round(($currentDir / $totalDirs) * 100, 2)
            
            Write-Progress -Activity "Processing Empty Directories" `
                          -Status "Checking directory $currentDir of $totalDirs ($percentComplete%)" `
                          -PercentComplete $percentComplete `
                          -CurrentOperation $dir.Name

            if (!(Get-ChildItem -Path $dir.FullName -Force)) {
                Remove-Item $dir.FullName -Force -ErrorAction SilentlyContinue
                Write-Host "Removed empty directory: $($dir.FullName)" -ForegroundColor DarkGray
            }
        }

        # Clear the progress bar
        Write-Progress -Activity "Processing Empty Directories" -Completed
    }

    # Get volume information after deletion
    Write-Host "`nGetting drive information after deletion..." -ForegroundColor Cyan
    $volumeAfter = Get-Volume -DriveLetter $driveLetter -ErrorAction Stop
    Write-Host "Drive information after file deletion:" -ForegroundColor Yellow
    Show-DriveInfo -Volume $volumeAfter
    
    # Show space reclaimed
    $spaceReclaimed = $volumeAfter.SizeRemaining - $volumeBefore.SizeRemaining
    if ($spaceReclaimed -gt 0) {
        Write-Host "`nSpace reclaimed: $([math]::Round($spaceReclaimed/1MB, 2)) MB" -ForegroundColor Green
    } else {
        Write-Host "`nNo measurable space was reclaimed." -ForegroundColor Yellow
    }
}
catch {
    Write-Error "Error performing operation. Error: $_"
}
finally {
    # Stop logging
    try {
        Stop-Transcript
    }
    catch {
        Write-Error "Failed to stop transcript: $_"
    }
}
