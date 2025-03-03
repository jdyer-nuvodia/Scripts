# =============================================================================
# Script: Delete-OldFiles.ps1
# Created: 2024-02-20 17:15:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-05-22 15:43:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2
# Additional Info: Added drive space comparison before and after deletion
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
    - Silent operation with error suppression
.PARAMETER folderPath
    The path to the folder containing files to be cleaned up
.PARAMETER daysOld
    Number of days old the files must be to be deleted
.EXAMPLE
    .\Delete-OldFiles.ps1
    Deletes files older than 30 days from C:\windows\System32\winevt\logs
.NOTES
    Security Level: Medium
    Required Permissions: Administrator rights on target folder
    Validation Requirements: Verify folder path exists before execution
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$folderPath = "C:\windows\System32\winevt\logs",
    
    [Parameter(Mandatory=$false)]
    [int]$daysOld = 30
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
}

try {
    # Get the drive letter from the folder path
    $driveLetter = $folderPath.Substring(0, 1)
    
    # Get volume information before deletion
    Write-Host "Getting drive information before deletion..." -ForegroundColor Cyan
    $volumeBefore = Get-Volume -DriveLetter $driveLetter -ErrorAction Stop
    Write-Host "Drive information before file deletion:" -ForegroundColor Yellow
    Show-DriveInfo -Volume $volumeBefore

    # Get the current date
    $currentDate = Get-Date

    # Calculate the cutoff date
    $cutoffDate = $currentDate.AddDays(-$daysOld)

    # Get all files in the folder older than the cutoff date
    $oldFiles = Get-ChildItem -Path $folderPath -File | Where-Object { $_.LastWriteTime -lt $cutoffDate }

    # Display deletion information
    Write-Host "`nDeleting $($oldFiles.Count) files older than $daysOld days..." -ForegroundColor Cyan

    # Delete the old files silently
    foreach ($file in $oldFiles) {
        Remove-Item $file.FullName -Force -ErrorAction SilentlyContinue
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
