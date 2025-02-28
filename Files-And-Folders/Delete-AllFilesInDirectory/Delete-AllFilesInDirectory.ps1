# =============================================================================
# Script: Delete-AllFilesInDirectory.ps1
# Created: 2024-02-20 17:15:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2024-02-20 17:25:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1
# Additional Info: Added parameter support for target directory
# =============================================================================

<#
.SYNOPSIS
    Recursively deletes all files and folders in a specified directory.
.DESCRIPTION
    This script removes all files and folders within a specified directory.
    It performs the deletion in two steps:
    1. Removes all files recursively
    2. Removes all folders in descending order to handle nested directories
    
    Dependencies:
    - PowerShell 5.1 or higher
    - Appropriate permissions on target directory
.PARAMETER TargetPath
    The target directory path to clean up. This parameter is mandatory.
.EXAMPLE
    .\Delete-AllFilesInDirectory.ps1 -TargetPath "C:\TempFiles"
    Deletes all contents in the specified directory "C:\TempFiles"
.NOTES
    Security Level: High
    Required Permissions: Write access to target directory
    Validation Requirements: Verify target directory before execution
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$TargetPath
)

Write-Host "Starting directory cleanup process..." -ForegroundColor Cyan
Write-Host "Target directory: $TargetPath" -ForegroundColor Cyan

try {
    # Remove all files
    Write-Host "Removing files..." -ForegroundColor Cyan
    Get-ChildItem -Path $TargetPath -File -Recurse | ForEach-Object {
        Write-Host "Deleting file: $($_.FullName)" -ForegroundColor Yellow
        Remove-Item -Path $_.FullName -Force
    }

    # Remove all folders
    Write-Host "Removing folders..." -ForegroundColor Cyan
    Get-ChildItem -Path $TargetPath -Directory -Recurse | Sort-Object -Property FullName -Descending | ForEach-Object {
        Write-Host "Deleting folder: $($_.FullName)" -ForegroundColor Yellow
        Remove-Item -Path $_.FullName -Recurse -Force
    }

    Write-Host "Directory cleanup completed successfully!" -ForegroundColor Green
} catch {
    Write-Error "An error occurred during the cleanup process: $_"
    exit 1
}

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
    # Get all available volumes with drive letters and sort them
    $volumes = Get-Volume | 
        Where-Object { $_.DriveLetter } | 
        Sort-Object DriveLetter

    if ($volumes.Count -eq 0) {
        Write-Error "No drives with letters found on the system."
        exit
    }

    # Select the volume with lowest drive letter
    $lowestVolume = $volumes[0]
    
    Write-Host "Found lowest drive letter: $($lowestVolume.DriveLetter)" -ForegroundColor Yellow
    Show-DriveInfo -Volume $lowestVolume
}
catch {
    Write-Error "Error accessing drive information. Error: $_"
}