# =============================================================================
# Script: Get-LowestDriveVolume.ps1
# Created: 2025-02-27 21:26:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-27 21:26:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.3
# Additional Info: Modified to automatically select lowest drive letter
# =============================================================================

<#
.SYNOPSIS
    Retrieves volume information for the drive with the lowest letter.
.DESCRIPTION
    This script automatically finds and displays information for the drive
    with the lowest available drive letter in the system.
    - Automatically detects all drives
    - Selects the lowest available drive letter
    - Provides detailed volume information
.EXAMPLE
    .\Get-LowestDriveVolume.ps1
    Gets volume information for the drive with lowest letter
.NOTES
    Security Level: Low
    Required Permissions: No special permissions required
    Validation Requirements: None
#>

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