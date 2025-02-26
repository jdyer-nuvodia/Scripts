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
        Remove-Item -Path $_.FullName -Force
    }

    Write-Host "Directory cleanup completed successfully!" -ForegroundColor Green
} catch {
    Write-Error "An error occurred during the cleanup process: $_"
    exit 1
}