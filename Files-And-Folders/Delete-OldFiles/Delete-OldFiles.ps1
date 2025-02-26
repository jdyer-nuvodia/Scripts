# =============================================================================
# Script: Delete-OldFiles.ps1
# Created: 2024-02-20 17:15:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2024-02-20 17:20:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1
# Additional Info: Added parameter support with default values
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

# Get the current date
$currentDate = Get-Date

# Calculate the cutoff date
$cutoffDate = $currentDate.AddDays(-$daysOld)

# Get all files in the folder older than the cutoff date
$oldFiles = Get-ChildItem -Path $folderPath -File | Where-Object { $_.LastWriteTime -lt $cutoffDate }

# Delete the old files silently
foreach ($file in $oldFiles) {
    Remove-Item $file.FullName -Force -ErrorAction SilentlyContinue
}

Get-Volume $folderPath.Substring(0, [Math]::Min($folderPath.Length, 1))
