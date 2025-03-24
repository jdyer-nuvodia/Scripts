# =============================================================================
# Script: Delete-OldScreenshots.ps1
# Created: 2024-02-07 13:45:00 UTC
# Author: nunya-nunya
# Last Updated: 2025-02-26 23:45:00 UTC
# Updated By: nunya-nunya
# Version: 1.0
# Additional Info: Initial script creation with standard header format
# =============================================================================

<#
.SYNOPSIS
    Deletes screenshot files older than a specified number of days.
.DESCRIPTION
    This script automatically removes screenshot files from a specified folder
    that are older than a defined threshold (default 30 days).
    - Searches specified screenshots folder
    - Removes files older than threshold
    - Runs silently without user interaction
.PARAMETER StartPath
    The path to the screenshots folder
.PARAMETER daysOld
    Number of days old the files must be before deletion
.EXAMPLE
    .\Delete-OldScreenshots.ps1
    Deletes all screenshots older than 30 days from the specified folder
.NOTES
    Security Level: Low
    Required Permissions: File system read/write access to screenshots folder
    Validation Requirements: Verify folder path exists before execution
#>

# Set the folder path
$StartPath = "C:\Users\nunya\OneDrive - nunya\Pictures\Screenshots"

# Set the number of days old for files to be deleted
$daysOld = 30

# Get the current date
$currentDate = Get-Date

# Calculate the cutoff date
$cutoffDate = $currentDate.AddDays(-$daysOld)

# Get all files in the folder older than the cutoff date
$oldFiles = Get-ChildItem -Path $StartPath -File | Where-Object { $_.LastWriteTime -lt $cutoffDate }

# Delete the old files silently
foreach ($file in $oldFiles) {
    Remove-Item $file.FullName -Force -ErrorAction SilentlyContinue
}

# Exit silently
exit
