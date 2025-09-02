# =============================================================================
# Script: Delete-OldScreenshots.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.3.3
# Additional Info: Fixed header metadata for workflow validation
# =============================================================================

<#
.SYNOPSIS
    Deletes screenshot files older than a specified number of days.
.DESCRIPTION
    This script automatically removes screenshot files from a specified folder
    that are older than a defined threshold (default 30 days).
    - Searches specified screenshots folder
    - Removes files older than threshold
    - Provides progress feedback and logging
    - Handles errors gracefully

    Supports -WhatIf parameter to preview deletions without making them.
.PARAMETER FolderPath
    The path to the screenshots folder
.PARAMETER DaysOld
    Number of days old the files must be before deletion
.EXAMPLE
    .\Delete-OldScreenshots.ps1 -FolderPath "C:\Screenshots" -DaysOld 30
    Deletes all screenshots older than 30 days from the specified folder
.EXAMPLE
    .\Delete-OldScreenshots.ps1 -WhatIf
    Shows which files would be deleted without actually deleting them
.NOTES
    Security Level: Low
    Required Permissions: File system read/write access to screenshots folder
    Validation Requirements: Verify folder path exists before execution
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $false)]
    [ValidateScript({ Test-Path $_ })]
    [string]$FolderPath = "C:\Users\jdyer\OneDrive - Nuvodia\Pictures\Screenshots",

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 365)]
    [int]$DaysOld = 30
)

# Initialize logging
$LogPath = Join-Path $PSScriptRoot "DeleteScreenshots.log"
$ErrorActionPreference = "Stop"

function Write-LogMessage {
    param($Message, $Level = "Information")
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "$TimeStamp [$Level] $Message"
    Add-Content -Path $LogPath -Value $LogMessage
    switch ($Level) {
        "Information" { Write-Output $Message }
        "Success" { Write-Output "[SUCCESS] $Message" }
        "Warning" { Write-Warning $Message }
        "Error" { Write-Error $Message -ErrorAction Continue }
    }
}

try {
    Write-LogMessage "Starting screenshot cleanup process" "Information"
    Write-LogMessage "Target folder: $FolderPath" "Information"
    Write-LogMessage "Deleting files older than $DaysOld days" "Information"

    # Calculate cutoff date
    $cutoffDate = (Get-Date).AddDays(-$DaysOld)

    # Get old files
    $oldFiles = Get-ChildItem -Path $FolderPath -File | Where-Object { $_.LastWriteTime -lt $cutoffDate }
    $totalFiles = $oldFiles.Count

    if ($totalFiles -eq 0) {
        Write-LogMessage "No files found older than $DaysOld days" "Information"
        exit 0
    }

    Write-LogMessage "Found $totalFiles files to delete" "Information"
    $deleted = 0
    $failed = 0

    # Process files with progress bar
    foreach ($file in $oldFiles) {
        $percent = [math]::Round(($deleted + $failed) / $totalFiles * 100)
        Write-Progress -Activity "Deleting old screenshots" -Status "$percent% Complete" -PercentComplete $percent

        try {
            if ($PSCmdlet.ShouldProcess($file.FullName, "Delete screenshot")) {
                Remove-Item $file.FullName -Force
                $deleted++
                Write-LogMessage "Deleted: $($file.Name)" "Success"
            }
        } catch {
            $failed++
            Write-LogMessage "Failed to delete: $($file.Name). Error: $($_.Exception.Message)" "Error"
        }
    }

    Write-Progress -Activity "Deleting old screenshots" -Completed
    Write-LogMessage "Operation complete. Successfully deleted: $deleted files. Failed: $failed files" "Information"
} catch {
    Write-LogMessage "Script execution failed: $($_.Exception.Message)" "Error"
    exit 1
} finally {
    Write-Progress -Activity "Deleting old screenshots" -Completed
    Write-LogMessage "Script execution finished. Cleaning up resources." "Information"
    # Remove any temporary variables that might contain sensitive data
    Remove-Variable -Name oldFiles, file -ErrorAction SilentlyContinue
}
