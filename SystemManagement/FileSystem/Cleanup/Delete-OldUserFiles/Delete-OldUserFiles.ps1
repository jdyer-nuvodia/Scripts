# =============================================================================
# Script: Delete-OldUserFiles.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1.1
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
# =============================================================================

<#
.SYNOPSIS
    Deletes user directories that have been inactive for a specified number of months.

.DESCRIPTION
    This script performs the following actions:
    - Scans all user directories in the specified parent directory (default: C:\Users)
    - Calculates cutoff date based on current date minus specified months
    - Identifies user directories that have no files modified after the cutoff date
    - Deletes inactive user directories and provides detailed logging
    - Shows drive space comparison before and after deletion
    - Supports -WhatIf parameter to preview changes without making them
    - Creates detailed logs of every file and directory deleted

.PARAMETER ParentPath
    The parent path containing user directories to check (default: C:\Users)

.PARAMETER MonthsOld
    Number of months of inactivity before a user directory is considered for deletion (default: 6)

.PARAMETER ExcludeUsers
    Array of additional user directory names to exclude from deletion (case-insensitive).
    These will be added to the default system exclusions: Administrator, Default, Public, All Users, Default User

.EXAMPLE
    .\Delete-OldUserFiles.ps1
    Deletes user directories from C:\Users that have been inactive for 6 months or more

.EXAMPLE
    .\Delete-OldUserFiles.ps1 -ParentPath "D:\UserProfiles" -MonthsOld 12
    Deletes user directories from D:\UserProfiles that have been inactive for 12 months or more

.EXAMPLE
    .\Delete-OldUserFiles.ps1 -ExcludeUsers @("TestUser", "ServiceAccount") -WhatIf
    Shows what would be deleted, excluding TestUser and ServiceAccount in addition to the default system accounts
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $false)]
    [string]$ParentPath = "C:\Users",

    [Parameter(Mandatory = $false)]
    [int]$MonthsOld = 6,

    [Parameter(Mandatory = $false)]
    [string[]]$ExcludeUsers = @()
)

function Show-DriveInfo {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Volume
    )

    Write-Information "`nDrive Volume Details:" -InformationAction Continue
    Write-Information "------------------------" -InformationAction Continue
    Write-Information "Drive Letter: $($Volume.DriveLetter)" -InformationAction Continue
    Write-Information "Drive Label: $($Volume.FileSystemLabel)" -InformationAction Continue
    Write-Information "File System: $($Volume.FileSystem)" -InformationAction Continue
    Write-Information "Drive Type: $($Volume.DriveType)" -InformationAction Continue
    Write-Information "Size: $([math]::Round($Volume.Size/1GB, 2)) GB" -InformationAction Continue
    Write-Information "Free Space: $([math]::Round($Volume.SizeRemaining/1GB, 2)) GB" -InformationAction Continue
    Write-Information "Health Status: $($Volume.HealthStatus)" -InformationAction Continue
    Write-Information "" -InformationAction Continue
}

function Write-LogEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $true)]
        [string]$LogPath,

        [Parameter(Mandatory = $false)]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC"
    $logEntry = "[$timestamp] [$Level] $Message"
    $logEntry | Out-File -FilePath $LogPath -Append -Encoding UTF8

    switch ($Level) {
        "ERROR" { Write-Error $Message }
        "WARNING" { Write-Warning $Message }
        "SUCCESS" { Write-Information $Message -InformationAction Continue }
        default { Write-Verbose $Message }
    }
}

try {
    # Validate parent path exists
    if (-not (Test-Path -Path $ParentPath -PathType Container)) {
        throw "Parent path '$ParentPath' does not exist or is not a directory."
    }

    # Combine system exclusions with user-specified exclusions
    $systemExclusions = @("Administrator", "Default", "Public", "All Users", "Default User", "nuvodialocal")
    $allExcludedUsers = ($systemExclusions + $ExcludeUsers) | Select-Object -Unique

    # Setup logging
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $logFileName = "DeleteOldUserFiles_$($env:COMPUTERNAME)_$timestamp.log"
    $finalLogPath = Join-Path -Path $PSScriptRoot -ChildPath $logFileName
    $detailLogPath = Join-Path -Path $PSScriptRoot -ChildPath "DeletedUserFilesDetail_$($env:COMPUTERNAME)_$timestamp.log"

    # Start transcript
    Start-Transcript -Path $finalLogPath -Force

    Write-LogEntry -Message "Starting Delete-OldUserFiles script execution" -LogPath $detailLogPath -Level "INFO"
    Write-LogEntry -Message "Parent Path: $ParentPath" -LogPath $detailLogPath -Level "INFO"
    Write-LogEntry -Message "Months Old Threshold: $MonthsOld" -LogPath $detailLogPath -Level "INFO"
    Write-LogEntry -Message "System Excluded Users: $($systemExclusions -join ', ')" -LogPath $detailLogPath -Level "INFO"
    Write-LogEntry -Message "Additional Excluded Users: $($ExcludeUsers -join ', ')" -LogPath $detailLogPath -Level "INFO"
    Write-LogEntry -Message "All Excluded Users: $($allExcludedUsers -join ', ')" -LogPath $detailLogPath -Level "INFO"

    # Get the drive letter from the parent path
    $driveLetter = Split-Path -Path $ParentPath -Qualifier
    $driveLetterOnly = $driveLetter.TrimEnd(':')

    # Get volume information before deletion
    Write-Information "Getting drive information before deletion..." -InformationAction Continue
    $volumeBefore = Get-Volume -DriveLetter $driveLetterOnly -ErrorAction Stop
    Write-Information "Drive information before user directory deletion:" -InformationAction Continue
    Show-DriveInfo -Volume $volumeBefore

    # Calculate the cutoff date
    $cutoffDate = (Get-Date).AddMonths(-$MonthsOld)
    Write-LogEntry -Message "Cutoff date calculated: $cutoffDate" -LogPath $detailLogPath -Level "INFO"

    Write-Information "`nScanning for inactive user directories older than $MonthsOld months (before $cutoffDate)..." -InformationAction Continue

    # Get all user directories
    $userDirectories = Get-ChildItem -Path $ParentPath -Directory -ErrorAction SilentlyContinue

    if (-not $userDirectories) {
        Write-LogEntry -Message "No user directories found in $ParentPath" -LogPath $detailLogPath -Level "WARNING"
        return
    }

    Write-LogEntry -Message "Found $($userDirectories.Count) user directories to analyze" -LogPath $detailLogPath -Level "INFO"

    $inactiveDirectories = @()
    $totalDirs = $userDirectories.Count
    $currentDir = 0

    foreach ($userDir in $userDirectories) {
        $currentDir++
        $percentComplete = [math]::Round(($currentDir / $totalDirs) * 100, 2)

        Write-Progress -Activity "Analyzing User Directories" `
            -Status "Checking directory $currentDir of $totalDirs ($percentComplete%)" `
            -PercentComplete $percentComplete `
            -CurrentOperation $userDir.Name

        # Skip excluded users (case-insensitive)
        if ($allExcludedUsers -contains $userDir.Name) {
            Write-LogEntry -Message "Skipping excluded user directory: $($userDir.Name)" -LogPath $detailLogPath -Level "INFO"
            continue
        }

        try {
            # Check for any files/folders modified after cutoff date
            Write-LogEntry -Message "Analyzing user directory: $($userDir.FullName)" -LogPath $detailLogPath -Level "INFO"

            $recentItems = Get-ChildItem -Path $userDir.FullName -Recurse -Force -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -ge $cutoffDate }

            if (-not $recentItems) {
                # No items newer than cutoff, directory is inactive
                Write-LogEntry -Message "User directory '$($userDir.Name)' is inactive (no files modified since $cutoffDate)" -LogPath $detailLogPath -Level "INFO"
                $inactiveDirectories += $userDir
            } else {
                Write-LogEntry -Message "User directory '$($userDir.Name)' is active (has $($recentItems.Count) items modified since $cutoffDate)" -LogPath $detailLogPath -Level "INFO"
            }
        } catch {
            Write-LogEntry -Message "Error analyzing user directory '$($userDir.Name)': $_" -LogPath $detailLogPath -Level "ERROR"
        }
    }

    # Clear the progress bar
    Write-Progress -Activity "Analyzing User Directories" -Completed

    if ($inactiveDirectories.Count -eq 0) {
        Write-Information "`nNo inactive user directories found." -InformationAction Continue
        Write-LogEntry -Message "No inactive user directories found for deletion" -LogPath $detailLogPath -Level "INFO"
        return
    }

    Write-Information "`nFound $($inactiveDirectories.Count) inactive user directories for deletion:" -InformationAction Continue
    foreach ($dir in $inactiveDirectories) {
        Write-Information "  - $($dir.Name) (Last Modified: $($dir.LastWriteTime))" -InformationAction Continue
    }

    # Process inactive directories for deletion
    $totalInactive = $inactiveDirectories.Count
    $currentInactive = 0
    $deletedCount = 0
    $errorCount = 0

    foreach ($inactiveDir in $inactiveDirectories) {
        $currentInactive++
        $percentComplete = [math]::Round(($currentInactive / $totalInactive) * 100, 2)

        Write-Progress -Activity "Processing Inactive User Directories" `
            -Status "Processing $currentInactive of $totalInactive directories ($percentComplete%)" `
            -PercentComplete $percentComplete `
            -CurrentOperation $inactiveDir.Name

        try {
            if ($PSCmdlet.ShouldProcess($inactiveDir.FullName, "Delete inactive user directory and all contents")) {

                # Log all files and folders that will be deleted
                Write-LogEntry -Message "Beginning deletion of user directory: $($inactiveDir.FullName)" -LogPath $detailLogPath -Level "INFO"

                $allItems = Get-ChildItem -Path $inactiveDir.FullName -Recurse -Force -ErrorAction SilentlyContinue
                foreach ($item in $allItems) {
                    if ($item.PSIsContainer) {
                        Write-LogEntry -Message "Will delete directory: $($item.FullName)" -LogPath $detailLogPath -Level "INFO"
                    } else {
                        Write-LogEntry -Message "Will delete file: $($item.FullName) (Size: $($item.Length) bytes, Last Modified: $($item.LastWriteTime))" -LogPath $detailLogPath -Level "INFO"
                    }
                }

                # Delete the entire user directory
                Remove-Item -Path $inactiveDir.FullName -Recurse -Force -ErrorAction Stop

                Write-LogEntry -Message "Successfully deleted user directory: $($inactiveDir.FullName)" -LogPath $detailLogPath -Level "SUCCESS"
                $deletedCount++
            } else {
                Write-LogEntry -Message "WhatIf: Would delete user directory: $($inactiveDir.FullName)" -LogPath $detailLogPath -Level "INFO"
            }
        } catch {
            Write-LogEntry -Message "Failed to delete user directory '$($inactiveDir.FullName)': $_" -LogPath $detailLogPath -Level "ERROR"
            $errorCount++
        }
    }

    # Clear the progress bar
    Write-Progress -Activity "Processing Inactive User Directories" -Completed

    # Summary
    Write-Information "`n=== DELETION SUMMARY ===" -InformationAction Continue
    Write-Information "Total inactive directories found: $totalInactive" -InformationAction Continue
    Write-Information "Successfully deleted: $deletedCount" -InformationAction Continue
    Write-Information "Errors encountered: $errorCount" -InformationAction Continue

    Write-LogEntry -Message "Deletion Summary - Total: $totalInactive, Deleted: $deletedCount, Errors: $errorCount" -LogPath $detailLogPath -Level "INFO"

    # Get volume information after deletion
    if ($deletedCount -gt 0) {
        Write-Information "`nGetting drive information after deletion..." -InformationAction Continue
        $volumeAfter = Get-Volume -DriveLetter $driveLetterOnly -ErrorAction Stop
        Write-Information "Drive information after user directory deletion:" -InformationAction Continue
        Show-DriveInfo -Volume $volumeAfter

        # Show space reclaimed
        $spaceReclaimed = $volumeAfter.SizeRemaining - $volumeBefore.SizeRemaining
        if ($spaceReclaimed -gt 0) {
            $spaceReclaimedMB = [math]::Round($spaceReclaimed / 1MB, 2)
            $spaceReclaimedGB = [math]::Round($spaceReclaimed / 1GB, 2)
            Write-Information "`nSpace reclaimed: $spaceReclaimedMB MB ($spaceReclaimedGB GB)" -InformationAction Continue
            Write-LogEntry -Message "Space reclaimed: $spaceReclaimedMB MB ($spaceReclaimedGB GB)" -LogPath $detailLogPath -Level "SUCCESS"
        } else {
            Write-Information "`nNo measurable space was reclaimed." -InformationAction Continue
            Write-LogEntry -Message "No measurable space was reclaimed" -LogPath $detailLogPath -Level "WARNING"
        }
    }

    Write-LogEntry -Message "Delete-OldUserFiles script execution completed successfully" -LogPath $detailLogPath -Level "SUCCESS"
} catch {
    $errorMessage = "Critical error during script execution: $_"
    Write-Error $errorMessage
    Write-LogEntry -Message $errorMessage -LogPath $detailLogPath -Level "ERROR"
} finally {
    # Stop logging and merge detailed log
    try {
        # Only stop transcript if one is actually running
        try {
            $transcriptRunning = $true
            Stop-Transcript -ErrorAction Stop
        } catch {
            $transcriptRunning = $false
        }

        if (Test-Path -Path $detailLogPath) {
            # Append detailed log to transcript if it exists, otherwise create final log
            if ($transcriptRunning -or (Test-Path -Path $finalLogPath)) {
                "`n`n=== DETAILED OPERATION LOG ===" | Out-File -FilePath $finalLogPath -Append -Encoding UTF8
                Get-Content -Path $detailLogPath | Out-File -FilePath $finalLogPath -Append -Encoding UTF8
            } else {
                # If no transcript was created, rename detail log as final log
                Move-Item -Path $detailLogPath -Destination $finalLogPath -Force -ErrorAction SilentlyContinue
            }

            # Clean up temporary detailed log if it still exists
            if (Test-Path -Path $detailLogPath) {
                Remove-Item -Path $detailLogPath -Force -ErrorAction SilentlyContinue
            }
        }

        if (Test-Path -Path $finalLogPath) {
            Write-Information "`nDetailed log saved to: $finalLogPath" -InformationAction Continue
        }
    } catch {
        Write-Warning "Failed to finalize logging: $_"
        # Ensure we still inform about the detail log location if it exists
        if (Test-Path -Path $detailLogPath) {
            Write-Information "`nDetail log available at: $detailLogPath" -InformationAction Continue
        }
    }
}
