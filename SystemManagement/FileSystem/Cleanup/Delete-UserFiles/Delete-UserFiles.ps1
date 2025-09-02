# =============================================================================
# Script: Delete-UserFiles.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.1
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
# =============================================================================

<#
.SYNOPSIS
    Deletes specific user directories by inclusion or exclusion criteria.

.DESCRIPTION
    This script performs the following actions:
    - Scans all user directories in the specified parent directory (default: C:\Users)
    - Deletes user directories based on inclusion or exclusion criteria
    - Always protects system accounts: Administrator, Default, Public, All Users, Default User
    - Shows drive space comparison before and after deletion
    - Supports -WhatIf parameter to preview changes without making them
    - Creates detailed logs of every file and directory deleted

.PARAMETER ParentPath
    The parent path containing user directories to check (default: C:\Users)

.PARAMETER IncludeUsers
    Array of specific user directory names to include for deletion (case-insensitive).
    Only these users will be considered for deletion. Cannot be used with ExcludeUsers.

.PARAMETER ExcludeUsers
    Array of additional user directory names to exclude from deletion (case-insensitive).
    All users except these will be considered for deletion. System accounts are always excluded.
    Cannot be used with IncludeUsers.

.EXAMPLE
    .\Delete-UserFiles.ps1 -IncludeUsers @("OldUser1", "TempUser2") -WhatIf
    Shows what would be deleted for only OldUser1 and TempUser2 directories

.EXAMPLE
    .\Delete-UserFiles.ps1 -ExcludeUsers @("CurrentUser", "ActiveUser") -WhatIf
    Shows what would be deleted for all users except CurrentUser and ActiveUser (and system accounts)

.EXAMPLE
    .\Delete-UserFiles.ps1 -ParentPath "D:\UserProfiles" -IncludeUsers @("TestUser")
    Deletes only the TestUser directory from D:\UserProfiles
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $false)]
    [string]$ParentPath = "C:\Users",

    [Parameter(Mandatory = $false, ParameterSetName = 'Include')]
    [string[]]$IncludeUsers = @(),

    [Parameter(Mandatory = $false, ParameterSetName = 'Exclude')]
    [string[]]$ExcludeUsers = @()
)

# Validate that at least one parameter set is used
if ($IncludeUsers.Count -eq 0 -and $ExcludeUsers.Count -eq 0) {
    throw "You must specify either -IncludeUsers or -ExcludeUsers parameter with at least one username."
}

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

    # Define system exclusions that are always protected
    $systemExclusions = @("Administrator", "Default", "Public", "All Users", "Default User")

    # Setup logging
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $logFileName = "DeleteUserFiles_$($env:COMPUTERNAME)_$timestamp.log"
    $finalLogPath = Join-Path -Path $PSScriptRoot -ChildPath $logFileName
    $detailLogPath = Join-Path -Path $PSScriptRoot -ChildPath "DeletedUserFilesDetail_$($env:COMPUTERNAME)_$timestamp.log"

    # Start transcript
    Start-Transcript -Path $finalLogPath -Force

    Write-LogEntry -Message "Starting Delete-UserFiles script execution" -LogPath $detailLogPath -Level "INFO"
    Write-LogEntry -Message "Parent Path: $ParentPath" -LogPath $detailLogPath -Level "INFO"
    Write-LogEntry -Message "System Excluded Users: $($systemExclusions -join ', ')" -LogPath $detailLogPath -Level "INFO"

    if ($IncludeUsers.Count -gt 0) {
        Write-LogEntry -Message "Operation Mode: Include specific users" -LogPath $detailLogPath -Level "INFO"
        Write-LogEntry -Message "Include Users: $($IncludeUsers -join ', ')" -LogPath $detailLogPath -Level "INFO"
    } else {
        Write-LogEntry -Message "Operation Mode: Exclude specific users" -LogPath $detailLogPath -Level "INFO"
        Write-LogEntry -Message "Additional Excluded Users: $($ExcludeUsers -join ', ')" -LogPath $detailLogPath -Level "INFO"
        $allExcludedUsers = ($systemExclusions + $ExcludeUsers) | Select-Object -Unique
        Write-LogEntry -Message "All Excluded Users: $($allExcludedUsers -join ', ')" -LogPath $detailLogPath -Level "INFO"
    }

    # Get the drive letter from the parent path
    $driveLetter = Split-Path -Path $ParentPath -Qualifier
    $driveLetterOnly = $driveLetter.TrimEnd(':')

    # Get volume information before deletion
    Write-Information "Getting drive information before deletion..." -InformationAction Continue
    $volumeBefore = Get-Volume -DriveLetter $driveLetterOnly -ErrorAction Stop
    Write-Information "Drive information before user directory deletion:" -InformationAction Continue
    Show-DriveInfo -Volume $volumeBefore

    Write-Information "`nScanning for user directories to delete..." -InformationAction Continue

    # Get all user directories
    $userDirectories = Get-ChildItem -Path $ParentPath -Directory -ErrorAction SilentlyContinue

    if (-not $userDirectories) {
        Write-LogEntry -Message "No user directories found in $ParentPath" -LogPath $detailLogPath -Level "WARNING"
        return
    }

    Write-LogEntry -Message "Found $($userDirectories.Count) user directories to analyze" -LogPath $detailLogPath -Level "INFO"

    $targetDirectories = @()
    $totalDirs = $userDirectories.Count
    $currentDir = 0

    foreach ($userDir in $userDirectories) {
        $currentDir++
        $percentComplete = [math]::Round(($currentDir / $totalDirs) * 100, 2)

        Write-Progress -Activity "Analyzing User Directories" `
            -Status "Checking directory $currentDir of $totalDirs ($percentComplete%)" `
            -PercentComplete $percentComplete `
            -CurrentOperation $userDir.Name

        # Always skip system exclusions
        if ($systemExclusions -contains $userDir.Name) {
            Write-LogEntry -Message "Skipping system protected user directory: $($userDir.Name)" -LogPath $detailLogPath -Level "INFO"
            continue
        }

        $shouldDelete = $false

        if ($IncludeUsers.Count -gt 0) {
            # Include mode: only delete users in the include list
            if ($IncludeUsers -contains $userDir.Name) {
                $shouldDelete = $true
                Write-LogEntry -Message "User directory '$($userDir.Name)' is in include list - marked for deletion" -LogPath $detailLogPath -Level "INFO"
            } else {
                Write-LogEntry -Message "User directory '$($userDir.Name)' is not in include list - skipping" -LogPath $detailLogPath -Level "INFO"
            }
        } else {
            # Exclude mode: delete all users except those in the exclude list
            if ($allExcludedUsers -contains $userDir.Name) {
                Write-LogEntry -Message "Skipping excluded user directory: $($userDir.Name)" -LogPath $detailLogPath -Level "INFO"
            } else {
                $shouldDelete = $true
                Write-LogEntry -Message "User directory '$($userDir.Name)' is not excluded - marked for deletion" -LogPath $detailLogPath -Level "INFO"
            }
        }

        if ($shouldDelete) {
            $targetDirectories += $userDir
        }
    }

    # Clear the progress bar
    Write-Progress -Activity "Analyzing User Directories" -Completed

    if ($targetDirectories.Count -eq 0) {
        Write-Information "`nNo user directories found matching the specified criteria." -InformationAction Continue
        Write-LogEntry -Message "No user directories found matching the specified criteria for deletion" -LogPath $detailLogPath -Level "INFO"
        return
    }

    Write-Information "`nFound $($targetDirectories.Count) user directories for deletion:" -InformationAction Continue
    foreach ($dir in $targetDirectories) {
        Write-Information "  - $($dir.Name) (Last Modified: $($dir.LastWriteTime))" -InformationAction Continue
    }

    # Process target directories for deletion
    $totalTarget = $targetDirectories.Count
    $currentTarget = 0
    $deletedCount = 0
    $errorCount = 0

    foreach ($targetDir in $targetDirectories) {
        $currentTarget++
        $percentComplete = [math]::Round(($currentTarget / $totalTarget) * 100, 2)

        Write-Progress -Activity "Processing User Directories for Deletion" `
            -Status "Processing $currentTarget of $totalTarget directories ($percentComplete%)" `
            -PercentComplete $percentComplete `
            -CurrentOperation $targetDir.Name

        try {
            if ($PSCmdlet.ShouldProcess($targetDir.FullName, "Delete user directory and all contents")) {

                # Log all files and folders that will be deleted
                Write-LogEntry -Message "Beginning deletion of user directory: $($targetDir.FullName)" -LogPath $detailLogPath -Level "INFO"

                $allItems = Get-ChildItem -Path $targetDir.FullName -Recurse -Force -ErrorAction SilentlyContinue
                foreach ($item in $allItems) {
                    if ($item.PSIsContainer) {
                        Write-LogEntry -Message "Will delete directory: $($item.FullName)" -LogPath $detailLogPath -Level "INFO"
                    } else {
                        Write-LogEntry -Message "Will delete file: $($item.FullName) (Size: $($item.Length) bytes, Last Modified: $($item.LastWriteTime))" -LogPath $detailLogPath -Level "INFO"
                    }
                }

                # Delete the entire user directory
                Remove-Item -Path $targetDir.FullName -Recurse -Force -ErrorAction Stop

                Write-LogEntry -Message "Successfully deleted user directory: $($targetDir.FullName)" -LogPath $detailLogPath -Level "SUCCESS"
                $deletedCount++
            } else {
                Write-LogEntry -Message "WhatIf: Would delete user directory: $($targetDir.FullName)" -LogPath $detailLogPath -Level "INFO"
            }
        } catch {
            Write-LogEntry -Message "Failed to delete user directory '$($targetDir.FullName)': $_" -LogPath $detailLogPath -Level "ERROR"
            $errorCount++
        }
    }

    # Clear the progress bar
    Write-Progress -Activity "Processing User Directories for Deletion" -Completed

    # Summary
    Write-Information "`n=== DELETION SUMMARY ===" -InformationAction Continue
    Write-Information "Total directories found for deletion: $totalTarget" -InformationAction Continue
    Write-Information "Successfully deleted: $deletedCount" -InformationAction Continue
    Write-Information "Errors encountered: $errorCount" -InformationAction Continue

    Write-LogEntry -Message "Deletion Summary - Total: $totalTarget, Deleted: $deletedCount, Errors: $errorCount" -LogPath $detailLogPath -Level "INFO"

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

    Write-LogEntry -Message "Delete-UserFiles script execution completed successfully" -LogPath $detailLogPath -Level "SUCCESS"
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
