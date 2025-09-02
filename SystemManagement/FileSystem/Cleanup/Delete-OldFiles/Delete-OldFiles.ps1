# =============================================================================
# Script: Delete-OldFiles.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.10.2
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
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
    - Recursive deletion by default
    - Silent operation with error suppression

    Supports -WhatIf parameter to preview changes without making them.
.PARAMETER StartPath
    The path to the folder containing files to be cleaned up
.PARAMETER daysOld
    Number of days old the files must be to be deleted
.PARAMETER NoRecurse
    Optional switch to disable recursive deletion of files and empty directories
.EXAMPLE
    .\Delete-OldFiles.ps1
    Recursively deletes files older than 30 days from C:\windows\System32\winevt\logs
.EXAMPLE
    .\Delete-OldFiles.ps1 -StartPath "D:\Backups" -daysOld 90 -NoRecurse
    Deletes files older than 90 days from D:\Backups without recursing into subdirectories
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param(
    [Parameter(Mandatory = $false)]
    [string]$StartPath = "C:\windows\System32\winevt\logs",

    [Parameter(Mandatory = $false)]
    [int]$daysOld = 30,

    [Parameter(Mandatory = $false)]
    [switch]$NoRecurse
)

# Color support variables and Write-ColorOutput function
$Script:UseAnsiColors = $PSVersionTable.PSVersion.Major -ge 7
$Script:Colors = if ($Script:UseAnsiColors) {
    @{
        'White'    = "`e[37m"
        'Cyan'     = "`e[36m"
        'Green'    = "`e[32m"
        'Yellow'   = "`e[33m"
        'Red'      = "`e[31m"
        'Magenta'  = "`e[35m"
        'DarkGray' = "`e[90m"
        'Reset'    = "`e[0m"
    }
} else {
    @{
        'White'    = [ConsoleColor]::White
        'Cyan'     = [ConsoleColor]::Cyan
        'Green'    = [ConsoleColor]::Green
        'Yellow'   = [ConsoleColor]::Yellow
        'Red'      = [ConsoleColor]::Red
        'Magenta'  = [ConsoleColor]::Magenta
        'DarkGray' = [ConsoleColor]::DarkGray
        'Reset'    = ''
    }
}

function Write-ColorOutput {
    <#
    .SYNOPSIS
    Outputs colored text in a way that's compatible with PSScriptAnalyzer requirements.

    .DESCRIPTION
    This function provides colored output while maintaining compatibility with PSScriptAnalyzer
    by using only Write-Output and standard PowerShell cmdlets.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$Color = "White"
    )

    # Always use Write-Output to satisfy PSScriptAnalyzer
    # For PowerShell 7+, include ANSI color codes in the output
    if ($Script:UseAnsiColors) {
        $colorCode = $Script:Colors[$Color]
        $resetCode = $Script:Colors.Reset
        Write-Output "${colorCode}${Message}${resetCode}"
    } else {
        # For PowerShell 5.1, just output the message
        # Color formatting will be handled by the terminal/host if supported
        Write-Output $Message
    }
}


function Show-DriveInfo {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Volume
    )

    Write-ColorOutput -Message "`nDrive Volume Details:" -Color 'Green'
    Write-ColorOutput -Message "------------------------" -Color 'Green'
    Write-ColorOutput -Message "Drive Letter: $($Volume.DriveLetter)" -Color 'Cyan'
    Write-ColorOutput -Message "Drive Label: $($Volume.FileSystemLabel)" -Color 'Cyan'
    Write-ColorOutput -Message "File System: $($Volume.FileSystem)" -Color 'Cyan'
    Write-ColorOutput -Message "Drive Type: $($Volume.DriveType)" -Color 'Cyan'
    Write-ColorOutput -Message "Size: $([math]::Round($Volume.Size/1GB, 2)) GB" -Color 'Cyan'
    Write-ColorOutput -Message "Free Space: $([math]::Round($Volume.SizeRemaining/1GB, 2)) GB" -Color 'Cyan'
    Write-ColorOutput -Message "Health Status: $($Volume.HealthStatus)" -Color 'Cyan'
    Write-ColorOutput -Message "" -Color "White"
}

try {
    # Setup logging
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $finalLogPath = Join-Path $PSScriptRoot "OldFilesDeleted_$($env:COMPUTERNAME)_$timestamp.log"
    $tempDeletionLog = Join-Path $PSScriptRoot "DeletedFiles_$($env:COMPUTERNAME)_$timestamp.tmp"
    Start-Transcript -Path $finalLogPath -Force

    # Get the drive letter from the folder path
    $driveLetter = $StartPath.Substring(0, 1)

    # Get volume information before deletion
    Write-ColorOutput -Message "Getting drive information before deletion..." -Color 'Cyan'
    $volumeBefore = Get-Volume -DriveLetter $driveLetter -ErrorAction Stop
    Write-ColorOutput -Message "Drive information before file deletion:" -Color 'Yellow'
    Show-DriveInfo -Volume $volumeBefore

    # Get the current date
    $currentDate = Get-Date

    # Calculate the cutoff date
    $cutoffDate = $currentDate.AddDays(-$daysOld)

    Write-ColorOutput -Message "`nDeleting files$(if(!$NoRecurse) { ' and directories' }) older than $daysOld days..." -Color 'Cyan'

    # Get files to delete based on recursion setting
    $oldFiles = if (!$NoRecurse) {
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

        try {
            if ($PSCmdlet.ShouldProcess($file.FullName, "Delete file")) {
                Remove-Item $file.FullName -Force -ErrorAction Stop
                "Deleted file: $($file.FullName)" | Out-File -FilePath $tempDeletionLog -Append
            }
        } catch {
            "Failed to delete file: $($file.FullName) - Error: $_" | Out-File -FilePath $tempDeletionLog -Append
        }
    }

    # Clear the progress bar
    Write-Progress -Activity "Deleting Old Files" -Completed

    # Only process directories if -NoRecurse is not specified
    if (!$NoRecurse) {
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
                try {
                    if ($PSCmdlet.ShouldProcess($dir.FullName, "Delete empty directory")) {
                        Remove-Item $dir.FullName -Force -ErrorAction Stop
                        "Deleted empty directory: $($dir.FullName)" | Out-File -FilePath $tempDeletionLog -Append
                    }
                } catch {
                    "Failed to delete directory: $($dir.FullName) - Error: $_" | Out-File -FilePath $tempDeletionLog -Append
                }
            }
        }

        # Clear the progress bar
        Write-Progress -Activity "Processing Empty Directories" -Completed
    }

    # Get volume information after deletion
    Write-ColorOutput -Message "`nGetting drive information after deletion..." -Color 'Cyan'
    $volumeAfter = Get-Volume -DriveLetter $driveLetter -ErrorAction Stop
    Write-ColorOutput -Message "Drive information after file deletion:" -Color 'Yellow'
    Show-DriveInfo -Volume $volumeAfter

    # Show space reclaimed
    $spaceReclaimed = $volumeAfter.SizeRemaining - $volumeBefore.SizeRemaining
    if ($spaceReclaimed -gt 0) {
        Write-ColorOutput -Message "`nSpace reclaimed: $([math]::Round($spaceReclaimed/1MB, 2)) MB" -Color 'Green'
    } else {
        Write-ColorOutput -Message "`nNo measurable space was reclaimed." -Color 'Yellow'
    }
} catch {
    Write-Error "Error performing operation. Error: $_"
} finally {
    # Stop logging
    try {
        Stop-Transcript
        if (Test-Path $tempDeletionLog) {
            # Append deletion log to transcript and cleanup
            "`nDetailed Deletion Log:" | Out-File -FilePath $finalLogPath -Append
            Get-Content $tempDeletionLog | Out-File -FilePath $finalLogPath -Append
            Remove-Item $tempDeletionLog -Force
        }
    } catch {
        Write-Error "Failed to stop transcript or merge logs: $_"
    }
}
