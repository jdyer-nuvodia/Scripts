# =============================================================================
# Script: Delete-AllFilesInDirectory.ps1
# Created: 2024-02-20 17:15:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-06-14 14:30:21 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2.1
# Additional Info: Fixed syntax error with missing closing brace
# =============================================================================

<#
.SYNOPSIS
    Recursively deletes all files and folders in a specified directory.
.DESCRIPTION
    This script removes all files and folders within a specified directory.
    It performs the deletion in three steps:
    1. Takes ownership of all files and folders recursively
    2. Removes all files recursively
    3. Removes all folders in descending order to handle nested directories
    
    Dependencies:
    - PowerShell 5.1 or higher
    - Appropriate permissions on target directory
    - Administrative rights (for taking ownership)
.PARAMETER TargetPath
    The target directory path to clean up. This parameter is mandatory.
.EXAMPLE
    .\Delete-AllFilesInDirectory.ps1 -TargetPath "C:\TempFiles"
    Deletes all contents in the specified directory "C:\TempFiles"
.NOTES
    Security Level: High
    Required Permissions: Write access to target directory, Administrative rights
    Validation Requirements: Verify target directory before execution
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$TargetPath
)

# Function to take ownership of files and folders
function Set-Ownership {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    
    Write-Host "Taking ownership of path: $Path" -ForegroundColor Cyan
    
    try {
        # Take ownership using icacls command
        $takeOwnResult = Start-Process -FilePath "icacls.exe" -ArgumentList "`"$Path`" /takeown /T /C /Q" -NoNewWindow -PassThru -Wait
        if ($takeOwnResult.ExitCode -ne 0) {
            Write-Host "Warning: Failed to take ownership of $Path (Exit code: $($takeOwnResult.ExitCode))" -ForegroundColor Yellow
        }
        
        # Grant full control to the current user/system
        $grantResult = Start-Process -FilePath "icacls.exe" -ArgumentList "`"$Path`" /grant *S-1-5-18:F /T /C /Q" -NoNewWindow -PassThru -Wait
        if ($grantResult.ExitCode -ne 0) {
            Write-Host "Warning: Failed to grant permissions on $Path (Exit code: $($grantResult.ExitCode))" -ForegroundColor Yellow
        }
        
        return $true
    }
    catch {
        Write-Host "Error taking ownership of $Path`: $($_.Exception.Message)" -ForegroundColor Yellow
        return $false
    }
}

Write-Host "Starting directory cleanup process..." -ForegroundColor Cyan
Write-Host "Target directory: $TargetPath" -ForegroundColor Cyan

try {
    # Step 1: Take ownership of the target directory and all contents
    Write-Host "Taking ownership of all files and folders..." -ForegroundColor Cyan
    $ownershipResult = Set-Ownership -Path $TargetPath
    
    if (-not $ownershipResult) {
        Write-Host "Continuing with deletion despite ownership issues. Some files may be skipped." -ForegroundColor Yellow
    }
    
    # Remove all files
    Write-Host "Removing files..." -ForegroundColor Cyan
    Get-ChildItem -Path $TargetPath -File -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            Write-Host "Deleting file: $($_.FullName)" -ForegroundColor Yellow
            Remove-Item -Path $_.FullName -Force -ErrorAction Stop
        }
        catch {
            Write-Host "Failed to delete file $($_.FullName): $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    # Remove all folders with improved error handling and long path support
    Write-Host "Removing folders..." -ForegroundColor Cyan
    Get-ChildItem -Path $TargetPath -Directory -Recurse -ErrorAction SilentlyContinue | 
        Sort-Object -Property FullName -Descending | 
        ForEach-Object {
            $StartPath = $_.FullName
            Write-Host "Attempting to delete folder: $StartPath" -ForegroundColor Yellow
            try {
                # Enable long path support if needed
                if ($StartPath.Length -ge 260) {
                    $StartPath = "\\?\$StartPath"
                    Write-Host "Using long path format: $StartPath" -ForegroundColor Yellow
                }
                
                # Try up to 3 times with a small delay between attempts
                $maxAttempts = 3
                $attempt = 1
                $success = $false
                
                while (-not $success -and $attempt -le $maxAttempts) {
                    try {
                        Remove-Item -LiteralPath $StartPath -Recurse -Force -ErrorAction Stop
                        $success = $true
                        Write-Host "Successfully deleted folder: $($_.FullName)" -ForegroundColor Green
                    }
                    catch {
                        if ($attempt -lt $maxAttempts) {
                            Write-Host "Attempt $attempt failed, retrying in 2 seconds..." -ForegroundColor Yellow
                            Start-Sleep -Seconds 2
                            $attempt++
                        }
                        else {
                            Write-Host "Failed to delete folder after $maxAttempts attempts: $($_.FullName)" -ForegroundColor Yellow
                            Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Yellow
                            # Continue with other folders instead of stopping
                        }
                    }
                }
            }
            catch {
                Write-Host "Error processing folder $($_.FullName): $($_.Exception.Message)" -ForegroundColor Yellow
                # Continue with other folders
            }
        }

    Write-Host "Directory cleanup completed successfully!" -ForegroundColor Green
} 
catch {
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