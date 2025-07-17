# =============================================================================
# Script: Delete-AllFilesInDirectory.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.3.6
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
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

    Supports -WhatIf parameter to preview changes without making them.

    Dependencies:
    - PowerShell 5.1 or higher
    - Appropriate permissions on target directory
    - Administrative rights (for taking ownership)
.PARAMETER TargetPath
    The target directory path to clean up. This parameter is mandatory.
.EXAMPLE
    .\Delete-AllFilesInDirectory.ps1 -TargetPath "C:\TempFiles"
    Deletes all contents in the specified directory "C:\TempFiles"
.EXAMPLE
    .\Delete-AllFilesInDirectory.ps1 -TargetPath "C:\TempFiles" -WhatIf
    Shows what files and folders would be deleted without actually deleting them
.NOTES
    Security Level: High
    Required Permissions: Write access to target directory, Administrative rights
    Validation Requirements: Verify target directory before execution
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $true)]
    [string]$TargetPath
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


# Function to take ownership of files and folders
function Set-Ownership {
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([System.Boolean])]
    param (
        [Parameter(Mandatory = $true)]
        [string]$StartPath
    )

    if ($PSCmdlet.ShouldProcess($StartPath, "Take ownership of path and set permissions")) {
        Write-ColorOutput -Message "Taking ownership of path: $StartPath" -Color 'Cyan'
        try {
            # Take ownership using takeown command
            $takeOwnResult = Start-Process -FilePath "takeown.exe" -ArgumentList "/F `"$StartPath`" /R /D Y" -NoNewWindow -PassThru -Wait
            if ($takeOwnResult.ExitCode -ne 0) {
                Write-ColorOutput -Message "Warning: Failed to take ownership of $StartPath (Exit code: $($takeOwnResult.ExitCode))" -Color 'Yellow'
            }

            # Grant full control to the current user/system
            $grantResult = Start-Process -FilePath "icacls.exe" -ArgumentList "`"$StartPath`" /grant *S-1-5-18:F /T /C /Q" -NoNewWindow -PassThru -Wait
            if ($grantResult.ExitCode -ne 0) {
                Write-ColorOutput -Message "Warning: Failed to grant permissions on $StartPath (Exit code: $($grantResult.ExitCode))" -Color 'Yellow'
            }

            return $true
        } catch {
            Write-ColorOutput -Message "Error taking ownership of $StartPath`: $($_.Exception.Message)" -Color 'Yellow'
            return $false
        }
    } else {
        Write-ColorOutput -Message "Would take ownership of path: $StartPath" -Color 'Cyan'
        return $true
    }
}

Write-ColorOutput -Message "Starting directory cleanup process..." -Color 'Cyan'
Write-ColorOutput -Message "Target directory: $TargetPath" -Color 'Cyan'

try {
    # Step 1: Take ownership of the target directory and all contents
    Write-ColorOutput -Message "Taking ownership of all files and folders..." -Color 'Cyan'
    $ownershipResult = Set-Ownership -StartPath $script:TargetPath -WhatIf:$WhatIfPreference -Confirm:$ConfirmPreference

    if (-not $ownershipResult) {
        Write-ColorOutput -Message "Continuing with deletion despite ownership issues. Some files may be skipped." -Color 'Yellow'
    }

    # Remove all files
    Write-ColorOutput -Message "Removing files..." -Color 'Cyan'
    Get-ChildItem -Path $TargetPath -File -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            if ($PSCmdlet.ShouldProcess($_.FullName, "Delete file")) {
                Write-ColorOutput -Message "Deleting file: $($_.FullName)" -Color 'Yellow'
                Remove-Item -Path $_.FullName -Force -ErrorAction Stop
            }
        } catch {
            Write-ColorOutput -Message "Failed to delete file $($_.FullName): $($_.Exception.Message)" -Color 'Yellow'
        }
    }

    # Remove all folders with improved error handling and long path support
    Write-ColorOutput -Message "Removing folders..." -Color 'Cyan'
    Get-ChildItem -Path $TargetPath -Directory -Recurse -ErrorAction SilentlyContinue |
        Sort-Object -Property FullName -Descending |
        ForEach-Object {
            $StartPath = $_.FullName
            if ($PSCmdlet.ShouldProcess($StartPath, "Delete folder")) {
                Write-ColorOutput -Message "Attempting to delete folder: $StartPath" -Color 'Yellow'
                try {
                    # Enable long path support if needed
                    if ($StartPath.Length -ge 260) {
                        $StartPath = "\\?\$StartPath"
                        Write-ColorOutput -Message "Using long path format: $StartPath" -Color 'Yellow'
                    }

                    # Try up to 3 times with a small delay between attempts
                    $maxAttempts = 3
                    $attempt = 1
                    $success = $false

                    while (-not $success -and $attempt -le $maxAttempts) {
                        try {
                            Remove-Item -LiteralPath $StartPath -Recurse -Force -ErrorAction Stop
                            $success = $true
                            Write-ColorOutput -Message "Successfully deleted folder: $($_.FullName)" -Color 'Green'
                        } catch {
                            if ($attempt -lt $maxAttempts) {
                                Write-ColorOutput -Message "Attempt $attempt failed, retrying in 2 seconds..." -Color 'Yellow'
                                Start-Sleep -Seconds 2
                                $attempt++
                            } else {
                                Write-ColorOutput -Message "Failed to delete folder after $maxAttempts attempts: $($_.FullName)" -Color 'Yellow'
                                Write-ColorOutput -Message "Error: $($_.Exception.Message)" -Color 'Yellow'
                                # Continue with other folders instead of stopping
                            }
                        }
                    }
                } catch {
                    Write-ColorOutput -Message "Error processing folder $($_.FullName): $($_.Exception.Message)" -Color 'Yellow'
                    # Continue with other folders
                }
            }
        }

    Write-ColorOutput -Message "Directory cleanup completed successfully!" -Color 'Green'
} catch {
    Write-Error "An error occurred during the cleanup process: $_"
    exit 1
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

    Write-ColorOutput -Message "Found lowest drive letter: $($lowestVolume.DriveLetter)" -Color 'Yellow'
    Show-DriveInfo -Volume $lowestVolume
} catch {
    Write-Error "Error accessing drive information. Error: $_"
}
