# =============================================================================
# Script: Get-FolderSizes.ps1
# Created: 2025-02-05 00:55:03 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-06-11 22:15:00 UTC
# Updated By: jdyer-nuvodia
# Version: 2.10.1
# Additional Info: Fixed PSScriptAnalyzer issues including trailing whitespace and FollowJunctions parameter implementation
# =============================================================================

# Requires -Version 5.1

<#
.SYNOPSIS
    Ultra-fast directory scanner that analyzes folder sizes and identifies largest files.

.DESCRIPTION
    This script performs a high-performance recursive directory scan using
    optimized .NET methods for maximum performance, even when scanning system directories.

    Features:
    - Handles access-denied errors gracefully
    - Identifies largest files in each directory
    - Uses ANSI color codes for rich console output while remaining PSScriptAnalyzer compliant

.PARAMETER StartPath
    The starting directory path to analyze. Default is C:\.

.PARAMETER MaxDepth
    Maximum recursion depth for folder scanning. Default is 10.

.PARAMETER Top
    Number of top folders to display at each level. Default is 3.

.PARAMETER IncludeHiddenSystem
    Whether to include hidden and system files/folders in the analysis. Default is $true.

.PARAMETER FollowJunctions
    Whether to follow NTFS junction points and symbolic links. Default is $true.

.PARAMETER MaxThreads
    Maximum number of concurrent threads for parallel processing. Default is 10.

.PARAMETER OnlyPhysicalFiles
    When set to True, only count files that are physically stored on disk (not cloud placeholders). Default is $true.

.EXAMPLE
    .\Get-FolderSizes.ps1
    Analyzes all folders starting at C:\ with default settings.

.EXAMPLE
    .\Get-FolderSizes.ps1 -StartPath "D:\Projects" -MaxDepth 3 -Top 5
    Analyzes the D:\Projects folder, going 3 levels deep and showing the top 5 largest folders at each level.

.EXAMPLE
    .\Get-FolderSizes.ps1 -StartPath "C:\Users" -OnlyPhysicalFiles $false
    Analyzes all user profiles including cloud-stored files that may not be physically present on disk.

.NOTES
    Security Level: Medium
    Required Permissions:
    - Administrative access (recommended but not required)
    - Read access to scanned directories
    - Write access to script directory for logging

    Validation Requirements:
    - Check available memory (4GB+)
    - Validate write access to log directory

    Author:  jdyer-nuvodia
    Created: 2025-02-05 00:55:03 UTC
    Updated: 2025-06-11 20:45:00 UTC

    Requirements:
    - Windows PowerShell 5.1 or later
    - Administrative privileges recommended
    - Minimum 4GB RAM recommended for large directory structures
#>

param (
    [ValidateScript({
        if([string]::IsNullOrWhiteSpace($_)) {
            throw "Path cannot be empty or whitespace."
        }
        if(!(Test-Path $_)) {
            throw "Path '$_' does not exist."
        }
        return $true
    })]
    [string]$StartPath = 'C:\',  # Starting path for folder size analysis
    [int]$MaxDepth = 10,         # Used in recursive folder scan
    [ValidateRange(1, 50)]
    [int]$Top = 3,               # Used to determine how many top folders to display
    [bool]$IncludeHiddenSystem = $true,  # Used in file filtering logic
    [bool]$FollowJunctions = $true,      # Used for handling symbolic links and junction points
    [int]$MaxThreads = 10,               # Used in Start-FolderProcessing function
    [bool]$OnlyPhysicalFiles = $true     # Used throughout the script to filter cloud files
)

# Set global information action preference for the script to ensure output visibility
$InformationPreference = 'Continue'
$ErrorActionPreference = 'SilentlyContinue'

# Reference script parameters to ensure they're recognized as used
Write-Verbose "Script parameters: MaxDepth=$MaxDepth, Top=$Top, IncludeHiddenSystem=$IncludeHiddenSystem, FollowJunctions=$FollowJunctions, MaxThreads=$MaxThreads"

# ANSI color code definitions - Using PowerShell escape syntax
$script:ANSI = @{
    # Standard colors
    Reset     = "$([char]27)[0m"
    White     = "$([char]27)[97m"  # Standard info
    Cyan      = "$([char]27)[36m"  # Process updates
    Green     = "$([char]27)[32m"  # Success
    Yellow    = "$([char]27)[33m"  # Warnings
    Red       = "$([char]27)[31m"  # Errors
    Magenta   = "$([char]27)[35m"  # Debug info
    DarkGray  = "$([char]27)[90m"  # Less important details
    # Additional colors for specific usages
    Blue      = "$([char]27)[34m"
    DarkCyan  = "$([char]27)[36m"
    DarkGreen = "$([char]27)[32m"
    DarkRed   = "$([char]27)[31m"
}

# Console colors for diagnostic output
function Write-DiagnosticMessage {
    param (
        [string]$Message,
        [string]$Color = "White"
    )

    # Use ANSI color codes with Write-Information for colored output
    $timeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $colorCode = $script:ANSI.White  # Default color

    # Map color parameter to ANSI color code
    switch ($Color) {
        "Error" { $colorCode = $script:ANSI.Red }
        "Red" { $colorCode = $script:ANSI.Red }
        "Warning" { $colorCode = $script:ANSI.Yellow }
        "Yellow" { $colorCode = $script:ANSI.Yellow }
        "Success" { $colorCode = $script:ANSI.Green }
        "Green" { $colorCode = $script:ANSI.Green }
        "Cyan" { $colorCode = $script:ANSI.Cyan }
        "Magenta" { $colorCode = $script:ANSI.Magenta }
        "DarkGray" { $colorCode = $script:ANSI.DarkGray }
        default { $colorCode = $script:ANSI.White }
    }

    # Create colored output string
    $coloredMessage = "$colorCode[$timeStamp] $Message$($script:ANSI.Reset)"

    # Use Write-Information for console output with color
    Write-Information $coloredMessage -InformationAction Continue
}

# Initial diagnostic message to show script is starting
Write-DiagnosticMessage "Script starting - Get-FolderSizes.ps1" -Color Cyan
Write-DiagnosticMessage "PowerShell Version: $($PSVersionTable.PSVersion)" -Color Cyan
Write-DiagnosticMessage "Script executed by: $env:USERNAME on $env:COMPUTERNAME" -Color Cyan

# Start transcript logging
try {
    Write-DiagnosticMessage "Starting transcript logging..." -Color Cyan

    # Check for elevated privileges but do not prompt user - continue with limited functionality
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    if (-not $isAdmin) {
        Write-DiagnosticMessage "Running with limited privileges. Some directories may be inaccessible." -Color Yellow
    }

    # Use script directory for logs instead of C:\temp
    $transcriptPath = $PSScriptRoot

    # Ensure we have a valid path - script directory should always exist when running from a script
    if (Test-Path $transcriptPath) {
        $dateFormat = "yyyy-MM-dd_HH-mm-ss"  # Changed to double quotes
        $transcriptFile = Join-Path -Path $transcriptPath -ChildPath ("FolderScan_${env:COMPUTERNAME}_$(Get-Date -Format $dateFormat).log")
        Write-DiagnosticMessage "Starting transcript at: $transcriptFile" -Color DarkGray
        Start-Transcript -Path $transcriptFile -Force -ErrorAction SilentlyContinue

        if (Test-Path $transcriptFile) {
            Write-DiagnosticMessage "Transcript file created successfully" -Color Green
        } else {
            Write-DiagnosticMessage "Failed to create transcript file" -Color "Error"
        }
    } else {
        # Fallback to user temp directory if script path is not accessible for some reason
        $dateFormat = "yyyy-MM-dd_HH-mm-ss"  # Changed to double quotes
        $transcriptFile = Join-Path -Path $env:TEMP -ChildPath ("FolderScan_${env:COMPUTERNAME}_$(Get-Date -Format $dateFormat).log")
        Write-DiagnosticMessage "Could not access script directory, using $transcriptFile instead" -Color Yellow
        Start-Transcript -Path $transcriptFile -Force -ErrorAction SilentlyContinue
    }

    Write-DiagnosticMessage "Transcript logging started successfully" -Color Green
} catch {
    Write-DiagnosticMessage "Failed to start transcript: $($_.Exception.Message)" -Color "Error"
}

#region Helper Functions

# Function to log to transcript only without console output
function Write-TranscriptOnly {
    param([string]$Message)
    $InformationPreference = 'Continue'
    Write-Information $Message 6> $null
    $InformationPreference = 'SilentlyContinue'
}

# Function to initialize color scheme for console output
function Show-ColorLegend {
    Write-Information "$($script:ANSI.Cyan)`n===== Console Output Color Legend =====$($script:ANSI.Reset)" -InformationAction Continue
    Write-Information "$($script:ANSI.White)White     - Standard information$($script:ANSI.Reset)" -InformationAction Continue
    Write-Information "$($script:ANSI.Cyan)Cyan      - Process updates and status$($script:ANSI.Reset)" -InformationAction Continue
    Write-Information "$($script:ANSI.Green)Green     - Successful operations and results$($script:ANSI.Reset)" -InformationAction Continue
    Write-Information "$($script:ANSI.Yellow)Yellow    - Warnings and attention needed$($script:ANSI.Reset)" -InformationAction Continue
    Write-Information "$($script:ANSI.Red)Red       - Errors and critical issues$($script:ANSI.Reset)" -InformationAction Continue
    Write-Information "$($script:ANSI.Magenta)Magenta   - Debug information$($script:ANSI.Reset)" -InformationAction Continue
    Write-Information "$($script:ANSI.DarkGray)DarkGray  - Technical details$($script:ANSI.Reset)" -InformationAction Continue
    Write-Information "$($script:ANSI.Cyan)======================================$($script:ANSI.Reset)`n" -InformationAction Continue
}

# New function to detect symbolic links and junction points
function Get-PathType {
    param (
        [string]$InputPath
    )

    try {
        # Special handling for OneDrive paths
        if ($InputPath -match "OneDrive -") {
            $dirInfo = New-Object System.IO.DirectoryInfo $InputPath

            if ($dirInfo.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
                # This is an OneDrive reparse point - special handling
                return @{
                    Type = "OneDriveFolder"
                    Target = "Cloud Storage"
                    IsReparsePoint = $true
                    IsOneDrive = $true
                }
            }
        }

        $dirInfo = New-Object System.IO.DirectoryInfo $InputPath
        if ($dirInfo.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
            # This is a reparse point (symbolic link, junction, etc.)
            $target = $null
            $type = "ReparsePoint"

            # Method 1: Try fsutil for most accurate results
            try {
                $fsutil = & fsutil reparsepoint query "$InputPath" 2>&1

                if ($fsutil -match "Symbolic Link") {
                    $type = "SymbolicLink"
                    # Improved parsing logic for symbolic links
                    $printNameLine = $fsutil | Where-Object { $_ -match "Print Name:" }
                    if ($printNameLine) {
                        $target = ($printNameLine -replace "^.*?Print Name:\s*", "").Trim()
                    }
                }
                elseif ($fsutil -match "Mount Point") {
                    $type = "MountPoint"
                    $printNameLine = $fsutil | Where-Object { $_ -match "Print Name:" }
                    if ($printNameLine) {
                        $target = ($printNameLine -replace "^.*?Print Name:\s*", "").Trim()
                    }
                }
                elseif ($fsutil -match "Junction") {
                    $type = "Junction"
                    $printNameLine = $fsutil | Where-Object { $_ -match "Print Name:" }
                    if ($printNameLine) {
                        $target = ($printNameLine -replace "^.*?Print Name:\s*", "").Trim()
                    }
                }
                # Check for OneDrive specific patterns in fsutil output
                elseif ($fsutil -match "OneDrive" -or $InputPath -match "OneDrive -") {
                    $type = "OneDriveFolder"
                    $target = "Cloud Storage"
                }
            }
            catch {
                Write-Verbose "fsutil method failed: $($_.Exception.Message)"
                # If path contains OneDrive, treat as OneDrive folder
                if ($InputPath -match "OneDrive -") {
                    $type = "OneDriveFolder"
                    $target = "Cloud Storage"
                }
            }

            # Method 2: Try .NET method if fsutil did not work or target is empty
            if ([string]::IsNullOrEmpty($target)) {
                try {
                    # For Windows 10/Server 2016+
                    if ($PSVersionTable.PSVersion.Major -ge 5) {
                        # Use reflection to access the Target property if available
                        $targetProperty = [System.IO.DirectoryInfo].GetProperty("Target", [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::Public)

                        if ($null -ne $targetProperty) {
                            $target = $targetProperty.GetValue($dirInfo)
                            if ($target -is [array] -and $target.Length -gt 0) {
                                $target = $target[0]  # Take first element if array
                            }
                        }
                    }
                }
                catch {
                    Write-Verbose ".NET target method failed: $($_.Exception.Message)"
                    # If path contains OneDrive, treat as OneDrive folder
                    if ($InputPath -match "OneDrive -") {
                        $type = "OneDriveFolder"
                        $target = "Cloud Storage"
                    }
                }
            }

            # Method 3: Use PowerShell native commands instead of findstr
            if ([string]::IsNullOrEmpty($target)) {
                try {
                    # Use Get-Item with -Force parameter to get link information
                    $item = Get-Item -Path $InputPath -Force -ErrorAction Stop

                    # Check for LinkType property (PowerShell 5.1+)
                    if ($item.PSObject.Properties.Name -contains "LinkType") {
                        if ($item.LinkType -eq "Junction") {
                            $type = "Junction"
                            if ($item.PSObject.Properties.Name -contains "Target") {
                                $target = $item.Target
                                if ($target -is [array] -and $target.Length -gt 0) {
                                    $target = $target[0]
                                }
                            }
                        }
                        elseif ($item.LinkType -eq "SymbolicLink") {
                            $type = "SymbolicLink"
                            if ($item.PSObject.Properties.Name -contains "Target") {
                                $target = $item.Target
                                if ($target -is [array] -and $target.Length -gt 0) {
                                    $target = $target[0]
                                }
                            }
                        }
                    }
                }
                catch {
                    Write-Verbose "PowerShell Get-Item method failed: $($_.Exception.Message)"
                }
            }

            # Final check - if we still have an Unknown Target and path has OneDrive, mark as OneDrive
            if (([string]::IsNullOrEmpty($target) -or $target -eq "Unknown Target") -and $InputPath -match "OneDrive -") {
                $type = "OneDriveFolder"
                $target = "Cloud Storage"
            }

            # Return results with either found target or "Unknown Target"
            return @{
                Type = $type
                Target = if ([string]::IsNullOrEmpty($target)) { "Unknown Target" } else { $target }
                IsReparsePoint = $true
                IsOneDrive = ($type -eq "OneDriveFolder")
            }
        }
        else {
            # Regular directory
            return @{
                Type = "Directory"
                Target = $null
                IsReparsePoint = $false
                IsOneDrive = $false
            }
        }
    }
    catch {
        Write-Warning "Error determining path type for $InputPath`: $($_.Exception.Message)"
        # Check if it might be an OneDrive path
        if ($InputPath -match "OneDrive -") {
            return @{
                Type = "OneDriveFolder"
                Target = "Cloud Storage"
                IsReparsePoint = $true
                IsOneDrive = $true
            }
        }

        return @{
            Type = "Unknown"
            Target = $null
            IsReparsePoint = $false
            IsOneDrive = $false
        }
    }
}

function Format-SizeWithPadding {
    param (
        [double]$Size,
        [int]$DecimalPlaces = 2,
        [string]$Unit = "GB"
    )

    switch ($Unit) {
        "GB" { $divider = 1GB }
        "MB" { $divider = 1MB }
        "KB" { $divider = 1KB }
        default { $divider = 1GB }
    }

    return "{0:F$DecimalPlaces}" -f ($Size / $divider)
}

function Format-Path {
    param (
        [string]$InputPath
    )
    try {
        $fullPath = [System.IO.Path]::GetFullPath($InputPath.Trim())
        return $fullPath
    }
    catch {
        Write-Warning "Error formatting path '$InputPath': $($_.Exception.Message)"
        return $InputPath
    }
}

function Write-TableHeader {
    param([int]$Width = 150)

    Write-Information ("-" * $Width) -InformationAction Continue
    Write-Information ("Folder Path".PadRight(50) + " | " +
                      "Size (GB)".PadLeft(11) + " | " +
                      "Subfolders".PadLeft(15) + " | " +
                      "Files".PadLeft(12) + " | " +
                      "Largest File (in this directory)") -InformationAction Continue
    Write-Information ("-" * $Width) -InformationAction Continue
}

function Write-TableRow {
    param(
        [string]$StartPath,
        [long]$Size,
        [int]$SubfolderCount,
        [int]$FileCount,
        [object]$LargestFile
    )

    $sizeGB = Format-SizeWithPadding -Size $Size -DecimalPlaces 2 -Unit "GB"
    $largestFileInfo = if ($LargestFile) {
        $largestFileSize = Format-SizeWithPadding -Size $LargestFile.Size -DecimalPlaces 2 -Unit "MB"
        $fileName = $LargestFile.Name

        # Check if this file has significantly different logical vs actual size
        if ($LargestFile.LogicalSize -and $LargestFile.LogicalSize -gt $LargestFile.Size * 2) {
            $logicalSizeMB = Format-SizeWithPadding -Size $LargestFile.LogicalSize -DecimalPlaces 2 -Unit "MB"
            "$fileName ($largestFileSize MB actual, $logicalSizeMB MB logical)"
        } else {
            "$fileName ($largestFileSize MB)"
        }
    } else {
        "No files found"
    }
      $outputLine = $StartPath.PadRight(50) + " | " +
                  $sizeGB.PadLeft(11) + " | " +
                  $SubfolderCount.ToString().PadLeft(15) + " | " +
                  $FileCount.ToString().PadLeft(12) + " | " +
                  $largestFileInfo

    # Select color code based on size
    $colorCode = $script:ANSI.DarkGray  # Default for small folders

    # Size-based conditional coloring - determine size threshold to set color
    try {
        $sizeGBValue = [double]($sizeGB -replace "GB", "").Trim()
        if ($sizeGBValue -gt 100) {
            $colorCode = $script:ANSI.Red  # Very large folders
        } elseif ($sizeGBValue -gt 20) {
            $colorCode = $script:ANSI.Yellow  # Large folders
        } elseif ($sizeGBValue -gt 5) {
            $colorCode = $script:ANSI.White  # Medium folders
        }
    } catch {
        # If size parsing fails, use default color
        $colorCode = $script:ANSI.White
    }
      # Use Write-Information with ANSI color codes
    Write-Information "$colorCode$outputLine$($script:ANSI.Reset)" -InformationAction Continue
}

function Write-ProgressBar {
    param (
        [int]$Completed,
        [int]$Total,
        [int]$Width = 50
    )

    $percentComplete = [math]::Min(100, [math]::Floor(($Completed / $Total) * 100))
    $filledWidth = [math]::Floor($Width * ($percentComplete / 100))
    $bar = "[" + ("=" * $filledWidth).PadRight($Width) + "] $percentComplete% | Completed processing $Completed of $Total folders"

    # For progress indication, we'll use Write-Progress instead of Write-Host
    Write-Progress -Activity "Processing Folders" -Status $bar -PercentComplete $percentComplete -Id 2
    if ($Completed -eq $Total) {
        Write-Progress -Activity "Processing Folders" -Completed -Id 2
    }
}

# Function to provide disk usage analysis and explain discrepancies
function Write-DiskUsageAnalysis {
    param(
        [string]$StartPath,
        [long]$CalculatedSize
    )
    try {
        # Extract drive letter
        $drivePath = [System.IO.Path]::GetPathRoot($StartPath)
        $driveLetter = $drivePath.TrimEnd('\')

        Write-DiagnosticMessage "Checking disk usage for drive: $driveLetter" -Color "Cyan"
          # Get actual disk information correctly
        $driveInfo = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$driveLetter'"
        if ($driveInfo) {
            $totalDiskSize = [long]$driveInfo.Size
            $freeDiskSpace = [long]$driveInfo.FreeSpace
            $usedDiskSpace = $totalDiskSize - $freeDiskSpace
              $calculatedGB = [math]::Round($CalculatedSize / 1GB, 2)
            $actualUsedGB = [math]::Round($usedDiskSpace / 1GB, 2)
            $totalDiskGB = [math]::Round($totalDiskSize / 1GB, 2)
                Write-Information "`n=== Disk Usage Analysis ===" -InformationAction Continue
            Write-Information "Drive: $($driveInfo.DeviceID)" -InformationAction Continue
            Write-Information "Total Disk Size: $totalDiskGB GB" -InformationAction Continue
            Write-Information "Actual Used Space: $actualUsedGB GB" -InformationAction Continue
            Write-Information "Script Calculated: $calculatedGB GB" -InformationAction Continue

            $difference = $calculatedGB - $actualUsedGB

            if ([math]::Abs($difference) -gt 5) {
                # Use the dedicated function for size discrepancy reporting
                Write-SizeDiscrepancyWarning -ReportedSize $calculatedGB -DiskCapacity $actualUsedGB
            } else {
                Write-Information "Calculation accuracy: Good (within 5GB)" -InformationAction Continue
            }
            Write-Information "=============================`n" -InformationAction Continue
        }    } catch {
        Write-DiagnosticMessage "Could not retrieve disk information for analysis: $($_.Exception.Message)" -Color "Warning"
    }
}

# Function to provide detailed explanation of size discrepancy between calculated size and actual disk usage
function Write-SizeDiscrepancyWarning {
    param(
        [double]$ReportedSize,
        [double]$DiskCapacity
    )

    $difference = $ReportedSize - $DiskCapacity
    $roundedDiff = [math]::Round($difference, 2)

    Write-Information "Size Difference: $roundedDiff GB" -InformationAction Continue

    if ($difference -gt 0) {
        Write-Information "`nPossible reasons for over-reporting:" -InformationAction Continue
        Write-Information "- Sparse files (hibernation/page files) showing logical vs actual size" -InformationAction Continue
        Write-Information "- NTFS compression reducing actual disk usage" -InformationAction Continue
        Write-Information "- Hard links causing duplicate counting" -InformationAction Continue
        Write-Information "- Junction points creating circular references" -InformationAction Continue
    } else {
        Write-Information "`nPossible reasons for under-reporting:" -InformationAction Continue
        Write-Information "- Access denied to some system directories" -InformationAction Continue
        Write-Information "- File system metadata and journal overhead" -InformationAction Continue
        Write-Information "- Reserved disk space not included in calculations" -InformationAction Continue
        Write-Information "- Shadow copies and system restore points" -InformationAction Continue
    }
}

# Function to get drive statistics for a specific drive letter
function Get-DriveStat {
    param (
        [Parameter(Mandatory = $true)]
        [string]$DriveLetter
    )

    try {
        # Ensure the drive letter is properly formatted with colon
        $DriveLetter = $DriveLetter.TrimEnd(':') + ':'

        # Use CIM instance to get drive information
        $volume = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$DriveLetter'" -ErrorAction Stop

        if ($volume) {
            return $volume
        } else {
            Write-DiagnosticMessage "No volume found for drive letter: $DriveLetter" -Color "Warning"
            return $null
        }
    } catch {
        Write-DiagnosticMessage "Error getting drive stats for $DriveLetter`: $($_.Exception.Message)" -Color "Error"
        return $null
    }
}

# Function to display drive information in a formatted way
function Show-DriveInfo {
    param (
        [Parameter(Mandatory = $false)]
        [object]$Volume
    )

    if ($null -eq $Volume) {
        Write-Information "`n$($script:ANSI.Yellow)Drive information unavailable.$($script:ANSI.Reset)" -InformationAction Continue
        return
    }

    try {
        $freeSpaceGB = [math]::Round($Volume.FreeSpace / 1GB, 2)
        $totalSizeGB = [math]::Round($Volume.Size / 1GB, 2)
        $usedSpaceGB = [math]::Round(($Volume.Size - $Volume.FreeSpace) / 1GB, 2)
        $percentFree = [math]::Round(($Volume.FreeSpace / $Volume.Size) * 100, 2)

        Write-Information "`n$($script:ANSI.Cyan)Drive Information: $($Volume.DeviceID)$($script:ANSI.Reset)" -InformationAction Continue
        Write-Information "$($script:ANSI.White)Total Size: $totalSizeGB GB$($script:ANSI.Reset)" -InformationAction Continue
        Write-Information "$($script:ANSI.White)Used Space: $usedSpaceGB GB$($script:ANSI.Reset)" -InformationAction Continue
        Write-Information "$($script:ANSI.White)Free Space: $freeSpaceGB GB ($percentFree%)$($script:ANSI.Reset)" -InformationAction Continue

        # Color-code the free space percentage
        if ($percentFree -lt 10) {
            Write-Information "$($script:ANSI.Red)WARNING: Low disk space!$($script:ANSI.Reset)" -InformationAction Continue
        } elseif ($percentFree -lt 25) {
            Write-Information "$($script:ANSI.Yellow)Note: Disk space below 25%$($script:ANSI.Reset)" -InformationAction Continue
        }
    } catch {
        Write-DiagnosticMessage "Error displaying drive information: $($_.Exception.Message)" -Color "Error"
    }
}

#endregion

#region Setup

# Check for elevated privileges but do not prompt user - continue with limited functionality
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    Write-Warning "Running with limited privileges. Some directories may be inaccessible."
}

# Script Header in Transcript
Write-Information -MessageData "======================================================" -InformationAction Continue
Write-Information -MessageData "Folder Size Scanner - Execution Log" -InformationAction Continue
Write-Information -MessageData "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -InformationAction Continue
Write-Information -MessageData "Started (UTC): $((Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss'))" -InformationAction Continue
Write-Information -MessageData "User: $env:USERNAME" -InformationAction Continue
Write-Information -MessageData "Computer: $env:COMPUTERNAME" -InformationAction Continue
Write-Information -MessageData "Target Path: $StartPath" -InformationAction Continue
Write-Information -MessageData "Admin Privileges: $isAdmin" -InformationAction Continue
Write-Information -MessageData "OneDrive Mode: $(if($OnlyPhysicalFiles){"Only Physical Files"}else{"Include All Files"})" -InformationAction Continue
Write-Information -MessageData "Include Hidden/System: $IncludeHiddenSystem" -InformationAction Continue
Write-Information -MessageData "Follow Junctions: $FollowJunctions" -InformationAction Continue
Write-Information -MessageData "Max Thread Count: $MaxThreads" -InformationAction Continue
Write-Information -MessageData "======================================================" -InformationAction Continue
Write-Information -MessageData "" -InformationAction Continue

# Show color legend for user reference
Show-ColorLegend

# .NET Type Definition
# Use PowerShell functions instead of compiled C# to avoid compilation errors

# Define constants for file attributes
$script:FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS = 0x00400000
$script:FILE_ATTRIBUTE_RECALL_ON_OPEN = 0x00040000
$script:FILE_ATTRIBUTE_REPARSE_POINT = 0x00000400
$script:FILE_ATTRIBUTE_SYSTEM = 0x00000004
$script:FILE_ATTRIBUTE_HIDDEN = 0x00000002

# Special system files that can cause size inconsistencies
$script:SpecialSystemFiles = @(
    "hiberfil.sys",
    "pagefile.sys",
    "swapfile.sys"
)

# Define a PowerShell class to replace the C# FileDetails class
class FileDetails {
    [string]$Name
    [string]$Path
    [long]$Size
    [long]$LogicalSize

    FileDetails([string]$name, [string]$path, [long]$size, [long]$logicalSize) {
        $this.Name = $name
        $this.Path = $path
        $this.Size = $size
        $this.LogicalSize = $logicalSize
    }
}

# Create a simple static class emulator using the PowerShell script scope
# to hold our folder size helper methods

# Check if a file is physically stored or is a cloud placeholder
function script:IsFilePhysicallyStored {
    param([string]$FilePath)

    try {
        # First check if this is in an OneDrive folder
        if ($FilePath -match "OneDrive -") {
            # Use Get-PathType to determine if it's a placeholder or local file
            $dirPath = [System.IO.Path]::GetDirectoryName($FilePath)
            $pathType = Get-PathType -InputPath $dirPath

            # If the path is detected as OneDrive, we need to check file attributes
            if ($pathType.IsOneDrive) {
                $attrs = [System.IO.File]::GetAttributes($FilePath)

                # Check if it has cloud attributes (placeholder)
                $isPlaceholder = (([int]$attrs -band $script:FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS) -ne 0) -or
                                 (([int]$attrs -band $script:FILE_ATTRIBUTE_RECALL_ON_OPEN) -ne 0)

                # If it's not a placeholder, it's stored locally
                return -not $isPlaceholder
            }
        }

        # For non-OneDrive files or if OneDrive check fails, check attributes directly
        $attrs = [System.IO.File]::GetAttributes($FilePath)

        # Check if it has cloud attributes (placeholder)
        $isPlaceholder = (([int]$attrs -band $script:FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS) -ne 0) -or
                         (([int]$attrs -band $script:FILE_ATTRIBUTE_RECALL_ON_OPEN) -ne 0)

        # If it's not a placeholder, it's stored locally
        return -not $isPlaceholder
    }
    catch {
        # Default to true if we can't check (better to overcount than undercount)
        return $true
    }
}

# Check if a file is a special system file that might cause size inconsistencies
function script:IsSpecialSystemFile {
    param([string]$FilePath)

    try {
        $fileName = [System.IO.Path]::GetFileName($FilePath)

        # Check against our list of special files
        return $script:SpecialSystemFiles -contains $fileName
    }
    catch {
        return $false
    }
}

# Get the actual size of a file, which might be different than the logical size
function script:GetActualFileSize {
    param([string]$FilePath)

    try {
        $file = New-Object System.IO.FileInfo($FilePath)
        return $file.Length  # For now, just return logical size
    }
    catch {
        return 0
    }
}

# Calculate the total size of a directory
function script:GetDirectorySize {
    param(
        [string]$Path,
        [bool]$OnlyPhysicalFiles,
        [bool]$FollowJunctions = $true
    )

    $size = 0
    $stack = New-Object System.Collections.Generic.Stack[string]
    $stack.Push($Path)

    while ($stack.Count -gt 0) {
        $dir = $stack.Pop()
        try {
            foreach ($file in [System.IO.Directory]::GetFiles($dir)) {
                try {
                    # Skip non-physical files if requested
                    if ($OnlyPhysicalFiles -and -not (script:IsFilePhysicallyStored $file)) {
                        continue
                    }

                    # Add file size
                    $size += (New-Object System.IO.FileInfo($file)).Length
                }
                catch {
                    # Log error but continue processing
                    Write-Verbose "Error processing file $file`: $($_.Exception.Message)"
                }
            }

            foreach ($subDir in [System.IO.Directory]::GetDirectories($dir)) {
                # Check if it's a junction point or symbolic link
                $dirInfo = New-Object System.IO.DirectoryInfo($subDir)
                $isReparsePoint = $dirInfo.Attributes.ToString() -match 'ReparsePoint'

                # Skip junction points if not following them
                if ($isReparsePoint -and -not $FollowJunctions) {
                    continue
                }

                $stack.Push($subDir)
            }
        }
        catch {
            Write-Verbose "Error accessing directory $dir`: $($_.Exception.Message)"
        }
    }
    return $size
}

# Count the number of files and subdirectories in a directory
function script:GetDirectoryCounts {
    param(
        [string]$Path,
        [bool]$OnlyPhysicalFiles,
        [bool]$FollowJunctions = $true
    )

    $files = 0
    $folders = 0
    $stack = New-Object System.Collections.Generic.Stack[string]
    $stack.Push($Path)

    while ($stack.Count -gt 0) {
        $dir = $stack.Pop()
        try {
            if ($OnlyPhysicalFiles) {
                $filesList = [System.IO.Directory]::GetFiles($dir)
                foreach ($file in $filesList) {
                    if (script:IsFilePhysicallyStored $file) {
                        $files++
                    }
                }
            }
            else {
                $files += [System.IO.Directory]::GetFiles($dir).Length
            }

            $subDirs = [System.IO.Directory]::GetDirectories($dir)
            foreach ($subDir in $subDirs) {
                # Check if it's a junction point or symbolic link
                $dirInfo = New-Object System.IO.DirectoryInfo($subDir)
                $isReparsePoint = $dirInfo.Attributes.ToString() -match 'ReparsePoint'

                # Skip junction points if not following them
                if ($isReparsePoint -and -not $FollowJunctions) {
                    continue
                }

                $folders++
                $stack.Push($subDir)
            }
        }
        catch {
            Write-Verbose "Error counting files/folders in $dir`: $($_.Exception.Message)"
        }
    }
    return [System.Tuple]::Create($files, $folders)
}

# Find the largest file in a directory
function script:GetLargestFile {
    param(
        [string]$Path,
        [bool]$OnlyPhysicalFiles,
        [bool]$FollowJunctions = $true
    )

    try {
        # This will be our largest file across all directories
        $largestFileInfo = $null
        $largestFileSize = 0

        # Use a stack for depth-first traversal
        $stack = New-Object System.Collections.Generic.Stack[string]
        $stack.Push($Path)

        while ($stack.Count -gt 0) {
            $currentDir = $stack.Pop()

            try {
                # Get files in the current directory
                $allFiles = (New-Object System.IO.DirectoryInfo($currentDir)).GetFiles("*.*", [System.IO.SearchOption]::TopDirectoryOnly)

                # Filter for physical files if requested
                $filteredFiles = $allFiles
                if ($OnlyPhysicalFiles) {
                    $filteredFiles = $allFiles | Where-Object { script:IsFilePhysicallyStored $_.FullName }
                }

                # Find largest file in current directory
                $currentDirLargestFile = $filteredFiles | Sort-Object Length -Descending | Select-Object -First 1

                if ($currentDirLargestFile -and $currentDirLargestFile.Length -gt $largestFileSize) {
                    $largestFileInfo = $currentDirLargestFile
                    $largestFileSize = $currentDirLargestFile.Length
                }

                # Process subdirectories
                $subDirs = (New-Object System.IO.DirectoryInfo($currentDir)).GetDirectories()
                foreach ($subDir in $subDirs) {
                    # Check if it's a junction point or symbolic link
                    $isReparsePoint = $subDir.Attributes.ToString() -match 'ReparsePoint'

                    # Skip junction points if not following them
                    if ($isReparsePoint -and -not $FollowJunctions) {
                        continue
                    }

                    $stack.Push($subDir.FullName)
                }
            }
            catch {
                Write-Verbose "Error accessing directory $currentDir`: $($_.Exception.Message)"
                continue
            }
        }

        if ($largestFileInfo) {
            return [FileDetails]::new(
                $largestFileInfo.Name,
                $largestFileInfo.FullName,
                (script:GetActualFileSize $largestFileInfo.FullName),
                $largestFileInfo.Length
            )
        }

        return $null
    }
    catch {
        return $null
    }
}

Write-Information "Ultra-fast folder analysis starting at: $StartPath" -InformationAction Continue
Write-Information "Script started by: $env:USERNAME at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -InformationAction Continue

#endregion

#region Folder Scanning Logic

# New function to process folders in parallel using runspaces
function Start-FolderProcessing {
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([System.Collections.Hashtable])]
    param(
        [array]$Folders,
        [int]$MaxThreads,
        [bool]$OnlyPhysicalFiles,
        [bool]$FollowJunctions
    )

    if (-not $PSCmdlet.ShouldProcess("$($Folders.Count) folders with $MaxThreads threads", "Start parallel folder processing")) {
        return $null
    }

    $RunspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads)
    $RunspacePool.Open()
    $FolderSizeMap = @{}
    $Runspaces = @()
    $activeRunspaces = 0
    $processedCount = 0
    $totalFolders = $Folders.Count
      Write-Information "`nParallel Processing Configuration:" -InformationAction Continue
    Write-Information "Maximum Threads: $MaxThreads" -InformationAction Continue
    Write-Information "Total Folders to Process: $totalFolders" -InformationAction Continue
    Write-Information "Only Physical Files: $OnlyPhysicalFiles" -InformationAction Continue
      foreach ($folder in $Folders) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $RunspacePool
        $activeRunspaces++

        # Calculate progress percentage with a maximum of 100
        $progressPercent = [Math]::Min(100, [Math]::Round(($activeRunspaces / $MaxThreads) * 100))
        Write-Progress -Activity "Processing Folders" -Status "Active Runspaces: $activeRunspaces/$MaxThreads" -PercentComplete $progressPercent -Id 1
        [void]$ps.AddScript({
            param($StartPath, $OnlyPhysicalFiles, $FollowJunctions)

            $threadId = [System.Threading.Thread]::CurrentThread.ManagedThreadId
            Write-Verbose "Thread $threadId processing: $StartPath"

            try {
                $counts = script:GetDirectoryCounts -Path $StartPath -OnlyPhysicalFiles $OnlyPhysicalFiles -FollowJunctions $FollowJunctions
                $size = script:GetDirectorySize -Path $StartPath -OnlyPhysicalFiles $OnlyPhysicalFiles -FollowJunctions $FollowJunctions
                $largestFile = script:GetLargestFile -Path $StartPath -OnlyPhysicalFiles $OnlyPhysicalFiles -FollowJunctions $FollowJunctions

                return @{
                    Success = $true
                    StartPath = $StartPath
                    Size = $size
                    FileCount = $counts.Item1
                    FolderCount = $counts.Item2
                    LargestFile = $largestFile
                    ThreadId = $threadId
                }
            }
            catch {
                return @{
                    Success = $false
                    StartPath = $StartPath
                    Error = $_.Exception.Message
                    ThreadId = $threadId
                }
            }
        }).AddArgument($folder.FullName).AddArgument($OnlyPhysicalFiles).AddArgument($FollowJunctions)

        $Runspaces += [PSCustomObject]@{
            Instance = $ps
            Handle = $ps.BeginInvoke()
            Folder = $folder.FullName
            StartTime = [DateTime]::Now
        }
    }

    Write-TranscriptOnly "`n`nProcessing Results:"

    foreach ($r in $Runspaces) {
        try {
            $processedCount++
            $percentComplete = [math]::Round(($processedCount / $totalFolders) * 100, 1)

            $result = $r.Instance.EndInvoke($r.Handle)
            $processingTime = ([DateTime]::Now - $r.StartTime).TotalSeconds

            # Log detailed progress to transcript only
            Write-TranscriptOnly "`nProgress: $processedCount/$totalFolders ($percentComplete%)"
            Write-TranscriptOnly "Processing: $($r.Folder)"

            if ($result.Success) {
                Write-TranscriptOnly "Thread $($result.ThreadId) completed: $($result.StartPath) in $($processingTime.ToString('0.00'))s"

                $FolderSizeMap[$result.StartPath] = @{
                    Size = $result.Size
                    FileCount = $result.FileCount
                    FolderCount = $result.FolderCount
                    LargestFile = $result.LargestFile
                }
            }            else {
                Write-Error "Thread $($result.ThreadId) failed: $($r.Folder) - $($result.Error)"
            }
            $activeRunspaces--
        }
        catch {
            Write-Error "Critical error in runspace for folder $($r.Folder): $($_.Exception.Message)"
            $activeRunspaces--
        }
        finally {
            $r.Instance.Dispose()
        }
    }
    Write-Information -MessageData "`n`nParallel Processing Summary:" -InformationAction Continue
    Write-Information -MessageData "Total Folders Processed: $processedCount" -InformationAction Continue
    Write-Information -MessageData "Maximum Concurrent Threads: $MaxThreads" -InformationAction Continue

    $RunspacePool.Close()
    $RunspacePool.Dispose()
    return $FolderSizeMap
}

# Modify the Get-FolderSize function to use parallel processing
function Get-FolderSize {
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param (
        [string]$StartPath,
        [int]$CurrentDepth,
        [int]$MaxDepth,
        [int]$Top,
        [bool]$OnlyPhysicalFiles
    )

    try {
        # Validate input path
        if([string]::IsNullOrWhiteSpace($StartPath)) {
            Write-Warning "Invalid path: Path cannot be empty or whitespace"
            return @{
                ProcessedFolders = $false
                HasSubfolders = $false
                CompletionMessageShown = $false
            }
        }

        # Normalize path to ensure consistent formatting
        try {
            $StartPath = [System.IO.Path]::GetFullPath($StartPath)
        } catch {
            Write-Warning "Error normalizing path '$StartPath': $($_.Exception.Message)"
            return @{
                ProcessedFolders = $false
                HasSubfolders = $false
                CompletionMessageShown = $false
            }
        }

        if ($CurrentDepth -gt $MaxDepth) {
            return @{
                ProcessedFolders = $false
                HasSubfolders = $false
                CompletionMessageShown = $false
            }
        }

        $StartPath = Format-Path $StartPath
        if (-not (Test-Path -Path $StartPath -PathType Container)) {
            Write-Warning "Path '$StartPath' does not exist or is not a directory."
            return @{
                ProcessedFolders = $false
                HasSubfolders = $false
                CompletionMessageShown = $false
            }        }

        Write-Information "`nTop $Top Largest Folders in: $StartPath" -InformationAction Continue
        Write-Information "" -InformationAction Continue

        # First, analyze the root path itself
        if ($CurrentDepth -eq 1) {            $rootSize = script:GetDirectorySize -Path $StartPath -OnlyPhysicalFiles $OnlyPhysicalFiles -FollowJunctions $FollowJunctions
            $rootCounts = script:GetDirectoryCounts -Path $StartPath -OnlyPhysicalFiles $OnlyPhysicalFiles -FollowJunctions $FollowJunctions
            $rootLargestFile = script:GetLargestFile -Path $StartPath -OnlyPhysicalFiles $OnlyPhysicalFiles -FollowJunctions $FollowJunctions

            Write-TableHeader
            Write-TableRow -StartPath $StartPath `
                          -Size $rootSize `
                          -SubfolderCount $rootCounts.Item2 `                          -FileCount $rootCounts.Item1 `
                          -LargestFile $rootLargestFile
            Write-Information ("-" * 150) -InformationAction Continue
            Write-Information "" -InformationAction Continue
        }

        # Get all immediate subfolders in the root and process them
        $rootFolders = try {
            if ($IncludeHiddenSystem) {
                Get-ChildItem -Path $StartPath -Directory -Force -ErrorAction Stop
            }
            else {
                Get-ChildItem -Path $StartPath -Directory -ErrorAction Stop
            }
        } catch {
            Write-Warning "Error getting root folders in '$StartPath': $($_.Exception.Message)"
            @()        }

        # Process root level folders first
        if ($rootFolders -and $rootFolders.Count -gt 0) {
            Write-Information "Processing $($rootFolders.Count) folders in root directory..." -InformationAction Continue

            # Process root folders in parallel using the MaxThreads parameter
            $folderResults = Start-FolderProcessing -Folders $rootFolders -MaxThreads $MaxThreads -OnlyPhysicalFiles $OnlyPhysicalFiles -FollowJunctions $FollowJunctions

            # Convert results to sorted array
            $sortedFolders = $folderResults.GetEnumerator() | ForEach-Object {
                [PSCustomObject]@{
                    Path = $_.Key
                    Size = $_.Value.Size
                    FileCount = $_.Value.FileCount
                    FolderCount = $_.Value.FolderCount
                    LargestFile = $_.Value.LargestFile
                }
            } | Sort-Object -Property Size -Descending
              # Display table of root folders
            Write-TableHeader

            # Get top folders but ensure we do not exceed available folders
            $topFoldersCount = [Math]::Min($Top, $sortedFolders.Count)
            $topFolders = $sortedFolders | Select-Object -First $topFoldersCount

            # Debug information
            Write-DiagnosticMessage "Found $($sortedFolders.Count) sorted folders, displaying top $topFoldersCount" -Color "Cyan"

            # Display table rows for each top folder
            if ($topFolders -and $topFolders.Count -gt 0) {
                foreach ($folder in $topFolders) {
                    Write-TableRow -StartPath $folder.Path -Size $folder.Size -SubfolderCount $folder.FolderCount -FileCount $folder.FileCount -LargestFile $folder.LargestFile
                }
                Write-Information -MessageData ("-" * 150) -InformationAction Continue
                Write-Information -MessageData "" -InformationAction Continue
            } else {
                Write-DiagnosticMessage "No top folders to display" -Color "Warning"                Write-Information -MessageData "No folders found for display." -InformationAction Continue
                Write-Information -MessageData ("-" * 150) -InformationAction Continue
                Write-Information -MessageData "" -InformationAction Continue
            }

            # Process all top folders up to max depth
            $completionMessageShown = $false
            if ($CurrentDepth + 1 -le $MaxDepth -and $sortedFolders.Count -gt 0) {
                # Process top 5 (or as specified by $Top) folders at each level
                $foldersToProcess = $sortedFolders | Select-Object -First $Top
                  foreach ($folder in $foldersToProcess) {
                    # Display header for this subfolder level with a clear descending message                    Write-Information -MessageData "`n$($script:ANSI.Green)Descending into largest subfolder: $($folder.Path)$($script:ANSI.Reset)" -InformationAction Continue
                    Write-Information -MessageData "`n$($script:ANSI.Cyan)Analyzing subfolder: $($folder.Path)$($script:ANSI.Reset)" -InformationAction Continue
                    Write-Information -MessageData "$($script:ANSI.DarkGray)$("-" * 100)$($script:ANSI.Reset)" -InformationAction Continue

                    # First display information about this specific folder
                    $subPathInfo = Get-Item -Path $folder.Path -Force -ErrorAction SilentlyContinue
                    if ($subPathInfo) {
                        $subFolderCount = (Get-ChildItem -Path $folder.Path -Directory -Force -ErrorAction SilentlyContinue).Count
                        $subFileCount = (Get-ChildItem -Path $folder.Path -File -Force -ErrorAction SilentlyContinue).Count
                        Write-Information "Files: $subFileCount, Folders: $subFolderCount, Total Size: $([math]::Round($folder.Size / 1GB, 2)) GB" -InformationAction Continue
                    }
                      # Get subfolders for this folder and process them
                    $subFolders = try {
                        if ($IncludeHiddenSystem) {
                            Get-ChildItem -Path $folder.Path -Directory -Force -ErrorAction Stop
                        }
                        else {
                            Get-ChildItem -Path $folder.Path -Directory -ErrorAction Stop
                        }
                    } catch {
                        Write-Warning "Error getting subfolders in '$($folder.Path)': $($_.Exception.Message)"
                        @()
                    }

                    # Process subfolders if they exist and display a table for them
                    if ($subFolders -and $subFolders.Count -gt 0) {
                        Write-Information "`n$($script:ANSI.DarkCyan)Processing $($subFolders.Count) subfolders in: $($folder.Path)$($script:ANSI.Reset)" -InformationAction Continue

                        # Process subfolders in parallel
                        $subFolderResults = Start-FolderProcessing -Folders $subFolders -MaxThreads $MaxThreads -OnlyPhysicalFiles $OnlyPhysicalFiles -FollowJunctions $FollowJunctions

                        # Convert results to sorted array
                        $sortedSubFolders = $subFolderResults.GetEnumerator() | ForEach-Object {
                            [PSCustomObject]@{
                                Path = $_.Key
                                Size = $_.Value.Size
                                FileCount = $_.Value.FileCount
                                FolderCount = $_.Value.FolderCount
                                LargestFile = $_.Value.LargestFile
                            }
                        } | Sort-Object -Property Size -Descending

                        # Display table of subfolders if we have any
                        if ($sortedSubFolders.Count -gt 0) {
                            Write-TableHeader

                            # Get top subfolders but ensure we do not exceed available folders
                            $topSubFoldersCount = [Math]::Min($Top, $sortedSubFolders.Count)
                            $topSubFolders = $sortedSubFolders | Select-Object -First $topSubFoldersCount

                            foreach ($subFolder in $topSubFolders) {
                                Write-TableRow -StartPath $subFolder.Path -Size $subFolder.Size -SubfolderCount $subFolder.FolderCount -FileCount $subFolder.FileCount -LargestFile $subFolder.LargestFile
                            }

                            Write-Information ("-" * 150) -InformationAction Continue
                            Write-Information "" -InformationAction Continue
                        }
                    }

                    # Call recursively for deeper levels and capture the structured return value
                    if ($CurrentDepth + 1 -lt $MaxDepth) {
                        $result = Get-FolderSize -StartPath $folder.Path -CurrentDepth ($CurrentDepth + 1) -MaxDepth $MaxDepth -Top $Top -OnlyPhysicalFiles $OnlyPhysicalFiles

                        if ($result.ProcessedFolders -eq $true -and
                            $result.HasSubfolders -eq $true -and
                            $result.CompletionMessageShown -eq $false) {
                            Write-Information "`n$($script:ANSI.Green)Completed processing subfolder: $($folder.Path)$($script:ANSI.Reset)" -InformationAction Continue
                            $completionMessageShown = $true
                        } else {
                            $completionMessageShown = $result.CompletionMessageShown
                        }
                    }
                }
            }

            return @{
                ProcessedFolders = $true
                HasSubfolders = $true
                CompletionMessageShown = $completionMessageShown
            }
        } else {
            Write-Warning "No subfolders found to process."
            return @{
                ProcessedFolders = $true
                HasSubfolders = $false
                CompletionMessageShown = $false
            }
        }
    }    catch {
        Write-Warning "Error processing folder '$StartPath': $($_.Exception.Message)"
        return @{
            ProcessedFolders = $false
            HasSubfolders = $false
            CompletionMessageShown = $false        }
    }
} # End of Get-FolderSize function

# Start the Recursive Scan
# Get the total calculated size for analysis
$rootSize = script:GetDirectorySize -Path $StartPath -OnlyPhysicalFiles $OnlyPhysicalFiles -FollowJunctions $FollowJunctions

# Perform disk usage analysis
Write-DiskUsageAnalysis -StartPath $StartPath -CalculatedSize $rootSize

# Start the folder size analysis with recursive processing
Write-Information -MessageData "$($script:ANSI.Cyan)`nBeginning recursive folder scan with max depth of $MaxDepth and top $Top folders at each level$($script:ANSI.Reset)" -InformationAction Continue

# Clear any existing progress bars before starting main processing
Write-Progress -Activity "Processing Folders" -Completed -Id 1
Write-Progress -Activity "Processing Folders" -Completed -Id 2

$result = Get-FolderSize -StartPath $StartPath -CurrentDepth 1 -MaxDepth $MaxDepth -Top $Top -OnlyPhysicalFiles $OnlyPhysicalFiles

# Get drive information for completion
$driveLetter = $StartPath.Substring(0, 1)
Write-DiagnosticMessage "Getting drive stats for: $driveLetter" -Color "Cyan"
$driveInfo = Get-DriveStat -DriveLetter $driveLetter

# Show drive information
Show-DriveInfo -Volume $driveInfo

# Script completed successfully
Write-Information -MessageData "$($script:ANSI.Green)`nScript finished at $((Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')) (UTC)$($script:ANSI.Reset)" -InformationAction Continue

# Stop the transcript at the end and properly release file handles
try {
    # Check if transcript is active before trying to stop it
    $transcriptStatus = Get-PSCallStack | Where-Object { $_.Command -eq 'Start-Transcript' }

    if ($transcriptStatus) {
        # Stop the transcript only if it's running
        Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
        Write-Information -MessageData "$($script:ANSI.Cyan)Transcript saved, output file is $script:TranscriptFile$($script:ANSI.Reset)" -InformationAction Continue

        # Force garbage collection to release any remaining file handles
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
} catch {
    # Ignore errors stopping transcript
    Write-DiagnosticMessage "Warning: Error stopping transcript: $($_.Exception.Message)" -Color "Warning"
}

#endregion
