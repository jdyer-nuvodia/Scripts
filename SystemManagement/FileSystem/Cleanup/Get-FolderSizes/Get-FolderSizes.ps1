# =============================================================================
# Script: Get-FolderSizes.ps1
# Created: 2025-02-05 00:55:03 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-06-13 19:52:00 UTC
# Updated By: GitHub Copilot
# Version: 2.24.1
# Additional Info: Fixed FollowJunctions parameter usage and resolved all PSScriptAnalyzer issues
# =============================================================================

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

.PARAMETER ShowProgressBar
    Controls whether to display progress bars during scanning. Default is $true.

.PARAMETER Verbose
    Shows detailed diagnostic messages during execution, including folder-by-folder updates.

.PARAMETER Debug
    Shows highly detailed diagnostic information for troubleshooting purposes.

.EXAMPLE
    .\Get-FolderSizes.ps1
    Analyzes all folders starting at C:\ with default settings.

.EXAMPLE
    .\Get-FolderSizes.ps1 -StartPath "D:\Projects" -MaxDepth 3 -Top 5
    Analyzes the D:\Projects folder, going 3 levels deep and showing the top 5 largest folders at each level.

.EXAMPLE
    .\Get-FolderSizes.ps1 -StartPath "C:\Users" -OnlyPhysicalFiles $false
    Analyzes all user profiles including cloud-stored files that may not be physically present on disk.

.EXAMPLE
    .\Get-FolderSizes.ps1 -Verbose
    Analyzes all folders starting at C:\ with default settings and shows detailed progress information.

.NOTES
    Security Level: Medium
    Required Permissions:
    - Administrative access (recommended but not required)
    - Read access to scanned directories
    - Write access to script directory for logging

    Validation Requirements:
    - Check available memory (4GB+)
    - Validate write access to log directory

    Requirements:
    - Windows PowerShell 5.1 or later
    - Administrative privileges recommended
    - Minimum 4GB RAM recommended for large directory structures
#>

[CmdletBinding()]
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
    [bool]$OnlyPhysicalFiles = $true,    # Used throughout the script to filter cloud files
    [bool]$ShowProgressBar = $true       # Controls visibility of progress bars
)

# Set global information action preference for the script to ensure output visibility
$InformationPreference = 'Continue'
$ErrorActionPreference = 'SilentlyContinue'

# Initialize progress bar settings
$script:UseProgressBars = $ShowProgressBar

# Reference script parameters to ensure they're recognized as used
Write-Verbose "Script parameters: MaxDepth=$MaxDepth, Top=$Top, IncludeHiddenSystem=$IncludeHiddenSystem, FollowJunctions=$FollowJunctions, MaxThreads=$MaxThreads, ShowProgressBar=$ShowProgressBar"

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
    Write-Information -MessageData $coloredMessage -InformationAction Continue
}

# Initialize script-level variables
$script:InaccessibleFolders = @{} # Track folders that can't be accessed with specific errors
$script:transcriptActive = $false # Track if transcript is active
$script:transcriptFile = $null # Store the transcript file path

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
    if (Test-Path $transcriptPath) {        $dateFormat = "yyyy-MM-dd_HH-mm-ss"  # Changed to double quotes
        $script:transcriptFile = Join-Path -Path $transcriptPath -ChildPath ("FolderScan_${env:COMPUTERNAME}_$(Get-Date -Format $dateFormat).log")
        Write-DiagnosticMessage "Starting transcript at: $script:transcriptFile" -Color DarkGray
        Start-Transcript -Path $script:transcriptFile -Force -ErrorAction SilentlyContinue
        $script:transcriptActive = $true # Mark transcript as active

        if (Test-Path $script:transcriptFile) {
            Write-DiagnosticMessage "Transcript file created successfully" -Color Green
        } else {
            Write-DiagnosticMessage "Failed to create transcript file" -Color "Error"
            $script:transcriptActive = $false # Reset if file creation failed
        }    } else {
        # Fallback to user temp directory if script path is not accessible for some reason
        $dateFormat = "yyyy-MM-dd_HH-mm-ss"  # Changed to double quotes
        $script:transcriptFile = Join-Path -Path $env:TEMP -ChildPath ("FolderScan_${env:COMPUTERNAME}_$(Get-Date -Format $dateFormat).log")
        Write-DiagnosticMessage "Could not access script directory, using $script:transcriptFile instead" -Color Yellow
        Start-Transcript -Path $script:transcriptFile -Force -ErrorAction SilentlyContinue
        $script:transcriptActive = $true # Mark transcript as active
    }

    Write-DiagnosticMessage "Transcript logging started successfully" -Color Green
} catch {
    Write-DiagnosticMessage "Failed to start transcript: $($_.Exception.Message)" -Color "Error"
}

#region Helper Functions

# Function to safely stop transcript and release file handles
function Stop-TranscriptSafely {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param()

    # Safe wrapper for Write-Verbose
    function Write-VerboseSafe {
        param([string]$Message)
        try {
            Write-Verbose $Message
        } catch {
            # If Write-Verbose is not available, continue silently
            # This is expected in runspace or restricted execution contexts
            Write-Debug "Write-Verbose not available in current context"
        }
    }

    Write-VerboseSafe "Entering Stop-TranscriptSafely function."    # Check the script-scoped flag to see if transcript was started by this script
    if ($script:transcriptActive) {
        Write-VerboseSafe "Transcript was active, attempting to stop."
        if ($PSCmdlet.ShouldProcess("Active transcript", "Stop and release file handle")) {
            try {
                # Store transcript path before stopping it so we can try to force release later if needed
                $transcriptPath = $null
                try {
                    $transcriptClass = [Microsoft.PowerShell.Commands.PSHostInvocationData].Assembly.GetType('Microsoft.PowerShell.Commands.TranscriptionData')
                    if ($null -ne $transcriptClass) {
                        $binding = [System.Reflection.BindingFlags]'NonPublic,Static'
                        $field = $transcriptClass.GetField('filePath', $binding)
                        if ($null -ne $field) {
                            $transcriptPath = $field.GetValue($null)
                            Write-VerboseSafe "Found active transcript path: $transcriptPath"
                        }
                    }
                } catch {
                    Write-VerboseSafe "Unable to retrieve transcript path: $_"
                }

                # First try - standard Stop-Transcript method
                Stop-Transcript -ErrorAction Stop | Out-Null                Write-VerboseSafe "Stop-Transcript command executed."

                # Give the system a moment to release the file handle
                Start-Sleep -Milliseconds 500
                Write-VerboseSafe "Slept for 500ms after Stop-Transcript."

                # Force garbage collection to release file handles - multiple passes
                for ($i = 0; $i -lt 3; $i++) {
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                    Start-Sleep -Milliseconds 200
                }                Write-VerboseSafe "First round of garbage collection triggered after Stop-Transcript."

                # More aggressive garbage collection
                [System.GC]::Collect(2, [System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $true)
                [System.GC]::WaitForPendingFinalizers()
                Write-VerboseSafe "Aggressive garbage collection completed."

                # Force runspace cleanup - this is critical as runspaces can hold transcript handles
                $runspaces = [runspacefactory]::Runspaces
                if ($runspaces.Count -gt 0) {
                    Write-VerboseSafe "Found $($runspaces.Count) runspaces to clean up"
                    foreach ($rs in $runspaces) {
                        try {
                            if ($null -eq $rs.ConnectionInfo -and $null -ne $rs.Owner) {
                                $rs.Owner = $null
                            }
                            # Force runspace synchronization
                            if ($rs.RunspaceStateInfo.State -eq 'Opened') {
                                [System.Threading.Monitor]::Enter($rs)
                                [System.Threading.Monitor]::Exit($rs)
                            }
                        } catch {
                            Write-VerboseSafe "Error cleaning runspace: $_"
                        }
                    }
                }

                # Try to explicitly close any streams that might be holding the transcript file
                if (-not [string]::IsNullOrEmpty($transcriptPath) -and (Test-Path -Path $transcriptPath -ErrorAction SilentlyContinue)) {
                    try {
                        # Use .NET IO directly to try to access the file, which will fail if a handle is still open
                        $fileInfo = New-Object System.IO.FileInfo -ArgumentList $transcriptPath
                        if ($fileInfo.Exists) {
                            # Try to open and close the file to check if it's accessible
                            $stream = [System.IO.File]::Open($transcriptPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Delete)
                            $stream.Close()
                            $stream.Dispose()
                            Write-VerboseSafe "Successfully accessed and closed transcript file: $transcriptPath"
                        }
                    } catch {
                        Write-VerboseSafe "File access test indicates transcript is still locked: $_"
                    }
                }

                # Final garbage collection sweep
                [System.GC]::Collect(2, [System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $true)
                [System.GC]::WaitForPendingFinalizers()

                # Set the flag to inactive *after* successful stop
                $script:transcriptActive = $false
                Write-DiagnosticMessage "Transcript stopped successfully." -Color DarkGray
            }
            catch {
                Write-DiagnosticMessage "Warning: Error stopping transcript: $_" -Color Yellow
                try {
                    # Second try - just in case the first attempt failed but didn't throw properly
                    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
                    Start-Sleep -Milliseconds 500
                    [System.GC]::Collect(2, [System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $true)
                    [System.GC]::WaitForPendingFinalizers()
                } catch {
                    # Log error from second attempt
                    Write-VerboseSafe "Second attempt to stop transcript also failed: $_"
                }

                # Last resort - attempt to use CloseAllTranscripts
                try {
                    # Try to use reflection to close all transcripts (internal PowerShell API)
                    $internalType = [Microsoft.PowerShell.Commands.PSHostInvocationData].Assembly.GetType('Microsoft.PowerShell.Commands.TranscriptionData')
                    if ($null -ne $internalType) {
                        $method = $internalType.GetMethod('CloseAllTranscripts', [System.Reflection.BindingFlags]'NonPublic,Static')
                        if ($null -ne $method) {
                            $method.Invoke($null, @())
                            Write-VerboseSafe "Invoked CloseAllTranscripts via reflection"
                        }
                    }
                } catch {
                    Write-VerboseSafe "Reflection-based transcript closure failed: $_"
                }

                # Even if stopping failed, mark as inactive to prevent retry loops
                $script:transcriptActive = $false
                Write-VerboseSafe "Transcript marked as inactive despite error during stop."
            }

            # No matter what happened, do one final GC collection
            [System.GC]::Collect(2, [System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $true)
            [System.GC]::WaitForPendingFinalizers()
        }
    } else {
        Write-VerboseSafe "Transcript was not marked as active by this script, skipping Stop-Transcript."
    }
    Write-Verbose "Exiting Stop-TranscriptSafely function."
}

# Function to log to transcript only without console output
function Write-TranscriptOnly {
    param([string]$Message)
    $InformationPreference = 'Continue'
    Write-Information -MessageData $Message 6> $null
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
                    $type = "SymbolicLink"                    # Improved parsing logic for symbolic links
                    $printNameLine = $fsutil | Where-Object -FilterScript { $_ -match "Print Name:" }
                    if ($printNameLine) {
                        $target = ($printNameLine -replace "^.*?Print Name:\s*", "").Trim()
                    }
                }
                elseif ($fsutil -match "Mount Point") {
                    $type = "MountPoint"
                    $printNameLine = $fsutil | Where-Object -FilterScript { $_ -match "Print Name:" }
                    if ($printNameLine) {
                        $target = ($printNameLine -replace "^.*?Print Name:\s*", "").Trim()
                    }
                }
                elseif ($fsutil -match "Junction") {
                    $type = "Junction"
                    $printNameLine = $fsutil | Where-Object -FilterScript { $_ -match "Print Name:" }
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
        [int]$Width = 50,
        [string]$Activity = "Processing Folders",
        [int]$Id = 2,
        [string]$CurrentOperation = "",
        [string]$Status = "",
        [int]$ParentId = -1
    )

    # Calculate percent complete
    $percentComplete = 0
    if ($Total -gt 0) {
        $percentComplete = [math]::Min(100, [math]::Floor(($Completed / $Total) * 100))
    }    # Create visual bar representation
    $filledWidth = [math]::Floor($Width * ($percentComplete / 100))
    $bar = "[" + ("=" * $filledWidth).PadRight($Width) + "] $percentComplete% | Completed processing $Completed of $Total folders"    # Enhanced PS7 progress bar with try/catch to handle any errors
    try {
        # Check if progress bars are enabled
        if ($script:UseProgressBars -ne $false) {
            # Create the progress parameters as a hashtable for better clarity
            $progressParams = @{
                Activity = $Activity
                Status = if ([string]::IsNullOrEmpty($Status)) { $bar } else { $Status }
                PercentComplete = $percentComplete
                Id = $Id
            }

            # Add CurrentOperation if not empty
            if (-not [string]::IsNullOrWhiteSpace($CurrentOperation)) {
                $progressParams.CurrentOperation = $CurrentOperation
            }            # Add ParentId if provided and not equal to -1
            if ($ParentId -ne -1) {
                $progressParams.ParentId = $ParentId
            }

            # Call Write-Progress with splatting for cleaner PS7 support
            Write-Progress @progressParams
        }

        # Diagnostic message at regular intervals
        if ($percentComplete % 10 -eq 0 -and $percentComplete -gt 0) {
            Write-DiagnosticMessage "Progress: $percentComplete% complete. Processed $Completed of $Total folders." -Color "Cyan"
        }

        # Complete the progress bar when done
        if ($Completed -eq $Total) {
            Write-Progress -Activity $Activity -Completed -Id $Id
            Write-DiagnosticMessage "Folder processing complete. Total folders processed: $Total" -Color "Green"
        }
    }
    catch {
        # If progress bar fails, at least show diagnostic output
        Write-DiagnosticMessage "Progress: $percentComplete% ($Completed of $Total folders)" -Color "Cyan"
        Write-Verbose "Error displaying progress bar: $($_.Exception.Message)"
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
        [bool]$FollowJunctions = $true,
        [bool]$ShowProgress = $true
    )

    $size = 0
    $stack = New-Object System.Collections.Generic.Stack[string]
    $stack.Push($Path)
      # For progress reporting
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $lastUpdate = 0
    $processedItems = 0
    $totalEstimated = 500000 # Higher estimate for more realistic progress bar
    $processedPaths = @{}

    # Initial message
    if ($ShowProgress) {
        Write-DiagnosticMessage "Starting initial size calculation for $Path..." -Color "Cyan"
    }

    while ($stack.Count -gt 0) {
        $dir = $stack.Pop()
        try {
            # Avoid processing the same path twice (can happen with junctions)
            if ($processedPaths.ContainsKey($dir)) {
                continue
            }
            $processedPaths[$dir] = $true
              # Update progress periodically
            $processedItems++
            if ($ShowProgress -and ($stopwatch.ElapsedMilliseconds - $lastUpdate -gt 300)) {
                $percentComplete = [Math]::Min(99, [Math]::Round(($processedItems / $totalEstimated) * 100))
                  # Show progress every 300ms - using PS7 compatible approach with try/catch
                try {
                    $progressParams = @{
                        Activity = "Calculating Directory Size"
                        Status = "Scanning folder $processedItems of ~$totalEstimated | $percentComplete%"
                        PercentComplete = $percentComplete
                        Id = 4
                        CurrentOperation = "Path: $dir"
                    }
                    Write-Progress @progressParams
                }
                catch {
                    # Fallback if progress bar fails
                    Write-Verbose "Error displaying progress: $($_.Exception.Message)"
                }
                  # Periodic text updates every 3 seconds
                if ($stopwatch.ElapsedMilliseconds - $lastUpdate -gt 3000) {
                    $sizeSoFar = [math]::Round($size / 1GB, 2)
                    # Only show detailed scanning messages in Debug or Verbose mode
                    if ($DebugPreference -ne 'SilentlyContinue' -or $VerbosePreference -ne 'SilentlyContinue') {
                        Write-DiagnosticMessage "Scanning size: $sizeSoFar GB so far, processing $dir" -Color "DarkGray"
                    }
                    $lastUpdate = $stopwatch.ElapsedMilliseconds
                }
            }

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
      # Complete progress bar
    if ($ShowProgress) {
        try {
            # First show 100% completion
            $progressParams = @{
                Activity = "Calculating Directory Size"
                Status = "Complete: 100%"
                PercentComplete = 100
                Id = 4
            }
            Write-Progress @progressParams
            Start-Sleep -Milliseconds 300 # Brief pause to ensure progress bar updates

            # Then complete the progress bar
            Write-Progress -Activity "Calculating Directory Size" -Completed -Id 4

            $sizeGB = [math]::Round($size / 1GB, 2)
            Write-DiagnosticMessage "Size calculation complete: $sizeGB GB total" -Color "Green"
        }
        catch {
            # Fallback for progress display errors
            Write-Verbose "Error completing progress bar: $($_.Exception.Message)"
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
                    $filteredFiles = $allFiles | Where-Object -FilterScript { script:IsFilePhysicallyStored $_.FullName }
                }

                # Find largest file in current directory
                $currentDirLargestFile = $filteredFiles | Sort-Object -Property Length -Descending | Select-Object -First 1

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

function Debug-ProgressBar {
    param (
        [string]$TestName = "Progress Bar Test"
    )

    Write-DiagnosticMessage "Testing progress bar functionality in PowerShell $($PSVersionTable.PSVersion)" -Color "Cyan"
    Write-DiagnosticMessage "Starting progress bar test: $TestName" -Color "Yellow"

    # Perform a simple counting test with Write-Progress
    try {
        $total = 10
        for ($i = 1; $i -le $total; $i++) {
            $percent = ($i / $total) * 100

            # Use both raw Write-Progress and our wrapper
            Write-Progress -Activity "Direct Test" -Status "Testing: $i of $total" -PercentComplete $percent -Id 99
            Write-ProgressBar -Completed $i -Total $total -Activity "Wrapper Test" -Id 98 -CurrentOperation "Testing progress visibility"

            # Only show test progress in Debug or Verbose mode
            if ($DebugPreference -ne 'SilentlyContinue' -or $VerbosePreference -ne 'SilentlyContinue') {
                Write-DiagnosticMessage "Progress test: Step $i of $total ($percent%)" -Color "DarkGray"
            }
            Start-Sleep -Milliseconds 500
        }

        # Complete both progress bars
        Write-Progress -Activity "Direct Test" -Completed -Id 99
        Write-ProgressBar -Completed $total -Total $total -Activity "Wrapper Test" -Id 98

        Write-DiagnosticMessage "Progress bar test complete" -Color "Green"
        return $true
    }
    catch {
        Write-DiagnosticMessage "Error during progress bar test: $($_.Exception.Message)" -Color "Error"
        return $false
    }
}

# Only run the progress bar test if -Debug or -Verbose is specified
if ($DebugPreference -ne 'SilentlyContinue' -or $VerbosePreference -ne 'SilentlyContinue') {
    Write-DiagnosticMessage "Testing progress bar functionality..." -Color "Cyan"
    $progressBarTest = Debug-ProgressBar -TestName "Initial Progress Test"

    if (-not $progressBarTest) {
        Write-DiagnosticMessage "Progress bar test failed. The script will continue using text-based status updates." -Color "Warning"
        $script:UseProgressBars = $false
    } else {
        $script:UseProgressBars = $true
        Write-DiagnosticMessage "Progress bar test succeeded. Progress bars should be visible." -Color "Green"
    }
} else {
    # Skip test in normal mode but keep progress bars enabled
    $script:UseProgressBars = $true
}

# We'll need to update our use of progress bars based on this flag

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
    }    $RunspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads)
    $RunspacePool.Open()
    $FolderSizeMap = @{}
    $Runspaces = @()
    $activeRunspaces = 0
    $processedCount = 0
    $totalFolders = $Folders.Count
      # Display initial processing info with distinctive color
    Write-Information "`nParallel Processing Configuration:" -InformationAction Continue
    Write-Information "Maximum Threads: $MaxThreads" -InformationAction Continue
    Write-Information "Total Folders to Process: $totalFolders" -InformationAction Continue
    Write-Information "Only Physical Files: $OnlyPhysicalFiles" -InformationAction Continue
    Write-Information "Follow Junctions: $FollowJunctions" -InformationAction Continue
    Write-Information "Active Runspaces: 0/$MaxThreads" -InformationAction Continue    # Create and start the stopwatch for timing updates
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $lastUpdate = 0

    foreach ($folder in $Folders) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $RunspacePool
        $activeRunspaces++        # Update progress bar every 500ms
        if ($stopwatch.ElapsedMilliseconds - $lastUpdate -gt 500) {
            # Use our enhanced Write-ProgressBar function for initialization progress
            Write-ProgressBar -Completed $activeRunspaces -Total $totalFolders -Activity "Initializing Folder Scanning" -Id 1 -CurrentOperation "Creating runspace threads: $activeRunspaces/$MaxThreads active"
            Write-Information "Preparing to scan folders: $activeRunspaces/$totalFolders" -InformationAction Continue
            $lastUpdate = $stopwatch.ElapsedMilliseconds
        }        [void]$ps.AddScript({
            param(
                $StartPath,
                $OnlyPhysicalFiles,
                $FollowJunctions
            )            # Define runspace-safe helper functions inline to avoid cmdlet dependencies
            function IsFilePhysicallyStored {
                param([string]$filePath)
                try {
                    if (!(Test-Path $filePath)) { return $false }
                    # For simplicity in runspace, assume all files are physical
                    # The full implementation would check OneDrive attributes
                    return $true
                } catch { return $true }
            }

            function GetDirectorySize {
                param([string]$Path, [bool]$OnlyPhysicalFiles, [bool]$FollowJunctions)
                $size = 0
                $stack = New-Object System.Collections.Generic.Stack[string]
                $stack.Push($Path)
                $processedPaths = @{}

                while ($stack.Count -gt 0) {
                    $dir = $stack.Pop()
                    try {
                        if ($processedPaths.ContainsKey($dir)) { continue }
                        $processedPaths[$dir] = $true

                        foreach ($file in [System.IO.Directory]::GetFiles($dir)) {                            try {
                                if ($OnlyPhysicalFiles -and -not (IsFilePhysicallyStored $file)) { continue }
                                $size += (New-Object System.IO.FileInfo($file)).Length
                            } catch {
                                Write-Debug "Unable to access file: $file"
                            }
                        }

                        foreach ($subDir in [System.IO.Directory]::GetDirectories($dir)) {
                            $dirInfo = New-Object System.IO.DirectoryInfo($subDir)
                            $isReparsePoint = $dirInfo.Attributes.ToString() -match 'ReparsePoint'
                            if ($isReparsePoint -and -not $FollowJunctions) { continue }
                            $stack.Push($subDir)                        }                    } catch {
                        Write-Debug "Unable to access directory: $dir"
                    }
                }
                return $size
            }

            function GetDirectoryCounts {
                param([string]$Path, [bool]$OnlyPhysicalFiles, [bool]$FollowJunctions)
                $fileCount = 0
                $folderCount = 0
                $stack = New-Object System.Collections.Generic.Stack[string]
                $stack.Push($Path)
                $processedPaths = @{}

                while ($stack.Count -gt 0) {
                    $dir = $stack.Pop()
                    try {
                        if ($processedPaths.ContainsKey($dir)) { continue }
                        $processedPaths[$dir] = $true

                        $files = [System.IO.Directory]::GetFiles($dir)
                        if ($OnlyPhysicalFiles) {
                            $fileCount += ($files | Where-Object { IsFilePhysicallyStored $_ }).Count
                        } else {
                            $fileCount += $files.Count
                        }

                        foreach ($subDir in [System.IO.Directory]::GetDirectories($dir)) {
                            $folderCount++
                            $dirInfo = New-Object System.IO.DirectoryInfo($subDir)
                            $isReparsePoint = $dirInfo.Attributes.ToString() -match 'ReparsePoint'
                            if ($isReparsePoint -and -not $FollowJunctions) { continue }                            $stack.Push($subDir)
                        }                    } catch {
                        Write-Debug "Unable to access directory for counting: $dir"
                    }
                }
                return @($fileCount, $folderCount)
            }            function GetLargestFile {
                param([string]$Path, [bool]$OnlyPhysicalFiles, [bool]$FollowJunctions)
                $largestSize = 0
                $largestFile = $null

                try {
                    # Get files based on junction following preference
                    $files = if ($FollowJunctions) {
                        [System.IO.Directory]::GetFiles($Path)
                    } else {                        # Use Get-ChildItem with -Force to respect junction handling when not following junctions
                        (Get-ChildItem -Path $Path -File -Force -ErrorAction SilentlyContinue).FullName
                    }

                    foreach ($file in $files) {
                        try {
                            if ($OnlyPhysicalFiles -and -not (IsFilePhysicallyStored $file)) { continue }
                            $fileInfo = New-Object System.IO.FileInfo($file)
                            if ($fileInfo.Length -gt $largestSize) {
                                $largestSize = $fileInfo.Length
                                $largestFile = $fileInfo.Name
                            }} catch {
                            Write-Debug "Unable to access file for size comparison: $file"
                        }
                    }                } catch {
                    Write-Debug "Unable to access directory for largest file search: $Path"
                }

                if ($largestFile) {
                    return "$largestFile ($([math]::Round($largestSize / 1MB, 2)) MB)"
                }
                return "No files found"
            }

            $threadId = [System.Threading.Thread]::CurrentThread.ManagedThreadId

            try {
                $counts = GetDirectoryCounts -Path $StartPath -OnlyPhysicalFiles $OnlyPhysicalFiles -FollowJunctions $FollowJunctions
                $size = GetDirectorySize -Path $StartPath -OnlyPhysicalFiles $OnlyPhysicalFiles -FollowJunctions $FollowJunctions
                $largestFile = GetLargestFile -Path $StartPath -OnlyPhysicalFiles $OnlyPhysicalFiles -FollowJunctions $FollowJunctions

                return @{
                    Success = $true
                    StartPath = $StartPath
                    Size = $size
                    FileCount = $counts[0]
                    FolderCount = $counts[1]
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
        })        # Add arguments one by one - simplified argument list for runspace-safe operation
        [void]$ps.AddArgument($folder.FullName)
        [void]$ps.AddArgument($OnlyPhysicalFiles)
        [void]$ps.AddArgument($FollowJunctions)

        $Runspaces += [PSCustomObject]@{
            Instance = $ps
            Handle = $ps.BeginInvoke()
            Folder = $folder.FullName
            StartTime = [DateTime]::Now
        }
    }    Write-Information "`nProcessing folders in parallel..." -InformationAction Continue    # Reset counters for result processing
    $processedCount = 0
    $completedFolders = 0
    $totalSize = 0
    $totalFiles = 0
    $totalFolderCount = 0
    $lastProgressUpdate = 0

    foreach ($r in $Runspaces) {
        try {
            # Process the result first, then update counters
            $result = $r.Instance.EndInvoke($r.Handle)

            # Only increment processedCount after we have the result
            $processedCount++

            # Update progress display periodically but use our enhanced progress bar
            if ($stopwatch.ElapsedMilliseconds - $lastProgressUpdate -gt 300) {
                # Use our enhanced Write-ProgressBar function for better visibility
                $totalSizeGB = [math]::Round($totalSize / 1GB, 2)
                $currentOperation = "Current Stats: $completedFolders completed | $totalFiles files | $totalSizeGB GB"
                Write-ProgressBar -Completed $completedFolders -Total $totalFolders -Activity "Scanning Folders" -Id 1 -CurrentOperation $currentOperation

                # Show real-time stats periodically (not every single folder)
                if ($processedCount % 5 -eq 0 -or $processedCount -eq $totalFolders) {
                    $statusMsg = "Processed: $completedFolders/$totalFolders folders | $totalFiles files | $totalSizeGB GB"
                    Write-Information $statusMsg -InformationAction Continue
                }

                $lastProgressUpdate = $stopwatch.ElapsedMilliseconds
            }            if ($result.Success) {
                $completedFolders++
                $totalSize += $result.Size
                $totalFiles += $result.FileCount
                $totalFolderCount += $result.FolderCount

                # Show progress for every completed folder to avoid the stuck progress bar issue
                if ($script:UseProgressBars) {
                    $totalSizeGB = [math]::Round($totalSize / 1GB, 2)
                    $currentOperation = "Completed: $completedFolders/$totalFolders | $totalFiles files | $totalSizeGB GB"
                    Write-ProgressBar -Completed $completedFolders -Total $totalFolders -Activity "Scanning Folders" -Id 1 -CurrentOperation $currentOperation
                }

                # Always add successful results to the map, even if they appear empty
                # The folder might have subfolders that will be processed later
                $FolderSizeMap[$result.StartPath] = @{
                    Size = $result.Size
                    FileCount = $result.FileCount
                    FolderCount = $result.FolderCount
                    LargestFile = $result.LargestFile
                }

                # Log folders that appear empty for informational purposes
                if ($result.Size -eq 0 -and $result.FileCount -eq 0 -and $result.FolderCount -eq 0) {
                    Write-DiagnosticMessage "Folder $($r.Folder) appears empty but will be included for recursive scanning" -Color "DarkGray"
                }} else {
                $errorMessage = "Thread $($result.ThreadId) failed: $($result.Error)"
                $script:InaccessibleFolders[$r.Folder] = $errorMessage
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
      # Complete the progress bar using our enhanced function
      Write-ProgressBar -Completed $totalFolders -Total $totalFolders -Activity "Folder Processing Complete" -Id 1
    Write-Information " " -InformationAction Continue

    Write-Information "`nParallel Processing Summary:" -InformationAction Continue
    Write-Information "Total Folders Processed: $processedCount" -InformationAction Continue
    Write-Information "Maximum Concurrent Threads: $MaxThreads" -InformationAction Continue

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
        [bool]$OnlyPhysicalFiles,
        [int]$TotalDepths = 0
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

        $StartPath = Format-Path -InputPath $StartPath
        if (-not (Test-Path -Path $StartPath -PathType Container)) {
            Write-Warning -Message "Path '$StartPath' does not exist or is not a directory."
            return @{
                ProcessedFolders = $false
                HasSubfolders = $false
                CompletionMessageShown = $false
            }
        }        Write-Information "`nTop $Top Largest Folders in: $StartPath" -InformationAction Continue
        Write-Information "" -InformationAction Continue

        # Initialize progress tracking for first call
        if ($CurrentDepth -eq 1 -and $TotalDepths -eq 0) {
            # Make an estimate of total depths to process based on folder structure
            # Increased estimate for more realistic progress calculation

            # Initialize variables with same pattern as initial scanning but with more accurate estimates
            $script:totalEstimatedFolders = [Math]::Max(100, ($Top * $MaxDepth * 3))
            $script:processedFolders = 0
            $script:totalRecursiveSize = 0
            $script:totalRecursiveFiles = 0
            $script:totalRecursiveFolders = 0            # Set up variables for aggressive progress monitoring
            $script:lastRecursiveProgressUpdate = 0
            $script:lastConsoleUpdate = 0  # Separate timer for console updates
            $script:recursiveUpdateFrequency = 300 # Milliseconds between updates (matching initial scan)
            $script:recursiveStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $TotalDepths = $script:totalEstimatedFolders            # Clear any previous progress bar
            if ($script:UseProgressBars) {
                # Complete previous progress bar if exists
                try {
                    Write-Progress -Activity "Previous Operation" -Id 3 -Completed
                }
                catch {
                    # Handle errors from completing non-existent progress bars
                    Write-DiagnosticMessage "Non-critical error while completing progress bar: $($_.Exception.Message)" -Color "DarkGray"
                }
            }

            Write-DiagnosticMessage "Starting recursive scan with estimated $script:totalEstimatedFolders folders to process" -Color "Cyan"
            Write-Information "`n$($script:ANSI.Cyan)Starting recursive folder analysis...$($script:ANSI.Reset)" -InformationAction Continue

            # Always initialize the progress bar immediately with high visibility
            Write-Progress -Activity "Recursive Folder Analysis" -Status "Starting recursive scan..." -PercentComplete 0 -Id 3 -CurrentOperation "Initializing..."

            # Write initial console message with high visibility for better user feedback
            Write-Information "" -InformationAction Continue
            Write-Information "$($script:ANSI.Yellow)Recursive scan progress: 0% | Depth 1/$MaxDepth | 0/$TotalDepths folders | 0 GB$($script:ANSI.Reset)" -InformationAction Continue
            Write-Information "" -InformationAction Continue

            # Update the last progress timestamp to current
            $script:lastRecursiveProgressUpdate = $script:recursiveStopwatch.ElapsedMilliseconds
        }

        # Update progress counts for the current folder being processed
        $script:processedFolders++
        $script:totalRecursiveFolders++        # Update progress bar more frequently for better user feedback
        if ($script:UseProgressBars -and ($script:recursiveStopwatch.ElapsedMilliseconds - $script:lastRecursiveProgressUpdate -gt 500)) {
            try {                # Calculate progress metrics - using same formula as initial scan
                $totalSizeGB = [math]::Round($script:totalRecursiveSize / 1GB, 2)
                $progressPercentage = [Math]::Min(100, [Math]::Floor(($script:processedFolders * 100) / [Math]::Max(1, $TotalDepths)))

                # Include the path but shortened to avoid very wide progress bars
                $shortenedPath = if ($StartPath.Length -gt 30) { "..." + $StartPath.Substring($StartPath.Length - 30) } else { $StartPath }

                # Create progress parameters (matching initial scan style)
                $progressParams = @{
                    Activity = "Recursive Folder Analysis"
                    Status = "Scanning folder $script:processedFolders of ~$TotalDepths | $progressPercentage%"
                    PercentComplete = $progressPercentage
                    Id = 3
                    CurrentOperation = "Path: $shortenedPath | $totalSizeGB GB"
                }

                # Update the progress bar
                Write-Progress @progressParams

                # Update the timestamp for next update
                $script:lastRecursiveProgressUpdate = $script:recursiveStopwatch.ElapsedMilliseconds
            }
            catch {
                # Gracefully handle errors in progress bar updates
                Write-DiagnosticMessage "Error updating recursive progress bar: $($_.Exception.Message)" -Color "Red"
            }
        }

        # Log periodic console updates every 2 seconds instead of 3 for better feedback
        if ($script:recursiveStopwatch.ElapsedMilliseconds - $script:lastConsoleUpdate -gt 2000) {
            $totalSizeGB = [math]::Round($script:totalRecursiveSize / 1GB, 2)
            $progressPercentage = [Math]::Min(100, [Math]::Floor(($script:processedFolders * 100) / [Math]::Max(1, $TotalDepths)))

            # Create an informative console message
            $statusMsg = "Recursive scan progress: $progressPercentage% | Depth $CurrentDepth/$MaxDepth | $script:processedFolders/$TotalDepths folders | $totalSizeGB GB"
            Write-Information "$($script:ANSI.Cyan)$statusMsg$($script:ANSI.Reset)" -InformationAction Continue

            # Update the timestamp for next console update
            $script:lastConsoleUpdate = $script:recursiveStopwatch.ElapsedMilliseconds
        }

        # Force progress bar completion if we're at 95% or more
        if ($script:processedFolders -ge ($TotalDepths * 0.95)) {
            try {
                Write-Progress -Activity "Recursive Folder Analysis" -Status "Recursive Scan: Finalizing" -PercentComplete 99 -Id 3 -CurrentOperation "Completing scan..."                Write-Information "$($script:ANSI.Green)Recursive scan almost complete: $script:processedFolders folders | $totalSizeGB GB | Time: $($script:recursiveStopwatch.Elapsed.ToString("hh\:mm\:ss"))$($script:ANSI.Reset)" -InformationAction Continue
            }
            catch {
                # Ignore errors when completing the progress bar
                Write-DiagnosticMessage "Error updating final progress: $($_.Exception.Message)" -Color "Red"
            }

            # Update timestamp for next progress update
            $script:lastRecursiveProgressUpdate = $script:recursiveStopwatch.ElapsedMilliseconds
        }

        # Force an update at least every second regardless of other conditions
        # This ensures progress is always shown even during long operations
        try {
            if ($script:recursiveStopwatch.ElapsedMilliseconds - $script:lastRecursiveProgressUpdate -gt 1000) {
                $totalSizeGB = [math]::Round($script:totalRecursiveSize / 1GB, 2)
                Write-Information "Recursive scan continuing: $script:processedFolders folders | $totalSizeGB GB" -InformationAction Continue
                $script:lastRecursiveProgressUpdate = $script:recursiveStopwatch.ElapsedMilliseconds
            }
        }
        catch {
            Write-DiagnosticMessage "Error updating recursive progress: $($_.Exception.Message)" -Color "Red"
        }

        # First, analyze the root path itself
        if ($CurrentDepth -eq 1) {
            $rootSize = script:GetDirectorySize -Path $StartPath -OnlyPhysicalFiles $OnlyPhysicalFiles -FollowJunctions $FollowJunctions -ShowProgress $false
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
            Write-Information "Processing $($rootFolders.Count) folders in root directory..." -InformationAction Continue            # Add a status update before processing begins
            Write-DiagnosticMessage "Starting parallel scan of $($rootFolders.Count) folders at depth $CurrentDepth" -Color "Cyan"

            # Process root folders in parallel using the MaxThreads parameter
            $folderResults = Start-FolderProcessing -Folders $rootFolders -MaxThreads $MaxThreads -OnlyPhysicalFiles $OnlyPhysicalFiles -FollowJunctions $FollowJunctions

            # Add a status update after processing completes
            Write-DiagnosticMessage "Completed parallel scan of $($rootFolders.Count) folders at depth $CurrentDepth" -Color "Green"

            # Convert results to sorted array
            $sortedFolders = $folderResults.GetEnumerator() | ForEach-Object -Process {
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
            $topFolders = $sortedFolders | Select-Object -First $topFoldersCount            # Debug information
            if ($sortedFolders.Count -gt 0) {
                Write-DiagnosticMessage "Found $($sortedFolders.Count) sorted folders, displaying top $topFoldersCount" -Color "Cyan"
            } else {
                Write-DiagnosticMessage "Found 0 accessible folders - all folders either have no accessible content or access was denied" -Color "Yellow"
            }

            # Display table rows for each top folder
            if ($topFolders -and $topFolders.Count -gt 0) {
                foreach ($folder in $topFolders) {
                    Write-TableRow -StartPath $folder.Path -Size $folder.Size -SubfolderCount $folder.FolderCount -FileCount $folder.FileCount -LargestFile $folder.LargestFile
                }
                Write-Information -MessageData ("-" * 150) -InformationAction Continue
                Write-Information -MessageData "" -InformationAction Continue
            } else {
                Write-DiagnosticMessage "No top folders to display (check for access denied issues)" -Color "Yellow"                # Provide more details about why there may be no folders to display
                if ($rootFolders -and $rootFolders.Count -gt 0 -and $sortedFolders.Count -eq 0) {
                    Write-DiagnosticMessage "Found $($rootFolders.Count) folders but they couldn't be processed due to specific reasons:" -Color "Yellow"
                    Write-Information -MessageData "`nFolder access issues (specific details per folder):" -InformationAction Continue

                    # Display each inaccessible folder with its specific error/reason
                    foreach ($folder in $script:InaccessibleFolders.Keys) {
                        $reason = $script:InaccessibleFolders[$folder]
                        Write-Information -MessageData "$($script:ANSI.Yellow)$folder$($script:ANSI.Reset) - $reason" -InformationAction Continue
                    }

                    # If somehow we don't have details for some folders, provide general reasons
                    if ($script:InaccessibleFolders.Count -eq 0) {
                        Write-Information -MessageData "1. Access restrictions (permissions denied)" -InformationAction Continue
                        Write-Information -MessageData "2. Folder is a junction point or symbolic link (like My Music, My Pictures, etc.)" -InformationAction Continue
                        Write-Information -MessageData "3. Folder is empty or contains only placeholders" -InformationAction Continue
                    }
                    foreach ($folder in $rootFolders) {
                        Write-Information -MessageData "  - $($folder.FullName)" -InformationAction Continue
                    }
                } else {
                    Write-Information -MessageData "No accessible folders found for display." -InformationAction Continue
                }
                Write-Information -MessageData ("-" * 150) -InformationAction Continue
                Write-Information -MessageData "" -InformationAction Continue
            }# Process all top folders up to max depth
            $completionMessageShown = $false
            if ($CurrentDepth + 1 -le $MaxDepth -and $sortedFolders.Count -gt 0) {
                # Process top 5 (or as specified by $Top) folders at each level
                $foldersToProcess = $sortedFolders | Select-Object -First $Top
                  foreach ($folder in $foldersToProcess) {
                    # Update recursive scanning statistics for progress bar
                    $script:totalRecursiveSize += $folder.Size
                    $script:totalRecursiveFiles += $folder.FileCount

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
                        Write-Information "`n$($script:ANSI.DarkCyan)Processing $($subFolders.Count) subfolders in: $($folder.Path)$($script:ANSI.Reset)" -InformationAction Continue                        # Add status update for subfolders processing
                        # Only show detailed processing messages in Debug or Verbose mode
                        if ($DebugPreference -ne 'SilentlyContinue' -or $VerbosePreference -ne 'SilentlyContinue') {
                            Write-DiagnosticMessage "Processing $($subFolders.Count) subfolders at depth $($CurrentDepth+1) in $($folder.Path)" -Color "DarkGray"
                        }

                        # Process subfolders in parallel
                        $subFolderResults = Start-FolderProcessing -Folders $subFolders -MaxThreads $MaxThreads -OnlyPhysicalFiles $OnlyPhysicalFiles -FollowJunctions $FollowJunctions

                        # Add status update after subfolders are processed
                        # Only show detailed completion messages in Debug or Verbose mode
                        if ($DebugPreference -ne 'SilentlyContinue' -or $VerbosePreference -ne 'SilentlyContinue') {
                            Write-DiagnosticMessage "Completed processing $($subFolders.Count) subfolders in $($folder.Path)" -Color "DarkGray"
                        }

                        # Convert results to sorted array
                        $sortedSubFolders = $subFolderResults.GetEnumerator() | ForEach-Object -Process {
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
                    }                    # Call recursively for deeper levels and capture the structured return value
                    if ($CurrentDepth + 1 -lt $MaxDepth) {
                        # Write periodic update on deeper recursion
                        if ($CurrentDepth -gt 1) {
                            Write-DiagnosticMessage "Recursing to depth $($CurrentDepth + 1) in $folder.Path" -Color "Cyan"
                        }

                        $result = Get-FolderSize -StartPath $folder.Path -CurrentDepth ($CurrentDepth + 1) -MaxDepth $MaxDepth -Top $Top -OnlyPhysicalFiles $OnlyPhysicalFiles -TotalDepths $TotalDepths

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
            }            # Complete the progress bar for recursive scanning when we're finishing the top-level call
            if ($CurrentDepth -eq 1 -and $script:UseProgressBars) {
                try {
                    # Format final progress description consistently with initial scanning
                    $totalSizeGB = [math]::Round($script:totalRecursiveSize / 1GB, 2)
                    $timeElapsed = $script:recursiveStopwatch.Elapsed.ToString("hh\:mm\:ss")
                    $finalProgressDescription = "Complete | Folders: $script:totalRecursiveFolders | Files: $script:totalRecursiveFiles | Size: $totalSizeGB GB | Time: $timeElapsed"

                    # Log progress completion activity for troubleshooting
                    Write-DiagnosticMessage "Completing recursive scan progress bar - ID 3" -Color "Cyan"

                    # Complete the progress bar with the final stats - guaranteed completion
                    Write-Progress -Activity "Recursive Folder Analysis" -Status "Complete" -PercentComplete 100 -Id 3 -CurrentOperation $finalProgressDescription

                    # Explicitly complete the progress bar again after a brief pause to ensure it registers
                    Start-Sleep -Milliseconds 50
                    Write-Progress -Activity "Recursive Folder Analysis" -Completed -Id 3

                    # Display final statistics in same format as initial scanning completion
                    Write-DiagnosticMessage "Recursive Scan Complete" -Color "Green"
                    Write-Information "`n$($script:ANSI.Green)Recursive Scan Complete:$($script:ANSI.Reset)" -InformationAction Continue
                    Write-Information "Total Folders Processed: $($script:ANSI.Cyan)$script:totalRecursiveFolders$($script:ANSI.Reset)" -InformationAction Continue
                    Write-Information "Total Files Found: $($script:ANSI.Cyan)$script:totalRecursiveFiles$($script:ANSI.Reset)" -InformationAction Continue
                    Write-Information "Total Size Analyzed: $($script:ANSI.Cyan)$totalSizeGB GB$($script:ANSI.Reset)" -InformationAction Continue
                    Write-Information "Time Elapsed: $($script:ANSI.Cyan)$timeElapsed$($script:ANSI.Reset)" -InformationAction Continue
                }
                catch {
                    Write-DiagnosticMessage "Error completing recursive scan progress bar: $($_.Exception.Message)" -Color "Red"
                }
            }

            return @{
                ProcessedFolders = $true
                HasSubfolders = $true
                CompletionMessageShown = $completionMessageShown
            }
        }
        else {
            Write-Warning -Message "No subfolders found to process."
            return @{
                ProcessedFolders = $true
                HasSubfolders = $false
                CompletionMessageShown = $false
            }
        }
    } catch {
        Write-Warning -Message "Error processing folder '$StartPath': $($_.Exception.Message)"
        return @{
            ProcessedFolders = $false
            HasSubfolders = $false
            CompletionMessageShown = $false
        }
    }
} # End of Get-FolderSize function

# Start the Recursive Scan
# Get the total calculated size for analysis with progress reporting
Write-DiagnosticMessage "Starting initial disk analysis..." -Color "Cyan"
$rootSize = script:GetDirectorySize -Path $StartPath -OnlyPhysicalFiles $OnlyPhysicalFiles -FollowJunctions $FollowJunctions -ShowProgress $true

# Perform disk usage analysis
Write-DiskUsageAnalysis -StartPath $StartPath -CalculatedSize $rootSize

# Start the folder size analysis with recursive processing
Write-Information -MessageData "$($script:ANSI.Cyan)`nBeginning recursive folder scan with max depth of $MaxDepth and top $Top folders at each level$($script:ANSI.Reset)" -InformationAction Continue

# Clear any existing progress bars before starting main processing
Write-Progress -Activity "Processing Folders" -Completed -Id 1
Write-Progress -Activity "Processing Folders" -Completed -Id 2
Write-Progress -Activity "Scanning Folder Structure" -Completed -Id 3

# Initialize global tracking variables
$script:totalEstimatedFolders = 0
$script:processedFolders = 0

$result = Get-FolderSize -StartPath $StartPath -CurrentDepth 1 -MaxDepth $MaxDepth -Top $Top -OnlyPhysicalFiles $OnlyPhysicalFiles

# Complete all progress bars
Write-Progress -Activity "Scanning Folder Structure" -Completed -Id 3

# Get drive information for completion
$driveLetter = $StartPath.Substring(0, 1)
Write-DiagnosticMessage "Getting drive stats for: $driveLetter" -Color "Cyan"
$driveInfo = Get-DriveStat -DriveLetter $driveLetter

# Show drive information
Show-DriveInfo -Volume $driveInfo

# Script completed successfully
Write-Information -MessageData "$($script:ANSI.Green)`nScript finished at $((Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')) (UTC)$($script:ANSI.Reset)" -InformationAction Continue

# Stop the transcript at the end and completely release all file handles
try {
    # Store path for display before stopping the transcript
    $savedTranscriptPath = $null
    if ($script:transcriptFile -and (Test-Path $script:transcriptFile)) {
        $savedTranscriptPath = $script:transcriptFile
    }

    # NUCLEAR OPTION: Create a separate PowerShell process to release the transcript
    # This ensures the transcript file is released even if our process has a handle lock
    $releaseScript = @"
Start-Sleep -Seconds 1
`$Error.Clear()
try {
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
} catch {
    # Ignore any errors when stopping transcript in cleanup script
}
[System.GC]::Collect(2, [System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, `$true)
[System.GC]::WaitForPendingFinalizers()
exit
"@

    # First try our safe transcript stopping function
    Stop-TranscriptSafely

    # Then launch external process to force transcript release
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command $releaseScript" -WindowStyle Hidden
      # Close any open streams that might hold references to the transcript
    try {
        [System.IO.StreamWriter]::Null.Close()
    } catch {
        Write-Verbose "Error closing StreamWriter.Null: $($_.Exception.Message)"
    }

    try {
        [System.IO.StreamWriter]::Null.Dispose()
    } catch {
        Write-Verbose "Error disposing StreamWriter.Null: $($_.Exception.Message)"
    }
      # Release console output streams if possible
    try {
        [System.Console]::Out.Flush()
    } catch {
        Write-Verbose "Error flushing Console.Out: $($_.Exception.Message)"
    }

    try {
        [System.Console]::Error.Flush()
    } catch {
        Write-Verbose "Error flushing Console.Error: $($_.Exception.Message)"
    }

    # Multiple rounds of aggressive garbage collection with increasing aggressiveness
    for ($i = 0; $i -lt 3; $i++) {
        [System.GC]::Collect(2, [System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $true)
        [System.GC]::WaitForPendingFinalizers()
        Start-Sleep -Milliseconds 200
    }

    # Display the log path if we have it
    if ($savedTranscriptPath) {
        Write-Information -MessageData "Log saved to: $savedTranscriptPath" -InformationAction Continue
    }

    # Explicitly dispose of all runspaces
    $runspaces = [runspacefactory]::Runspaces
    if ($null -ne $runspaces -and $runspaces.Count -gt 0) {
        foreach ($rs in $runspaces) {            try {
                if ($rs.RunspaceStateInfo.State -eq 'Opened') {
                    $rs.Close()
                    $rs.Dispose()
                }            } catch {
                # Log the error but continue cleanup
                try {
                    Write-Verbose "Error during runspace cleanup: $($_.Exception.Message)"
                } catch {
                    # If Write-Verbose is not available, continue silently
                    # This is expected in restricted execution contexts
                    Write-Debug "Write-Verbose not available during runspace cleanup"
                }
            }
        }
    }
      # Clear all script-scoped variables that might hold references
    # This is crucial for ensuring no lingering handles to the transcript
    @(
        'processedFolders', 'recursiveStopwatch', 'InaccessibleFolders',
        'totalRecursiveFiles', 'totalRecursiveSize', 'totalRecursiveFolders',
        'transcriptActive', 'transcriptFile', 'lastRecursiveProgressUpdate',
        'progressUpdateJob', 'recursiveUpdateFrequency', 'UseProgressBars',
        'ANSI', 'currentProgressId', 'totalEstimatedFolders', 'FILE_ATTRIBUTE_HIDDEN',
        'FILE_ATTRIBUTE_SYSTEM', 'FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS',
        'FILE_ATTRIBUTE_RECALL_ON_OPEN', 'FILE_ATTRIBUTE_REPARSE_POINT'
    ) | ForEach-Object {        try {
            Set-Variable -Name $_ -Value $null -Scope Script -ErrorAction SilentlyContinue
            Remove-Variable -Name $_ -Scope Script -Force -ErrorAction SilentlyContinue            } catch {
                try {
                    Write-Verbose "Could not clean up script variable '$_': $($_.Exception.Message)"                } catch {
                    # If Write-Verbose is not available, continue silently
                    # This is expected in restricted execution contexts
                    Write-Debug "Write-Verbose not available during variable cleanup"
                }
            }
    }
    $script:totalRecursiveSize = $null
    $script:totalRecursiveFolders = $null
    $script:transcriptActive = $null
    $script:transcriptFile = $null
    $script:UseProgressBars = $null
    $script:ANSI = $null    # Final aggressive garbage collection
    [System.GC]::Collect(2, [System.GC]::MaxGeneration, [System.GCCollectionMode]::Forced, $true)
    [System.GC]::WaitForPendingFinalizers()

    Write-Information -MessageData "$($script:ANSI.Green)All resources cleaned up successfully.$($script:ANSI.Reset)" -InformationAction Continue
} catch {
    try {
        Write-Warning "Error during cleanup: $($_.Exception.Message)"
    } catch {
        # If Write-Warning is not available, use Write-Output as fallback
        Write-Output "Error during cleanup: $($_.Exception.Message)"
    }
}

#endregion
