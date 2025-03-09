# =============================================================================
# Script: Get-FolderSizes.ps1
# Created: 2/5/2025 00:55:03 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-09 16:46:00 UTC
# Updated By: jdyer-nuvodia
# Version: 2.1.10
# Additional Info: Fixed thread progress messages not appearing in transcript log
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
    - Creates detailed log file of the scan
    - Continues with limited functionality if admin rights unavailable
    - Supports custom depth limitation
    - Properly handles symbolic links and junction points
    - Includes hidden and system folders like "All Users"

    Dependencies:
    - Windows PowerShell 5.1 or later
    - Administrative privileges recommended but not required
    - Minimum 4GB RAM recommended

    Performance Impact:
    - CPU: Medium to High during scan
    - Memory: Medium (4GB+ recommended)
    - Disk I/O: Low to Medium
    - Network: Low (unless scanning network paths)

.PARAMETER Path
    The root directory path to start scanning from. Defaults to "C:\"

.PARAMETER MaxDepth
    Maximum depth of recursion for the directory scan. Defaults to 10 levels deep.

.PARAMETER Top
    Number of largest folders to display at each level. Defaults to 3. Range: 1-50.
    
.PARAMETER IncludeHiddenSystem
    Include hidden and system folders in the scan. Defaults to $true.

.PARAMETER FollowJunctions
    Follow junction points and symbolic links when calculating sizes. Defaults to $true.

.PARAMETER MaxThreads
    Maximum number of parallel threads to use for processing folders. Defaults to 10.
    Higher values may improve performance on systems with many CPU cores but will use more memory.

.EXAMPLE
    .\Get-FolderSizes.ps1
    Scans the C:\ drive with default settings

.EXAMPLE
    .\Get-FolderSizes.ps1 -Path "D:\Users" -MaxDepth 5
    Scans the D:\Users directory with a maximum depth of 5 levels

.EXAMPLE
    .\Get-FolderSizes.ps1 -Path "\\server\share"
    Scans a network share starting from the root

.EXAMPLE
    .\Get-FolderSizes.ps1 -Top 10
    Scans the C:\ drive and shows the 10 largest folders at each level
    
.EXAMPLE
    .\Get-FolderSizes.ps1 -IncludeHiddenSystem $false
    Scans the C:\ drive but excludes hidden and system folders

.EXAMPLE
    .\Get-FolderSizes.ps1 -Path "D:\Data" -MaxThreads 20
    Scans the D:\Data directory using 20 parallel threads for faster processing on multi-core systems.

.NOTES
    Security Level: Medium
    Required Permissions: 
    - Administrative access (recommended but not required)
    - Read access to scanned directories
    - Write access to C:\temp for logging
    
    Validation Requirements:
    - Check available memory (4GB+)
    - Validate write access to log directory

    Author:  jdyer-nuvodia
    Created: 2025-02-05 00:55:03 UTC
    Updated: 2025-06-04 17:25:00 UTC

    Requirements:
    - Windows PowerShell 5.1 or later
    - Administrative privileges recommended
    - Minimum 4GB RAM recommended for large directory structures

    Version History:
    1.0.0 - Initial release
    1.0.1 - Fixed compatibility issues with older PowerShell versions
    1.0.2 - Added ThreadJob module handling and fallback mechanism
    1.0.8 - Fixed handling of special characters in ThreadJobs processing
    1.1.0 - Modified for silent non-interactive operation with automatic dependency installation
    1.2.0 - Updated output formatting to display results in tabular format with progress indicators
    1.4.0 - Modified to only descend into the largest folder at each directory level
    1.5.0 - Added proper support for symbolic links and junction points
    1.5.1 - Fixed 'findstr' command not found errors by using PowerShell native commands
    1.5.2 - Added special handling for OneDrive reparse points
    1.5.3 - Fixed redundant completion messages in recursive processing
    1.5.4 - Eliminated redundant completion messages in recursive processing
    1.5.5 - Completely redesigned recursive processing to prevent redundant messages
    1.5.6 - Fixed Script Analyzer warnings for unused variables
    1.5.7 - Fixed recursive processing of completion messages with completion state tracking
    1.5.8 - Suppressed return value output in console
    1.6.0 - Added support for hidden and system folders like "All Users"
    1.6.1 - Suppressed mountpoint and junction output messages
    1.6.2 - Fixed catch block structure for proper exception handling
    1.6.3 - Added pre-emptive NuGet provider installation to prevent prompts
    1.6.4 - Fixed invalid assignment expressions for preference variables
    1.6.5 - Fixed parameter syntax error with path value
    1.6.6 - Fixed parameter syntax by removing trailing comma in path value
    1.6.7 - Eliminated GUI window flash during NuGet provider installation
    1.6.8 - Fixed variable name conflicts causing incorrect path targeting
    1.6.9 - Eliminated PowerShell window by using background jobs instead of Process
    1.7.0 - Standardized console output colors to match organizational standards
    1.7.1 - Enhanced silent NuGet provider installation to prevent prompts
    1.7.2 - Attempted fix for remaining NuGet silent install prompts
    1.7.3 - Moved transcript logging prior to NuGet provider installation
    1.7.4 - Added Initialize-ThreadJobModule function to avoid reference errors
    1.7.5 - Moved Initialize-ThreadJobModule function above usage
    1.7.6 - Moved Initialize-ThreadJobModule function to top of script
    1.7.7 - Changed log file location to use script directory instead of C:\temp
    1.7.8 - Added verbose diagnostic logging for NuGet provider installation
    1.7.9 - Fixed unsupported -Scope parameter in Set-PSRepository command
    1.8.0 - Fixed duplicate transcript initialization causing file access errors
    1.8.1 - Fixed UTC timestamp formatting in completion message
    1.8.2 - Implemented foolproof NuGet provider silent installation
    1.9.0 - Replaced ThreadJob with runspace pools for better performance
    1.9.1 - Fixed syntax error in comment escaping
    1.9.2 - Fixed PSGallery repository name quoting in Set-PSRepository command
    1.9.3 - Fixed string formatting in transcript path creation
    1.9.4 - Fixed string formatting in date format variable
    1.9.5 - Fixed string formatting in Get-PathType error handling
    1.9.6 - Fixed string formatting escape sequence in Get-PathType error handling
    1.9.7 - Fixed string formatting using double quotes to prevent parser error
    1.9.8 - Fixed parser error in Get-PathType using string concatenation
    1.9.9 - Fixed parser error in Get-PathType using string concatenation
    1.9.10 - Fixed syntax errors and parser issues in string handling
    2.0.0 - Removed ThreadJob and NuGet dependencies for simpler execution
    2.1.0 - Added multi-threading using runspaces for improved performance
    2.1.1 - Added MaxThreads parameter documentation and examples
    2.1.2 - Added parallel execution diagnostics and monitoring
    2.1.3 - Removed redundant transcript stopped message
    2.1.4 - Fixed missing catch block and closing brace syntax errors
    2.1.5 - Suppressed processing progress messages from console output
    2.1.6 - Fixed thread completion messages appearing in console output
    2.1.7 - Fixed thread messages still appearing in console output
    2.1.8 - Fixed remaining thread completion messages in console output
    2.1.9 - Fixed thread progress messages appearing in console during processing
    2.1.10 - Fixed thread progress messages not appearing in transcript log
#>

param (
    [string]$Path = 'C:',
    [int]$MaxDepth = 10,
    [ValidateRange(1, 50)]
    [int]$Top = 3,
    [bool]$IncludeHiddenSystem = $true,
    [bool]$FollowJunctions = $true,
    [int]$MaxThreads = 10
)

# Console colors for diagnostic output
function Write-DiagnosticMessage {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [string]$Color = "White"
    )
    
    # Always display diagnostic messages regardless of preference variables
    $timeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    if ($Color -eq "Error") {
        # Using Write-Error would respect $ErrorActionPreference, so we'll use Write-Host with Red color
        Write-Host "[$timeStamp] ERROR: $Message" -ForegroundColor Red
    } else {
        Write-Host "[$timeStamp] $Message" -ForegroundColor $Color
    }
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

# Continue with the rest of the script...

#region Helper Functions

# Function to initialize color scheme for console output
function Show-ColorLegend {
    Write-Host "`n===== Console Output Color Legend =====" -ForegroundColor White
    Write-Host "White     - Standard information" -ForegroundColor White
    Write-Host "Cyan      - Process updates and status" -ForegroundColor Cyan
    Write-Host "Green     - Successful operations and results" -ForegroundColor Green
    Write-Host "Yellow    - Warnings and attention needed" -ForegroundColor Yellow
    Write-Host "Red       - Errors and critical issues" -ForegroundColor Red
    Write-Host "Magenta   - Debug information" -ForegroundColor Magenta
    Write-Host "DarkGray  - Technical details" -ForegroundColor DarkGray
    Write-Host "======================================`n" -ForegroundColor White
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
} # Add missing closing brace

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
    
    Write-Host ("-" * $Width)
    Write-Host ("Folder Path".PadRight(50) + " | " + 
                "Size (GB)".PadLeft(11) + " | " + 
                "Subfolders".PadLeft(15) + " | " + 
                "Files".PadLeft(12) + " | " + 
                "Largest File (in this directory)")
    Write-Host ("-" * $Width)
}

function Write-TableRow {
    param(
        [string]$FolderPath,
        [long]$Size,
        [int]$SubfolderCount,
        [int]$FileCount,
        [object]$LargestFile
    )
    
    $sizeGB = Format-SizeWithPadding -Size $Size -DecimalPlaces 2 -Unit "GB"
    $largestFileInfo = if ($LargestFile) {
        $largestFileSize = Format-SizeWithPadding -Size $LargestFile.Size -DecimalPlaces 2 -Unit "MB"
        "$($LargestFile.Name) ($largestFileSize MB)"
    } else {
        "No files found"
    }
    
    Write-Host ($FolderPath.PadRight(50) + " | " + 
                $sizeGB.PadLeft(11) + " | " + 
                $SubfolderCount.ToString().PadLeft(15) + " | " + 
                $FileCount.ToString().PadLeft(12) + " | " + 
                $largestFileInfo)
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
    
    Write-Host "`r$bar" -NoNewline
    if ($Completed -eq $Total) {
        Write-Host ""  # Add a newline when complete
    }
}

# Add this helper function after the other helper functions
function Write-TranscriptLog {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [string]$Category = "Progress"
    )
    
    # Write to transcript but not console
    Write-Information -MessageData $Message -Tags $Category 6>&1 | Out-File -FilePath $transcriptFile -Append
}

#endregion

#region Setup

# Check for elevated privileges but do not prompt user - continue with limited functionality
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    Write-Host "Running with limited privileges. Some directories may be inaccessible." -ForegroundColor Yellow
}

# Script Header in Transcript
Write-Host "======================================================" -ForegroundColor White
Write-Host "Folder Size Scanner - Execution Log" -ForegroundColor White
Write-Host "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor White
Write-Host "Started (UTC): $((Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor White
Write-Host "User: $env:USERNAME" -ForegroundColor White
Write-Host "Computer: $env:COMPUTERNAME" -ForegroundColor White
Write-Host "Target Path: $Path" -ForegroundColor White
Write-Host "Admin Privileges: $isAdmin" -ForegroundColor White
Write-Host "======================================================" -ForegroundColor White
Write-Host ""

# Show color legend for user reference
Show-ColorLegend

# .NET Type Definition
Remove-TypeData -TypeName "FastFileScanner" -ErrorAction SilentlyContinue
Remove-TypeData -TypeName "FolderSizeHelper" -ErrorAction SilentlyContinue

# Helper Type for Folder Processing
Add-Type -TypeDefinition @"
using System;
using System.IO;
using System.Linq;
using System.Security;
using System.Collections.Generic;
using System.Runtime.InteropServices;

public static class FolderSizeHelper
{
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    static extern bool GetDiskFreeSpaceEx(string lpDirectoryName,
        out ulong lpFreeBytesAvailable,
        out ulong lpTotalNumberOfBytes,
        out ulong lpTotalNumberOfFreeBytes);
        
    public static long GetDirectorySize(string path)
    {
        long size = 0;
        var stack = new Stack<string>();
        stack.Push(path);

        while (stack.Count > 0)
        {
            string dir = stack.Pop();
            try
            {
                foreach (string file in Directory.GetFiles(dir))
                {
                    try
                    {
                        size += new FileInfo(file).Length;
                    }
                    catch (Exception) { }
                }

                foreach (string subDir in Directory.GetDirectories(dir))
                {
                    stack.Push(subDir);
                }
            }
            catch (UnauthorizedAccessException) { }
            catch (SecurityException) { }
            catch (IOException) { }
            catch (Exception) { }
        }
        return size;
    }

    public static Tuple<int, int> GetDirectoryCounts(string path)
    {
        int files = 0;
        int folders = 0;
        var stack = new Stack<string>();
        stack.Push(path);

        while (stack.Count > 0)
        {
            string dir = stack.Pop();
            try
            {
                files += Directory.GetFiles(dir).Length;
                var subDirs = Directory.GetDirectories(dir);
                folders += subDirs.Length;
                foreach (var subDir in subDirs) {
                    stack.Push(subDir);
                }
            }
            catch (UnauthorizedAccessException) { }
            catch (SecurityException) { }
            catch (IOException) { }
            catch (Exception) { }
        }
        return new Tuple<int, int>(files, folders);
    }

    public static FileDetails GetLargestFile(string path)
    {
        try
        {
            var fileInfo = new DirectoryInfo(path)
                .GetFiles("*.*", SearchOption.TopDirectoryOnly)
                .OrderByDescending(f => f.Length)
                .FirstOrDefault();
                
            if (fileInfo == null)
                return null;
                
            return new FileDetails
            {
                Name = fileInfo.Name,
                Path = fileInfo.FullName,
                Size = fileInfo.Length
            };
        }
        catch
        {
            return null;
        }
    }
    
    public class FileDetails
    {
        public string Name { get; set; }
        public string Path { get; set; }
        public long Size { get; set; }
    }
}
"@ -ErrorAction SilentlyContinue

$ErrorActionPreference = 'SilentlyContinue'

Write-Host "Ultra-fast folder analysis starting at: $Path" -ForegroundColor Cyan
Write-Host "Script started by: $env:USERNAME at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor White

#endregion

#region Folder Scanning Logic

# New function to process folders in parallel using runspaces
function Start-FolderProcessing {
    param(
        [array]$Folders,
        [int]$MaxThreads
    )
    
    $RunspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads)
    $RunspacePool.Open()
    $FolderSizeMap = @{}
    $Runspaces = @()
    $activeRunspaces = 0
    $processedCount = 0
    $totalFolders = $Folders.Count
    
    Write-Host "`nParallel Processing Configuration:" -ForegroundColor Cyan
    Write-Host "Maximum Threads: $MaxThreads" -ForegroundColor DarkGray
    Write-Host "Total Folders to Process: $totalFolders" -ForegroundColor DarkGray
    
    foreach ($folder in $Folders) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $RunspacePool
        $activeRunspaces++

        Write-Host "`rActive Runspaces: $activeRunspaces/$MaxThreads" -NoNewline -ForegroundColor Magenta

        [void]$ps.AddScript({
            param($FolderPath)
            
            $threadId = [System.Threading.Thread]::CurrentThread.ManagedThreadId
            # Return thread info instead of writing it
            try {
                $counts = [FolderSizeHelper]::GetDirectoryCounts($FolderPath)
                $size = [FolderSizeHelper]::GetDirectorySize($FolderPath)
                $largestFile = [FolderSizeHelper]::GetLargestFile($FolderPath)
                
                return @{
                    Success = $true
                    FolderPath = $FolderPath
                    Size = $size
                    FileCount = $counts.Item1
                    FolderCount = $counts.Item2
                    LargestFile = $largestFile
                    ThreadId = $threadId
                    Message = "Thread $threadId processing: $FolderPath"
                }
            }
            catch {
                return @{
                    Success = $false
                    FolderPath = $FolderPath
                    Error = $_.Exception.Message
                    ThreadId = $threadId
                    Message = "Thread $threadId failed: $FolderPath"
                }
            }
        }).AddArgument($folder.FullName)
        
        $Runspaces += [PSCustomObject]@{
            Instance = $ps
            Handle = $ps.BeginInvoke()
            Folder = $folder.FullName
            StartTime = [DateTime]::Now
        }
    }
    
    Write-Host "`n`nProcessing Results:" -ForegroundColor Cyan
    
    # Initialize Information stream
    $InformationPreference = 'Continue'
    
    foreach ($r in $Runspaces) {
        try {
            $processedCount++
            $percentComplete = [math]::Round(($processedCount / $totalFolders) * 100, 1)
            
            # Progress info to transcript only
            Write-TranscriptLog -Message "`rProgress: $processedCount/$totalFolders ($percentComplete%)" -Category "Progress"
            
            $result = $r.Instance.EndInvoke($r.Handle)
            $processingTime = ([DateTime]::Now - $r.StartTime).TotalSeconds
            
            if ($result.Success) {
                $FolderSizeMap[$result.FolderPath] = @{ 
                    Size = $result.Size
                    FileCount = $result.FileCount
                    FolderCount = $result.FolderCount
                    LargestFile = $result.LargestFile
                }
                # Completion info to transcript only
                Write-TranscriptLog -Message "$($result.Message) completed in $($processingTime.ToString('0.00'))s" -Category "ThreadComplete"
            }
            else {
                # Keep error messages visible in console
                Write-Host "`nThread $($result.ThreadId) failed: $($r.Folder) - $($result.Error)" -ForegroundColor Red
            }
            $activeRunspaces--
        }
        catch {
            # Keep error messages visible in console
            Write-Host "`nCritical error in runspace for folder $($r.Folder): $($_.Exception.Message)" -ForegroundColor Red
            $activeRunspaces--
        }
        finally {
            $r.Instance.Dispose()
        }
    }
    
    Write-Host "`n`nParallel Processing Summary:" -ForegroundColor Cyan
    Write-Host "Total Folders Processed: $processedCount" -ForegroundColor DarkGray
    Write-Host "Maximum Concurrent Threads: $MaxThreads" -ForegroundColor DarkGray
    
    $RunspacePool.Close()
    $RunspacePool.Dispose()
    return $FolderSizeMap
}

# Modify the Get-FolderSize function to use parallel processing
function Get-FolderSize {
    param (
        [string]$FolderPath,
        [int]$CurrentDepth,
        [int]$MaxDepth,
        [int]$Top
    )

    try {
        if ($CurrentDepth -gt $MaxDepth) {
            return @{ 
                ProcessedFolders = $false; 
                HasSubfolders = $false; 
                CompletionMessageShown = $false
            }
        }

        $folderPath = Format-Path $FolderPath
        if (-not (Test-Path -Path $folderPath -PathType Container)) {
            Write-Warning "Path '$FolderPath' does not exist or is not a directory."
            return @{ 
                ProcessedFolders = $false; 
                HasSubfolders = $false; 
                CompletionMessageShown = $false
            }
        }

        # Check if this path is a symbolic link, junction, or mount point
        $pathType = Get-PathType -InputPath $folderPath
        
        # Silently handle special paths - no console output for junction detection
        if ($pathType.Type -ne "Directory" -and $pathType.Type -ne "Unknown") {
            # If it is a link and we are configured to follow links, try to use the target path instead
            if ($FollowJunctions -and $pathType.Target -and $pathType.Target -ne "Unknown Target" -and $pathType.Target -ne "Cloud Storage") {
                # Handle relative paths in targets
                if (-not [System.IO.Path]::IsPathRooted($pathType.Target)) {
                    $targetPath = Join-Path (Split-Path $folderPath -Parent) $pathType.Target
                } else {
                    $targetPath = $pathType.Target
                }
                
                # Check if the target exists
                if (Test-Path -Path $targetPath -PathType Container) {
                    $folderPath = $targetPath
                }
            }
        }

        Write-Host "`nTop $Top Largest Folders in: $folderPath" -ForegroundColor Cyan
        Write-Host ""

        # Get Folder Size and Counts using .NET methods
        $counts = [FolderSizeHelper]::GetDirectoryCounts($folderPath)
        $folderCount = $counts.Item2

        # Get Largest File
        $largestFile = [FolderSizeHelper]::GetLargestFile($folderPath)

        # Display largest file information
        if ($largestFile) {
            Write-Host "`nLargest file in $folderPath :" -ForegroundColor White
            Write-Host "Name: $($largestFile.Name)" -ForegroundColor DarkGray
            $fileSize = if ($largestFile.Size -gt 1MB) {
                "$([Math]::Round($largestFile.Size / 1MB, 2)) MB"
            } elseif ($largestFile.Size -gt 1KB) {
                "$([Math]::Round($largestFile.Size / 1KB, 2)) KB"
            } else {
                "$($largestFile.Size) bytes"
            }
            Write-Host "Size: $fileSize" -ForegroundColor DarkGray
        }

        # Get Subfolders and Process - include hidden and system folders if specified
        $subFolders = try { 
            # Include hidden and system folders if specified
            if ($IncludeHiddenSystem) {
                Get-ChildItem -Path $folderPath -Directory -Force -ErrorAction Stop
            }
            else {
                Get-ChildItem -Path $folderPath -Directory -ErrorAction Stop
            }
        } catch { 
            Write-Warning "Error getting subfolders in '$folderPath': $($_.Exception.Message)"
            @() 
        }

        if ($subFolders -and $subFolders.Count -gt 0) {
            $folderCount = $subFolders.Count
            Write-Host "`nFound $folderCount subfolders to process..." -ForegroundColor Cyan
            
            # Process all folders in parallel using runspace pools
            $folderResults = Start-FolderProcessing -Folders $subFolders -MaxThreads $MaxThreads
            
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
            
            # Always display the table header
            Write-TableHeader
            
            # Get top folders but ensure we do not exceed available folders
            $topFoldersCount = [Math]::Min($Top, $sortedFolders.Count)
            $topFolders = $sortedFolders | Select-Object -First $topFoldersCount
            
            # Display each folder in table format
            foreach ($folder in $topFolders) {
                Write-TableRow -FolderPath $folder.Path -Size $folder.Size -SubfolderCount $folder.FolderCount -FileCount $folder.FileCount -LargestFile $folder.LargestFile
            }
            
            Write-Host ("-" * 150) -ForegroundColor DarkGray
            Write-Host ""
            
            # Process only the largest subfolder if within depth limit
            $completionMessageShown = $false
            if ($CurrentDepth + 1 -le $MaxDepth -and $sortedFolders.Count -gt 0) {
                $largestFolder = $sortedFolders[0] # Get the single largest folder
                
                Write-Host "`nDescending into largest subfolder: $($largestFolder.Path)" -ForegroundColor Cyan
                
                # Call recursively and capture the structured return value
                $result = Get-FolderSize -FolderPath $largestFolder.Path -CurrentDepth ($CurrentDepth + 1) -MaxDepth $MaxDepth -Top $Top
                
                # Only show a completion message if:
                # 1. Processing happened
                # 2. The child had subfolders
                # 3. No completion message has been shown in this branch yet
                if ($result.ProcessedFolders -eq $true -and 
                    $result.HasSubfolders -eq $true -and 
                    $result.CompletionMessageShown -eq $false) {
                    Write-Host "`nCompleted processing the largest subfolder." -ForegroundColor Green
                    $completionMessageShown = $true
                } else {
                    # Propagate the completion message state from child to parent
                    $completionMessageShown = $result.CompletionMessageShown
                }
            }
            
            # Return structured information about this level of processing
            return @{ 
                ProcessedFolders = $true;           # This level processed folders
                HasSubfolders = $true;              # This level had subfolders
                CompletionMessageShown = $completionMessageShown  # Track if any completion message was shown
            }
        } else {
            Write-Host "No subfolders found to process." -ForegroundColor Yellow
            # Return structured information - processed but had no subfolders
            return @{ 
                ProcessedFolders = $true;    # We did process this folder 
                HasSubfolders = $false;      # But it had no subfolders
                CompletionMessageShown = $false  # No completion message needed for leaf nodes
            }
        }
    }
    catch {
        Write-Warning "Error processing folder '$FolderPath': $($_.Exception.Message)"
        return @{ 
            ProcessedFolders = $false; 
            HasSubfolders = $false; 
            CompletionMessageShown = $false
        }
    }
}

# Start the Recursive Scan
Get-FolderSize -FolderPath $Path -CurrentDepth 1 -MaxDepth $MaxDepth -Top $Top | Out-Null

#endregion

#region Drive Information Display
function Show-DriveInfo {
    param (
        [Parameter(Mandatory=$true)]
        [object]$Volume
    )
    
    Write-Host "`nDrive Volume Details:" -ForegroundColor Green
    Write-Host "------------------------" -ForegroundColor Green
    Write-Host "Drive Letter: $($Volume.DriveLetter)" -ForegroundColor White
    Write-Host "Drive Label: $($Volume.FileSystemLabel)" -ForegroundColor White
    Write-Host "File System: $($Volume.FileSystem)" -ForegroundColor White
    Write-Host "Drive Type: $($Volume.DriveType)" -ForegroundColor White
    
    # Format size with appropriate colors based on values
    $totalSize = [math]::Round($Volume.Size/1GB, 2)
    $freeSpace = [math]::Round($Volume.SizeRemaining/1GB, 2)
    $freePercent = [math]::Round(($Volume.SizeRemaining / $Volume.Size) * 100, 1)
    
    Write-Host "Size: $totalSize GB" -ForegroundColor White
    Write-Host "Free Space: $freeSpace GB ($freePercent%)" -ForegroundColor $(if ($freePercent -lt 10) { "Red" } elseif ($freePercent -lt 20) { "Yellow" } else { "Green" })
    Write-Host "Health Status: $($Volume.HealthStatus)" -ForegroundColor White
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
catch { # Add missing catch block
    Write-Error "Error accessing drive information. Error: $_"
}
#endregion

# Stop Transcript
try {
    Stop-Transcript
} catch {
    Write-Warning "Failed to stop transcript: $_"
}

# Display single completion message with properly formatted UTC timestamp
Write-Host "`nScript finished at $((Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')) (UTC)" -ForegroundColor Green
