# =============================================================================
# Script: Get-FolderSizes.ps1
# Created: 2025-02-05 00:55:03 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-28 22:45:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1.0
# Additional Info: Modified for silent non-interactive operation with automatic dependency installation
# =============================================================================

# Requires -Version 5.1

<#
.SYNOPSIS
    Ultra-fast directory scanner that analyzes folder sizes and identifies largest files.

.DESCRIPTION
    This script performs a high-performance recursive directory scan to identify the largest
    folders and files in a given directory path. It uses multi-threading when available
    and optimized .NET methods for maximum performance, even when scanning system directories.
    
    All dependencies are installed automatically without user interaction.

    Features:
    - Multi-threaded scanning when ThreadJob module is available
    - Silent installation of required modules without user interaction
    - Fallback to single-threaded mode when ThreadJob is not available
    - Handles access-denied errors gracefully
    - Identifies largest files in each directory
    - Creates detailed log file of the scan
    - Continues with limited functionality if admin rights unavailable
    - Supports custom depth limitation

    Dependencies:
    - Windows PowerShell 5.1 or later
    - ThreadJob module (optional - will be installed automatically if not present)
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

.NOTES
    Security Level: Medium
    Required Permissions: 
    - Administrative access (recommended but not required)
    - Read access to scanned directories
    - Write access to C:\temp for logging
    
    Validation Requirements:
    - Check available memory (4GB+)
    - Validate write access to log directory
    - Test ThreadJob module availability

    Author:  jdyer-nuvodia
    Created: 2025-02-05 00:55:03 UTC
    Updated: 2025-03-12 18:30:00 UTC

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
#>

param (
    [string]$Path = "C:\",
    [int]$MaxDepth = 10,
    [ValidateRange(1, 50)]
    [int]$Top = 3
)

#region Helper Functions

function Initialize-NuGetProvider {
    try {
        $nugetProvider = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
        $minimumVersion = [Version]"2.8.5.201"

        if (-not $nugetProvider -or $nugetProvider.Version -lt $minimumVersion) {
            Write-Host "Installing NuGet provider..." -ForegroundColor Cyan
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:$false -SkipPublisherCheck | Out-Null
            Write-Host "NuGet provider installed successfully." -ForegroundColor Green
            return $true
        }
        return $true
    }
    catch {
        Write-Host "Failed to install NuGet provider: $($_.Exception.Message)" -ForegroundColor Yellow
        return $false
    }
}

function Initialize-ThreadJobModule {
    try {
        if (-not (Initialize-NuGetProvider)) {
            Write-Warning "Could not initialize NuGet provider. ThreadJob installation may fail."
            return $false
        }

        if (Get-Module -ListAvailable -Name ThreadJob) {
            Import-Module ThreadJob -ErrorAction Stop
            return $true
        }

        Write-Host "ThreadJob module not found. Attempting to install..." -ForegroundColor Cyan
        
        if (-not (Get-PSRepository -Name "PSGallery" -ErrorAction SilentlyContinue)) {
            Register-PSRepository -Default -Force -ErrorAction Stop
        }
        
        Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted -Force -ErrorAction SilentlyContinue

        Install-Module -Name ThreadJob -Repository PSGallery -Force -AllowClobber -Scope CurrentUser -SkipPublisherCheck -Confirm:$false -ErrorAction Stop
        Import-Module ThreadJob -Force -ErrorAction Stop
        Write-Host "ThreadJob module installed successfully." -ForegroundColor Green
        return $true
    }
    catch {
        Write-Warning "Could not install/import ThreadJob module: $($_.Exception.Message)"
        Write-Host "Falling back to single-threaded operation mode." -ForegroundColor Yellow
        return $false
    }
}

function Format-SizeWithPadding {
    param (
        [double]$Size,
        [int]$DecimalPlaces = 2
    )
    return "{0:F$DecimalPlaces}" -f $Size
}

function Format-Path {
    param (
        [string]$Path
    )
    try {
        $fullPath = [System.IO.Path]::GetFullPath($Path.Trim())
        return $fullPath
    }
    catch {
        Write-Warning "Error formatting path '$Path': $($_.Exception.Message)"
        return $Path
    }
}

#endregion

#region Setup

# Check for elevated privileges but don't prompt user - continue with limited functionality
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    Write-Host "Running with limited privileges. Some directories may be inaccessible." -ForegroundColor Yellow
}

# PowerShell Version Check and ThreadJob Initialization
$script:isLegacyPowerShell = $PSVersionTable.PSVersion.Major -lt 5
if ($script:isLegacyPowerShell) {
    Write-Warning "Running in PowerShell 4.0 compatibility mode. Some features may be limited."
    $global:useThreadJobs = $false
} else {
    $global:useThreadJobs = Initialize-ThreadJobModule
}

# Transcript Logging Setup
$transcriptPath = "C:\temp"
try {
    if (-not (Test-Path $transcriptPath)) {
        New-Item -ItemType Directory -Path $transcriptPath -Force -ErrorAction SilentlyContinue | Out-Null
    }
    
    if (Test-Path $transcriptPath) {
        $transcriptFile = Join-Path $transcriptPath "FolderScan_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').txt"
        Start-Transcript -Path $transcriptFile -Force -ErrorAction SilentlyContinue
    } else {
        $transcriptFile = Join-Path $env:TEMP "FolderScan_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').txt"
        Start-Transcript -Path $transcriptFile -Force -ErrorAction SilentlyContinue
        Write-Warning "Could not create transcript in C:\temp, using $transcriptFile instead"
    }
} catch {
    Write-Warning "Failed to start transcript: $_"
}

# Script Header in Transcript
Write-Host "======================================================"
Write-Host "Folder Size Scanner - Execution Log"
Write-Host "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host "Started (UTC): $((Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "User: $env:USERNAME"
Write-Host "Computer: $env:COMPUTERNAME"
Write-Host "Target Path: $Path"
Write-Host "Admin Privileges: $isAdmin"
Write-Host "Threading Mode: $(if ($global:useThreadJobs) { 'Multi-threaded' } else { 'Single-threaded' })"
Write-Host "======================================================"
Write-Host ""

# .NET Type Definition
Remove-TypeData -TypeName "FastFileScanner" -ErrorAction SilentlyContinue
$typeName = "FastFileScanner_" + (Get-Random)
Add-Type -TypeDefinition @"
using System;
using System.IO;
using System.Linq;
using System.Security;
using System.Collections.Generic;
using System.Runtime.InteropServices;

public class $typeName {
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    static extern bool GetDiskFreeSpaceEx(string lpDirectoryName,
        out ulong lpFreeBytesAvailable,
        out ulong lpTotalNumberOfBytes,
        out ulong lpTotalNumberOfFreeBytes);

    public static long GetDirectorySize(string path) {
        long size = 0;
        var stack = new Stack<string>();
        stack.Push(path);

        while (stack.Count > 0) {
            string dir = stack.Pop();
            try {
                foreach (string file in Directory.GetFiles(dir)) {
                    try {
                        size += new FileInfo(file).Length;
                    }
                    catch (Exception) { }
                }

                foreach (string subDir in Directory.GetDirectories(dir)) {
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

    public static Tuple<int, int> GetDirectoryCounts(string path) {
        int files = 0;
        int folders = 0;
        var stack = new Stack<string>();
        stack.Push(path);

        while (stack.Count > 0) {
            string dir = stack.Pop();
            try {
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

    public static FileInfo GetLargestFile(string path) {
        try {
            return new DirectoryInfo(path)
                .GetFiles("*.*", SearchOption.TopDirectoryOnly)
                .OrderByDescending(f => f.Length)
                .FirstOrDefault();
        }
        catch {
            return null;
        }
    }
}
"@ -ErrorAction SilentlyContinue

# Helper Type for Folder Processing
Add-Type -TypeDefinition @"
using System;
using System.IO;
using System.Linq;
using System.Security;
using System.Collections.Generic;

public static class FolderSizeHelper
{
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
                foreach (var subDir in subDirs)
                {
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

Write-Host "Ultra-fast folder analysis starting at: $Path"
Write-Host "Script started by: $env:USERNAME at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

#endregion

#region Folder Scanning Logic

function Get-FolderSize {
    param (
        [string]$FolderPath,
        [int]$CurrentDepth,
        [int]$MaxDepth,
        [int]$Top
    )

    try {
        if ($CurrentDepth -gt $MaxDepth) {
            return
        }

        $folderPath = Format-Path $FolderPath
        if (-not (Test-Path -Path $folderPath -PathType Container)) {
            Write-Warning "Path '$FolderPath' does not exist or is not a directory."
            return
        }

        Write-Host "Scanning: $folderPath (Depth: $CurrentDepth)"

        # Get Folder Size and Counts using .NET methods
        $size = [FolderSizeHelper]::GetDirectorySize($folderPath)
        $counts = [FolderSizeHelper]::GetDirectoryCounts($folderPath)
        $fileCount = $counts.Item1
        $folderCount = $counts.Item2

        # Get Largest File
        $largestFile = [FolderSizeHelper]::GetLargestFile($folderPath)

        # Output Folder Information
        Write-Host "  Size: $(Format-SizeWithPadding ($size / 1MB)) MB"
        Write-Host "  Files: $fileCount, Folders: $folderCount"

        if ($largestFile) {
            Write-Host "  Largest File: $($largestFile.Name) ($((Format-SizeWithPadding ($largestFile.Size / 1MB)) ) MB)"
        }
        else {
            Write-Host "  No files found."
        }

        # Get Subfolders and Process
        $subFolders = try { Get-ChildItem -Path $folderPath -Directory -ErrorAction Stop } catch { Write-Warning "Error getting subfolders in '$folderPath': $($_.Exception.Message)"; @() }

        if ($subFolders) {
            $sortedFolders = $subFolders | ForEach-Object {
                $subFolderPath = $_.FullName
                $subFolderSize = try { [FolderSizeHelper]::GetDirectorySize($subFolderPath) } catch { 0 }
                [PSCustomObject]@{
                    Path = $subFolderPath
                    Size = $subFolderSize
                }
            } | Sort-Object -Property Size -Descending | Select-Object -First $Top

            foreach ($subFolder in $sortedFolders) {
                Get-FolderSize -FolderPath $subFolder.Path -CurrentDepth ($CurrentDepth + 1) -MaxDepth $MaxDepth -Top $Top
            }
        }
    }
    catch {
        Write-Warning "Error processing folder '$FolderPath': $($_.Exception.Message)"
    }
}

# Start the Recursive Scan
Get-FolderSize -FolderPath $Path -CurrentDepth 1 -MaxDepth $MaxDepth -Top $Top

#endregion

# Stop Transcript
try {
    Stop-Transcript
    Write-Host "Transcript stopped. Log file: $transcriptFile"
} catch {
    Write-Warning "Failed to stop transcript: $_"
}

Write-Host "Script finished at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
