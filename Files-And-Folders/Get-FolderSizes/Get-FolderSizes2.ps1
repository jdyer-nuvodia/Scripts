# =============================================================================
# Script: Get-FolderSizes.ps1
# Created: 2025-02-05 00:55:03 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-28 21:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.9
# Additional Info: Fixed ThreadJob path issues in recursive directory scanning
# =============================================================================

<#
.SYNOPSIS
    Ultra-fast directory scanner that analyzes folder sizes and identifies largest files.

.DESCRIPTION
    This script performs a high-performance recursive directory scan to identify the largest
    folders and files in a given directory path. It uses multi-threading when available
    and optimized .NET methods for maximum performance, even when scanning system directories.

    Features:
    - Multi-threaded scanning when ThreadJob module is available
    - Fallback to single-threaded mode when ThreadJob is not available
    - Handles access-denied errors gracefully
    - Identifies largest files in each directory
    - Creates detailed log file of the scan
    - Requires administrative privileges
    - Supports custom depth limitation

    Dependencies:
    - Windows PowerShell 5.1 or later
    - ThreadJob module (optional - will be installed if not present)
    - Administrative privileges
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
    - Administrative access
    - Read access to scanned directories
    - Write access to C:\temp for logging
    
    Validation Requirements:
    - Verify administrative privileges
    - Check available memory (4GB+)
    - Validate write access to log directory
    - Test ThreadJob module availability

    Author:  jdyer-nuvodia
    Created: 2025-02-05 00:55:03 UTC
    Updated: 2025-02-07 15:41:22 UTC

    Requirements:
    - Windows PowerShell 5.1 or later
    - Administrative privileges
    - Minimum 4GB RAM recommended for large directory structures
    - ThreadJob module (optional - will be installed if not present)

    Version History:
    1.0.0 - Initial release
    1.0.1 - Fixed compatibility issues with older PowerShell versions
    1.0.2 - Added ThreadJob module handling and fallback mechanism
#>

#Requires -RunAsAdministrator

param (
    [string]$Path = "C:\",
    [int]$MaxDepth = 10,
    [ValidateRange(1, 50)]
    [int]$Top = 3
)

# Modified Initialize-NuGetProvider to run silently
function Initialize-NuGetProvider {
    try {
        # Check if NuGet provider is installed and meets minimum version
        $nugetProvider = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
        $minimumVersion = [Version]"2.8.5.201"

        if (-not $nugetProvider -or $nugetProvider.Version -lt $minimumVersion) {
            Write-Host "Installing NuGet provider..." -ForegroundColor Cyan
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:$false | Out-Null
            Write-Host "NuGet provider installed successfully." -ForegroundColor Green
        }
        return $true
    }
    catch {
        Write-Error "Failed to install NuGet provider: $($_.Exception.Message)"
        return $false
    }
}

# Modified Initialize-ThreadJobModule to run silently
function Initialize-ThreadJobModule {
    try {
        # Ensure NuGet provider is installed first
        if (-not (Initialize-NuGetProvider)) {
            Write-Warning "Could not initialize NuGet provider. ThreadJob installation may fail."
            return $false
        }

        # First try to import if it exists
        if (Get-Module -ListAvailable -Name ThreadJob) {
            Import-Module ThreadJob -ErrorAction Stop
            return $true
        }

        Write-Host "ThreadJob module not found. Attempting to install..." -ForegroundColor Cyan

        # Set up PSGallery as trusted repository if needed
        if (-not (Get-PSRepository -Name "PSGallery" -ErrorAction SilentlyContinue)) {
            Register-PSRepository -Default -ErrorAction Stop -Force
        }
        if ((Get-PSRepository -Name "PSGallery").InstallationPolicy -ne "Trusted") {
            Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted -ErrorAction Stop -Force
        }

        # Install the module
        Install-Module -Name ThreadJob -Repository PSGallery -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop -Confirm:$false
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

# Check PowerShell version and set compatibility mode
$script:isLegacyPowerShell = $PSVersionTable.PSVersion.Major -lt 5
if ($script:isLegacyPowerShell) {
    Write-Warning "Running in PowerShell 4.0 compatibility mode. Some features may be limited."
    $global:useThreadJobs = $false
} else {
    $global:useThreadJobs = Initialize-ThreadJobModule
}

# Setup transcript logging
$transcriptPath = "C:\temp"
if (-not (Test-Path $transcriptPath)) {
    New-Item -ItemType Directory -Path $transcriptPath -Force | Out-Null
}
$transcriptFile = Join-Path $transcriptPath "FolderScan_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').txt"
Start-Transcript -Path $transcriptFile -Force

# Add script header to transcript
Write-Host "======================================================"
Write-Host "Folder Size Scanner - Execution Log"
Write-Host "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host "Started (UTC): $((Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "User: $env:USERNAME"
Write-Host "Computer: $env:COMPUTERNAME"
Write-Host "Target Path: $Path"
Write-Host "Threading Mode: $(if ($global:useThreadJobs) { 'Multi-threaded' } else { 'Single-threaded' })"
Write-Host "======================================================"
Write-Host ""

# No longer checking/restarting for elevated privileges, assuming already run as admin

# Remove existing type if it exists
Remove-TypeData -TypeName "FastFileScanner" -ErrorAction SilentlyContinue

# Add .NET methods with unique type name
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
                // Use GetFiles with SearchOption.TopDirectoryOnly for better reliability
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
"@

# Add a static helper type for folder processing
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
"@ -ErrorAction Stop

$ErrorActionPreference = 'SilentlyContinue'

Write-Host "Ultra-fast folder analysis starting at: $Path"
Write-Host "Script started by: $env:USERNAME at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

function Format-SizeWithPadding {
    param (
        [double]$Size,
        [int]$DecimalPlaces = 2
    )
    return "{0:F$DecimalPlaces}" -f $Size
}

function Write-ProgressBar {
    param (
        [int]$Current,
        [int]$Total,
        [string]$Status,
        [int]$BarLength = 50
    )
    
    $percentComplete = if ($Total -eq 0) { 100 } else { [math]::Min(100, ($Current / $Total * 100)) }
    $filled = [math]::Round($BarLength * ($percentComplete / 100))
    $unfilled = $BarLength - $filled
    
    $progressBar = "[" + ("=" * $filled) + (" " * $unfilled) + "]"
    $percentage = "{0,3:N0}%" -f $percentComplete
    
    Write-Host "`r$progressBar $percentage | $Status" -NoNewline
    if ($Current -eq $Total) { Write-Host "" }
}

function Format-Bytes {
    param (
        [double]$Bytes
    )
    $kilobyte = 1024
    $megabyte = $kilobyte * 1024
    $gigabyte = $megabyte * 1024
    $terabyte = $gigabyte * 1024

    if ($Bytes -gt $terabyte) {
        return "$(Format-SizeWithPadding -Size ($Bytes / $terabyte)) TB"
    } elseif ($Bytes -gt $gigabyte) {
        return "$(Format-SizeWithPadding -Size ($Bytes / $gigabyte)) GB"
    } elseif ($Bytes -gt $megabyte) {
        return "$(Format-SizeWithPadding -Size ($Bytes / $megabyte)) MB"
    } elseif ($Bytes -gt $kilobyte) {
        return "$(Format-SizeWithPadding -Size ($Bytes / $kilobyte)) KB"
    } else {
        return "$([math]::Round($Bytes)) bytes"
    }
}

function Format-Path {
    param (
        [string]$Path
    )
    try {
        # Convert to full path and normalize separators
        $fullPath = [System.IO.Path]::GetFullPath($Path.Trim())
        return $fullPath
    }
    catch {
        Write-Warning "Error formatting path '$Path': $($_.Exception.Message)"
        return $Path
    }
}

function Escape-StringForInvokeExpression {
    param (
        [string]$InputString
    )
    # Properly escape special characters for PowerShell
    $escaped = $InputString -replace "'", "''"
    $escaped = $escaped -replace "`$", "`$`$"  # Escape $ characters
    $escaped = $escaped -replace "\\", "\\\\"  # Escape backslashes
    return $escaped
}

function Get-FolderSizes {
    param (
        [string]$FolderPath,
        [int]$CurrentDepth = 0
    )

    if ($CurrentDepth -ge $MaxDepth) { return $null }

    try {
        # Get all subdirectories - compatibility mode handling
        $folders = if ($script:isLegacyPowerShell) {
            @([System.IO.Directory]::GetDirectories($FolderPath))
        } else {
            @(Get-ChildItem -Path $FolderPath -Directory -Force -ErrorAction SilentlyContinue)
        }

        # Get largest file in current directory - using SafeInvoke pattern
        $safePathForInvokeExpression = Escape-StringForInvokeExpression -InputString $FolderPath
        $largestFile = & {
            try {
                # Use static helper class instead of Invoke-Expression
                [FolderSizeHelper]::GetLargestFile($FolderPath)
            } catch {
                Write-Warning "Error getting largest file in $FolderPath : $_"
                $null  # Ensure a null value is returned in case of error
            }
        }

        $folderData = foreach ($folder in $folders) {
            try {
                # Use static helper class for getting directory size and counts
                $size = [FolderSizeHelper]::GetDirectorySize($folder)
                $counts = [FolderSizeHelper]::GetDirectoryCounts($folder)
                $fileCount = $counts.Item1
                $folderCount = $counts.Item2

                [PSCustomObject]@{
                    Name        = $folder #$folder.Name
                    Path        = $folder #$folder.FullName
                    Size        = $size
                    FileCount   = $fileCount
                    FolderCount = $folderCount
                }
            }
            catch {
                Write-Warning "Error processing folder '$folder': $($_.Exception.Message)"
                # Continue to next folder
            }
        }

        # Sort folders by size
        $sortedFolders = $folderData | Sort-Object -Property Size -Descending | Select-Object -First $Top

        # Output folder information
        Write-Host "  " * $CurrentDepth "Folder: $($FolderPath)"
        if ($largestFile) {
            Write-Host "  " * ($CurrentDepth + 1) "Largest File: $($largestFile.Name) ($($largestFile.Size) bytes)"
        } else {
            Write-Host "  " * ($CurrentDepth + 1) "No files found."
        }

        foreach ($sf in $sortedFolders) {
            Write-Host "  " * ($CurrentDepth + 1) "- $($sf.Name) ($($sf.Size) bytes)"
        }

        # Recursive call for subfolders
        if ($global:useThreadJobs) {
            # Fixed threaded recursion - directly call function within job instead of dot-sourcing script
            foreach ($sf in $sortedFolders) {
                $formatBytes = ${function:Format-Bytes}
                $formatPath = ${function:Format-Path}
                $escapeString = ${function:Escape-StringForInvokeExpression}
                
                $job = Start-ThreadJob -ScriptBlock {
                    param (
                        $subFolderPath, 
                        $currentDepth, 
                        $maxDepth, 
                        $top,
                        $topLevel
                    )
                    
                    # Set up needed function and variable in job context
                    $global:useThreadJobs = $false  # Prevent nested thread jobs
                    $MaxDepth = $maxDepth
                    $Top = $top
                    
                    # Recursive scan within job (without threading)
                    & $topLevel -FolderPath $subFolderPath -CurrentDepth ($currentDepth + 1)
                    
                } -ArgumentList $sf.Path, $CurrentDepth, $MaxDepth, $Top, ${function:Get-FolderSizes}
                
                Write-Host "  " * ($CurrentDepth + 2) "Started thread job for $($sf.Path)"
            }
        } else {
            # Direct recursion
            foreach ($sf in $sortedFolders) {
                Get-FolderSizes -FolderPath $sf.Path -CurrentDepth ($CurrentDepth + 1)
            }
        }
    }
    catch {
        Write-Warning "Error scanning folder '$FolderPath': $($_.Exception.Message)"
    }
}

# Start the folder size analysis
Get-FolderSizes -FolderPath $Path -CurrentDepth 0

# Wait for thread jobs to complete
if ($global:useThreadJobs) {
    Write-Host "Waiting for all thread jobs to complete..."
    Get-Job | Wait-Job | Receive-Job
    Remove-Job -State Completed
    Write-Host "All thread jobs completed."
}

# Stop transcript logging
Write-Host ""
Write-Host "======================================================"
Write-Host "Folder Scan Completed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host "======================================================"
Stop-Transcript

Write-Host "Script execution completed. Log file: $transcriptFile"
