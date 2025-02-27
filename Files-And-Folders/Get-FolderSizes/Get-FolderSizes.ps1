# =============================================================================
# Script: Get-FolderSizes.ps1
# Created: 2025-02-05 00:55:03 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-27 20:34:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.5
# Additional Info: Enhanced path handling for special characters and spaces
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
            Register-PSRepository -Default -ErrorAction Stop
        }
        if ((Get-PSRepository -Name "PSGallery").InstallationPolicy -ne "Trusted") {
            Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted -ErrorAction Stop
        }

        # Install the module
        Install-Module -Name ThreadJob -Repository PSGallery -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
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

# Check for elevated privileges and restart if necessary
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Elevated privileges required. Attempting to restart script as Administrator..."
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" `"$Path`"" -Verb RunAs
    Exit
}

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

function Format-Path {
    param (
        [string]$Path
    )
    try {
        # Convert to full path and normalize separators
        $fullPath = [System.IO.Path]::GetFullPath($Path.Trim())
        
        # Handle special characters by using .NET path normalization
        if ($fullPath -match '[&\s\(\)\[\]\{\}\^\#\$\%\+\,\;\=\@]' -or $fullPath.Contains('"')) {
            return [System.IO.Path]::GetFullPath($fullPath)
        }
        
        return $fullPath
    }
    catch {
        Write-Warning "Error formatting path '$Path': $($_.Exception.Message)"
        return $Path
    }
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

        # Get largest file in current directory
        $largestFile = & { Invoke-Expression "[$typeName]::GetLargestFile('$FolderPath')" }
        if ($largestFile) {
            Write-Host "`nLargest file in $FolderPath :"
            Write-Host "Name: $($largestFile.Name)"
            Write-Host "Size: $(Format-SizeWithPadding ($largestFile.Length / 1GB)) GB"
        }

        Write-Host "`nFound $($folders.Count) subfolders to process..."
        $folderSizes = @()
        $processedCount = 0
        $totalDirs = $folders.Count

        # Process in smaller batches
        $batchSize = if ($script:isLegacyPowerShell) { 5 } else { 10 }
        for ($i = 0; $i -lt $folders.Count; $i += $batchSize) {
            $batch = $folders | Select-Object -Skip $i -First $batchSize
            $results = @()

            Write-ProgressBar -Current $processedCount -Total $totalDirs -Status "Processing folder batch $([math]::Floor($i/$batchSize + 1)) of $([math]::Ceiling($totalDirs/$batchSize))"

            if ($global:useThreadJobs -and -not $script:isLegacyPowerShell) {
                $jobs = @()
                foreach ($dir in $batch) {
                    $dirPath = if ($dir.FullName) { $dir.FullName } else { $dir }
                    $sanitizedPath = Format-Path $dirPath
                    $jobs += Start-ThreadJob -ThrottleLimit 10 -ArgumentList $dirPath, $sanitizedPath, $typeName -ScriptBlock {
                        param($originalPath, $path, $className)
                        try {
                            $size = & {
                                Add-Type -TypeDefinition @"
                                    using System.IO;
                                    public static class PathHelper {
                                        public static string NormalizePath(string path) {
                                            return Path.GetFullPath(path);
                                        }
                                    }
"@
                                Invoke-Expression "[$className]::GetDirectorySize('$([PathHelper]::NormalizePath($path))')"
                            }
                            $counts = Invoke-Expression "[$className]::GetDirectoryCounts('$path')"
                            $largestFile = Invoke-Expression "[$className]::GetLargestFile('$path')"

                            return [PSCustomObject]@{
                                Folder = $originalPath  # Use original path for display
                                SizeGB = $size / 1GB
                                TotalFiles = $counts.Item1
                                TotalSubfolders = $counts.Item2
                                LargestFile = if ($largestFile) {
                                    [PSCustomObject]@{
                                        Name = $largestFile.Name
                                        Path = $largestFile.FullName
                                        SizeGB = $largestFile.Length / 1GB
                                        SizeMB = $largestFile.Length / 1MB
                                    }
                                } else { $null }
                            }
                        }
                        catch {
                            Write-Warning "Error processing $path : $_"
                            return $null
                        }
                    }
                }

                # Wait for batch completion
                $results = $jobs | Wait-Job | Receive-Job
                $jobs | Remove-Job -Force
            }
            else {
                # Fallback method - optimized for PS 4.0
                foreach ($dir in $batch) {
                    $dirPath = if ($dir.FullName) { $dir.FullName } else { $dir }
                    $sanitizedPath = Format-Path $dirPath
                    try {
                        $size = Invoke-Expression "[$typeName]::GetDirectorySize('$sanitizedPath')"
                        $counts = Invoke-Expression "[$typeName]::GetDirectoryCounts('$sanitizedPath')"
                        $largestFile = Invoke-Expression "[$typeName]::GetLargestFile('$sanitizedPath')"

                        $results += [PSCustomObject]@{
                            Folder = $dirPath
                            SizeGB = $size / 1GB
                            TotalFiles = $counts.Item1
                            TotalSubfolders = $counts.Item2
                            LargestFile = if ($largestFile) {
                                [PSCustomObject]@{
                                    Name = $largestFile.Name
                                    Path = $largestFile.FullName
                                    SizeGB = $largestFile.Length / 1GB
                                    SizeMB = $largestFile.Length / 1MB
                                }
                            } else { $null }
                        }
                    }
                    catch {
                        Write-Warning "Error processing $dirPath : $_"
                    }
                }
            }

            $folderSizes += @($results | Where-Object { $_ -ne $null })
            $processedCount += $batch.Count
        }

        Write-ProgressBar -Current $totalDirs -Total $totalDirs -Status "Completed processing all folders"
        Write-Host ""

        Write-Host "`nCompleted processing all folders."
        return $folderSizes
    }
    catch {
        Write-Warning "Error processing directory '$FolderPath'. Error: $($_.Exception.Message)"
        return $null
    }
}

function Write-TableLine {
    param([int]$Length = 150)
    Write-Host ("-" * $Length)
}

$currentPath = $Path

while ($true) {
    $folderSizes = Get-FolderSizes -FolderPath $currentPath

    if ($null -eq $folderSizes -or $folderSizes.Count -eq 0) {
        Write-Host "`nReached end of directory tree at: $currentPath"
        break
    }

    # Display the top N largest folders in a table format
    Write-Host "`nTop $Top Largest Folders in: $currentPath`n"
    Write-TableLine
    $format = "{0,-50} | {1,10} | {2,15} | {3,12} | {4,-50}"
    Write-Host ($format -f "Folder Path", "Size (GB)", "Subfolders", "Files", "Largest File (in this directory)")
    Write-TableLine

    $topFolders = $folderSizes | Sort-Object -Property SizeGB -Descending | Select-Object -First $Top
    foreach ($folder in $topFolders) {
        $largestFileInfo = if ($folder.LargestFile) {
            if ($folder.LargestFile.SizeGB -ge 1) {
                "$($folder.LargestFile.Name) ($(Format-SizeWithPadding $folder.LargestFile.SizeGB) GB)"
            } else {
                "$($folder.LargestFile.Name) ($(Format-SizeWithPadding $folder.LargestFile.SizeMB) MB)"
            }
        } else {
            "No files"
        }
        
        $folderPathDisplay = if ($folder.Folder.Length -gt 47) {
            "..." + $folder.Folder.Substring($folder.Folder.Length - 44)
        } else {
            $folder.Folder
        }

        $largestFileDisplay = if ($largestFileInfo.Length -gt 47) {
            "..." + $largestFileInfo.Substring($largestFileInfo.Length - 44)
        } else {
            $largestFileInfo
        }

        Write-Host ($format -f $folderPathDisplay,
            (Format-SizeWithPadding $folder.SizeGB),
            $folder.TotalSubfolders,
            $folder.TotalFiles,
            $largestFileDisplay)
    }
    Write-TableLine

    # Descend into the largest folder
    $largestFolder = $topFolders | Select-Object -First 1
    Write-Host "`nDescending into: $($largestFolder.Folder)`n"
    $currentPath = $largestFolder.Folder
}

Get-Volume $Path.Substring(0, [Math]::Min($Path.Length, 1))

Write-Host "`nScript completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host "Script completed (UTC): $((Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "`nTranscript log file can be found here: $transcriptFile"
Write-Host "======================================================"

Stop-Transcript