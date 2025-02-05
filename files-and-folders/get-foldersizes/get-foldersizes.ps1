# get-foldersizes.ps1
# Author: jdyer-nuvodia
# Created: 2025-02-05 00:55:03 UTC
# Purpose: Ultra-fast directory scanner including system directories (read-only)

param (
    [string]$Path = "C:\",
    [int]$MaxDepth = 10
)

# Check for elevated privileges and restart if necessary
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Elevated privileges required. Attempting to restart script as Administrator..."
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" `"$Path`"" -Verb RunAs
    Exit
}

# Setup transcript logging with UTF8 encoding for proper table formatting
$transcriptPath = "C:\temp"
if (-not (Test-Path $transcriptPath)) {
    New-Item -ItemType Directory -Path $transcriptPath -Force | Out-Null
}
$transcriptFile = Join-Path $transcriptPath "FolderScan_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').txt"
Start-Transcript -Path $transcriptFile -Force -UseMinimalHeader

# Add script header to transcript
Write-Host "======================================================"
Write-Host "Folder Size Scanner - Execution Log"
Write-Host "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host "Started (UTC): $((Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "User: $env:USERNAME"
Write-Host "Computer: $env:COMPUTERNAME"
Write-Host "Target Path: $Path"
Write-Host "======================================================"
Write-Host ""

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
                size += Directory.GetFiles(dir).Sum(f => {
                    try { return new FileInfo(f).Length; }
                    catch { return 0; }
                });

                foreach (var subDir in Directory.GetDirectories(dir)) {
                    stack.Push(subDir);
                }
            }
            catch (UnauthorizedAccessException) { }
            catch (SecurityException) { }
            catch (IOException) { }
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

function Get-FolderSizes {
    param (
        [string]$FolderPath,
        [int]$CurrentDepth = 0
    )

    if ($CurrentDepth -ge $MaxDepth) { return $null }

    try {
        # Get all subdirectories
        $folders = @(Get-ChildItem -Path $FolderPath -Directory -Force -ErrorAction SilentlyContinue)
        if (-not $folders) {
            $folders = @([System.IO.Directory]::GetDirectories($FolderPath))
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
        $batchSize = 10
        for ($i = 0; $i -lt $folders.Count; $i += $batchSize) {
            $batch = $folders | Select-Object -Skip $i -First $batchSize
            $jobs = @()

            foreach ($dir in $batch) {
                $dirPath = $dir.FullName ?? $dir
                $jobs += Start-ThreadJob -ThrottleLimit 10 -ArgumentList $dirPath, $typeName -ScriptBlock {
                    param($path, $className)
                    try {
                        $size = Invoke-Expression "[$className]::GetDirectorySize('$path')"
                        $counts = Invoke-Expression "[$className]::GetDirectoryCounts('$path')"
                        $largestFile = Invoke-Expression "[$className]::GetLargestFile('$path')"

                        return [PSCustomObject]@{
                            Folder = $path
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
            $folderSizes += @($results | Where-Object { $_ -ne $null })
            $processedCount += $batch.Count
            Write-Host "`rProcessed $processedCount of $totalDirs folders..." -NoNewline
        }

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

    # Display the top 3 largest folders in a table format
    Write-Host "`nTop 3 Largest Folders in: $currentPath`n"
    Write-TableLine
    $format = "{0,-50} | {1,10} | {2,15} | {3,12} | {4,-50}"
    Write-Host ($format -f "Folder Path", "Size (GB)", "Subfolders", "Files", "Largest File (in this directory)")
    Write-TableLine

    $topFolders = $folderSizes | Sort-Object -Property SizeGB -Descending | Select-Object -First 3
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
        
        Write-Host ($format -f 
            ($folder.Folder.Length -gt 47 ? "..." + $folder.Folder.Substring($folder.Folder.Length - 44) : $folder.Folder),
            (Format-SizeWithPadding $folder.SizeGB),
            $folder.TotalSubfolders,
            $folder.TotalFiles,
            ($largestFileInfo.Length -gt 47 ? "..." + $largestFileInfo.Substring($largestFileInfo.Length - 44) : $largestFileInfo)
        )
    }
    Write-TableLine

    # Descend into the largest folder
    $largestFolder = $topFolders | Select-Object -First 1
    Write-Host "`nDescending into: $($largestFolder.Folder)`n"
    $currentPath = $largestFolder.Folder
}

Write-Host "`nScript completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host "Script completed (UTC): $((Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "`nTranscript log file can be found here: $transcriptFile"
Write-Host "======================================================"

Stop-Transcript