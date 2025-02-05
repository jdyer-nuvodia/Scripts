# get-foldersizes.ps1
# Author: jdyer-nuvodia
# Created: 2025-02-05 00:41:11 UTC
# Current User: jdyer-nuvodia
# Purpose: Ultra-fast directory scanner for large directories including system folders (read-only)

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

# Add .NET methods for high-performance file operations with system directory handling
Add-Type @"
using System;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Security.AccessControl;
using System.Collections.Generic;

public class FastFileScanner {
    public class FolderInfo {
        public string Path { get; set; }
        public long Size { get; set; }
        public int FileCount { get; set; }
        public int SubfolderCount { get; set; }
        public FileInfo LargestFile { get; set; }
    }

    public static FolderInfo ScanDirectory(string path) {
        try {
            var di = new DirectoryInfo(path);
            var files = new List<FileInfo>();
            
            try {
                files.AddRange(di.GetFiles("*", SearchOption.TopDirectoryOnly));
            }
            catch (UnauthorizedAccessException) {
                // Try alternate method for system directories
                foreach (string file in Directory.GetFiles(path, "*", SearchOption.TopDirectoryOnly)) {
                    try {
                        files.Add(new FileInfo(file));
                    }
                    catch { }
                }
            }

            var largestFile = files.OrderByDescending(f => f.Length).FirstOrDefault();
            var size = files.Sum(f => f.Length);

            int subFolderCount = 0;
            try {
                subFolderCount = di.GetDirectories("*", SearchOption.TopDirectoryOnly).Length;
            }
            catch {
                subFolderCount = Directory.GetDirectories(path, "*", SearchOption.TopDirectoryOnly).Length;
            }

            return new FolderInfo {
                Path = path,
                Size = size,
                FileCount = files.Count,
                SubfolderCount = subFolderCount,
                LargestFile = largestFile
            };
        }
        catch (Exception) {
            return null;
        }
    }

    public static long GetRecursiveSize(string path) {
        long size = 0;
        try {
            foreach (string file in Directory.GetFiles(path, "*", SearchOption.AllDirectories)) {
                try {
                    size += new FileInfo(file).Length;
                }
                catch { }
            }
        }
        catch { }
        return size;
    }

    public static int GetRecursiveFileCount(string path) {
        int count = 0;
        try {
            foreach (string file in Directory.GetFiles(path, "*", SearchOption.AllDirectories)) {
                try {
                    count++;
                }
                catch { }
            }
        }
        catch { }
        return count;
    }

    public static int GetRecursiveSubfolderCount(string path) {
        int count = 0;
        try {
            foreach (string dir in Directory.GetDirectories(path, "*", SearchOption.AllDirectories)) {
                try {
                    count++;
                }
                catch { }
            }
        }
        catch { }
        return count;
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
        # Scan current directory
        $currentInfo = [FastFileScanner]::ScanDirectory($FolderPath)
        if ($currentInfo.LargestFile) {
            Write-Host "`nLargest file in $FolderPath :"
            Write-Host "Name: $($currentInfo.LargestFile.Name)"
            Write-Host "Size: $(Format-SizeWithPadding ($currentInfo.LargestFile.Length / 1GB)) GB"
        }

        # Get subdirectories using .NET directly
        $subDirs = [System.IO.Directory]::GetDirectories($FolderPath)
        Write-Host "`nFound $($subDirs.Count) subfolders to process..."

        # Process directories in batches of 50 for better performance
        $batchSize = 50
        $folderSizes = @()
        $processedCount = 0
        $totalDirs = $subDirs.Count

        for ($i = 0; $i -lt $subDirs.Count; $i += $batchSize) {
            $batch = $subDirs | Select-Object -Skip $i -First $batchSize
            $jobs = @()

            foreach ($dir in $batch) {
                $jobs += Start-ThreadJob -ThrottleLimit 50 -ArgumentList $dir -ScriptBlock {
                    param($folderPath)
                    try {
                        $info = [FastFileScanner]::ScanDirectory($folderPath)
                        $recursiveSize = [FastFileScanner]::GetRecursiveSize($folderPath)
                        $recursiveFiles = [FastFileScanner]::GetRecursiveFileCount($folderPath)
                        $recursiveFolders = [FastFileScanner]::GetRecursiveSubfolderCount($folderPath)

                        return [PSCustomObject]@{
                            Folder = $folderPath
                            SizeGB = $recursiveSize / 1GB
                            TotalSubfolders = $recursiveFolders
                            TotalFiles = $recursiveFiles
                            LargestFile = if ($info.LargestFile) {
                                [PSCustomObject]@{
                                    Name = $info.LargestFile.Name
                                    Path = $info.LargestFile.FullName
                                    SizeGB = $info.LargestFile.Length / 1GB
                                    SizeMB = $info.LargestFile.Length / 1MB
                                }
                            } else { $null }
                        }
                    }
                    catch {
                        return $null
                    }
                }
            }

            # Wait for batch completion
            $results = $jobs | Wait-Job | Receive-Job
            $jobs | Remove-Job
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