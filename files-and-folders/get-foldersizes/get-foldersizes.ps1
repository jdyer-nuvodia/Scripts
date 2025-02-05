# get-foldersizes.ps1
# Author: jdyer-nuvodia
# Last Modified: 2025-02-05 00:32:45 UTC
# Purpose: High-performance directory scanner for finding largest folders and files (read-only)

param (
    [string]$Path = "C:\",
    [int]$MaxDepth = 10,
    [int]$ThrottleLimit = 50 # Number of concurrent operations
)

# Check for elevated privileges and restart if necessary
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Elevated privileges required. Attempting to restart script as Administrator..."
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" `"$Path`"" -Verb RunAs
    Exit
}

# Set global error action preference
$ErrorActionPreference = 'SilentlyContinue'

# Import required assemblies for faster file operations
Add-Type -TypeDefinition @"
    using System;
    using System.IO;
    public class FastFileInfo
    {
        public static long GetDirectorySize(string path)
        {
            long size = 0;
            try
            {
                var di = new DirectoryInfo(path);
                foreach (var fi in di.GetFiles("*", SearchOption.AllDirectories))
                {
                    size += fi.Length;
                }
            }
            catch { }
            return size;
        }
    }
"@

Write-Host "Analyzing folders in: $Path (Read-only scan - Optimized)"
Write-Host "Script started by: $env:USERNAME at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

function Get-FolderSizes {
    param (
        [string]$FolderPath,
        [int]$CurrentDepth = 0
    )

    if ($CurrentDepth -ge $MaxDepth) {
        return $null
    }

    try {
        # Get largest file in current directory first (optimized)
        $largestCurrentFile = Get-ChildItem -Path $FolderPath -File -Force | 
            Sort-Object -Property Length -Descending | 
            Select-Object -First 1

        if ($largestCurrentFile) {
            Write-Host "`nLargest file in $FolderPath :"
            Write-Host "Name: $($largestCurrentFile.Name)"
            Write-Host "Size: $([math]::round($largestCurrentFile.Length / 1GB, 2)) GB ($([math]::round($largestCurrentFile.Length / 1MB, 2)) MB)"
        }

        # Get all subdirectories
        $folders = @(Get-ChildItem -Path $FolderPath -Directory -Force)
        $folderSizes = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
        $totalItems = $folders.Count
        Write-Host "`nFound $totalItems subfolders to process..."
        $processedCount = 0

        # Process folders in parallel
        $folders | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
            try {
                # Get immediate files in the current directory
                $currentFiles = Get-ChildItem -Path $_.FullName -File -Force
                $largestCurrentFile = $currentFiles | Sort-Object -Property Length -Descending | Select-Object -First 1

                # Calculate directory size using optimized method
                $folderSize = [FastFileInfo]::GetDirectorySize($_.FullName)
                
                $folderInfo = [PSCustomObject]@{
                    Folder = $_.FullName
                    SizeGB = [math]::round($folderSize / 1GB, 2)
                    TotalSubfolders = (Get-ChildItem -Path $_.FullName -Directory -Recurse -Force | Measure-Object).Count
                    TotalFiles = (Get-ChildItem -Path $_.FullName -File -Recurse -Force | Measure-Object).Count
                    LargestFile = if ($largestCurrentFile) {
                        [PSCustomObject]@{
                            Name = $largestCurrentFile.Name
                            Path = $largestCurrentFile.FullName
                            SizeGB = [math]::round($largestCurrentFile.Length / 1GB, 2)
                            SizeMB = [math]::round($largestCurrentFile.Length / 1MB, 2)
                        }
                    } else { $null }
                }
                $using:folderSizes.Add($folderInfo)
                $using:processedCount++
                Write-Host "`rProcessed $($using:processedCount) of $($using:totalItems) folders..." -NoNewline
            }
            catch {
                Write-Warning "Cannot access path '$($_.FullName)'. Error: $($_.Exception.Message)"
            }
        }

        Write-Host "`nCompleted processing $processedCount folders."
        return @($folderSizes)
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

function Format-FileSize {
    param (
        [double]$SizeInBytes
    )
    if ($SizeInBytes -ge 1GB) {
        return "$([math]::round($SizeInBytes / 1GB, 2)) GB"
    } elseif ($SizeInBytes -ge 1MB) {
        return "$([math]::round($SizeInBytes / 1MB, 2)) MB"
    } else {
        return "$([math]::round($SizeInBytes / 1KB, 2)) KB"
    }
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
                "$($folder.LargestFile.Name) ($($folder.LargestFile.SizeGB) GB)"
            } else {
                "$($folder.LargestFile.Name) ($($folder.LargestFile.SizeMB) MB)"
            }
        } else {
            "No files"
        }
        
        Write-Host ($format -f 
            ($folder.Folder.Length -gt 47 ? "..." + $folder.Folder.Substring($folder.Folder.Length - 44) : $folder.Folder),
            $folder.SizeGB,
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