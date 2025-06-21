# =============================================================================
# Script: Find-LargestFolders.ps1
# Created: 2025-06-20 22:50:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-06-20 22:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.0
# Additional Info: Efficient script to find largest folders by drilling down into only the largest subdirectories
# =============================================================================

<#
.SYNOPSIS
Efficiently finds the largest folders by recursively drilling down into only the largest subdirectories.

.DESCRIPTION
This script scans a directory and identifies the largest subdirectories and files without scanning the entire drive.
It works by:
1. Scanning the top-level directories in the specified path
2. Identifying the largest folders and files
3. Recursively drilling down into only the largest folder
4. Repeating this process until reaching the specified depth or finding no more subdirectories

This approach is much faster than scanning entire drives and focuses on finding the largest space consumers.

.PARAMETER StartPath
The root directory path to begin analysis. Default is "C:\"

.PARAMETER MaxDepth
Maximum recursion depth for drilling down. Default is 10 levels.

.PARAMETER TopCount
Number of largest folders/files to display at each level. Default is 10.

.PARAMETER MinSizeGB
Minimum size in GB to display a folder/file. Default is 0.1 GB (100 MB).

.PARAMETER WhatIf
Shows what would be analyzed without performing the actual scan.

.EXAMPLE
.\Find-LargestFolders.ps1
Analyzes C:\ and drills down into the largest folders.

.EXAMPLE
.\Find-LargestFolders.ps1 -StartPath "D:\Data" -MaxDepth 5 -TopCount 5
Analyzes D:\Data with max depth of 5 levels, showing top 5 items at each level.

.EXAMPLE
.\Find-LargestFolders.ps1 -StartPath "C:\Users" -MinSizeGB 1.0
Analyzes C:\Users and only shows items larger than 1 GB.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $false, Position = 0)]
    [ValidateScript({
        if (-not (Test-Path -Path $_ -PathType Container)) {
            throw "Path '$_' does not exist or is not a directory."
        }
        return $true
    })]
    [string]$StartPath = "C:\",

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 50)]
    [int]$MaxDepth = 10,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 50)]
    [int]$TopCount = 10,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, [double]::MaxValue)]
    [double]$MinSizeGB = 0.1
)

# Script variables
$Script:StartTime = Get-Date
$Script:ProcessedFolders = 0
$Script:CurrentDepth = 0

# Color codes for different PowerShell versions
if ($PSVersionTable.PSVersion.Major -ge 7) {
    $Script:Colors = @{
        Reset      = "`e[0m"
        White      = "`e[37m"
        Cyan       = "`e[36m"
        Green      = "`e[32m"
        Yellow     = "`e[33m"
        Red        = "`e[31m"
        Magenta    = "`e[35m"
        DarkGray   = "`e[90m"
        Bold       = "`e[1m"
    }
} else {
    $Script:Colors = @{
        Reset      = ""
        White      = ""
        Cyan       = ""
        Green      = ""
        Yellow     = ""
        Red        = ""
        Magenta    = ""
        DarkGray   = ""
        Bold       = ""
    }
}

function Format-FileSize {
    param(
        [Parameter(Mandatory = $true)]
        [long]$SizeInBytes
    )

    if ($SizeInBytes -ge 1TB) {
        return "{0:N2} TB" -f ($SizeInBytes / 1TB)
    } elseif ($SizeInBytes -ge 1GB) {
        return "{0:N2} GB" -f ($SizeInBytes / 1GB)
    } elseif ($SizeInBytes -ge 1MB) {
        return "{0:N2} MB" -f ($SizeInBytes / 1MB)
    } elseif ($SizeInBytes -ge 1KB) {
        return "{0:N2} KB" -f ($SizeInBytes / 1KB)
    } else {
        return "$SizeInBytes bytes"
    }
}

function Get-DirectorySize {
    param(
        [string]$Path
    )

    try {
        # Get all files in the directory (not subdirectories)
        $files = Get-ChildItem -Path $Path -File -Force -ErrorAction SilentlyContinue
        $totalSize = ($files | Measure-Object -Property Length -Sum).Sum

        if ($null -eq $totalSize) {
            $totalSize = 0
        }

        return [PSCustomObject]@{
            Path = $Path
            SizeBytes = $totalSize
            FileCount = $files.Count
            IsAccessible = $true
            Error = $null
        }
    } catch {
        return [PSCustomObject]@{
            Path = $Path
            SizeBytes = 0
            FileCount = 0
            IsAccessible = $false
            Error = $_.Exception.Message
        }
    }
}

function Get-SubdirectoryTotalSize {
    param(
        [string]$Path
    )

    try {
        # Get total size including all subdirectories using Get-ChildItem -Recurse
        $files = Get-ChildItem -Path $Path -File -Recurse -Force -ErrorAction SilentlyContinue
        $totalSize = ($files | Measure-Object -Property Length -Sum).Sum

        if ($null -eq $totalSize) {
            $totalSize = 0
        }

        return $totalSize
    } catch {
        Write-Warning "Cannot access directory: $Path - $($_.Exception.Message)"
        return 0
    }
}

function Get-LargestItems {
    param(
        [string]$Path,
        [int]$TopCount,
        [double]$MinSizeGB
    )

    $results = @()
    $minSizeBytes = $MinSizeGB * 1GB

    try {
        Write-Output "$($Script:Colors.Cyan)Scanning: $Path$($Script:Colors.Reset)"

        # Get subdirectories
        $subdirs = Get-ChildItem -Path $Path -Directory -Force -ErrorAction SilentlyContinue

        foreach ($subdir in $subdirs) {
            $Script:ProcessedFolders++
            Write-Progress -Activity "Analyzing Folders" -Status "Processing: $($subdir.Name)" -PercentComplete -1

            $totalSize = Get-SubdirectoryTotalSize -Path $subdir.FullName

            if ($totalSize -ge $minSizeBytes) {
                $results += [PSCustomObject]@{
                    Type = "Folder"
                    Name = $subdir.Name
                    Path = $subdir.FullName
                    SizeBytes = $totalSize
                    SizeFormatted = Format-FileSize -SizeInBytes $totalSize
                }
            }
        }

        # Get files in current directory
        $files = Get-ChildItem -Path $Path -File -Force -ErrorAction SilentlyContinue

        foreach ($file in $files) {
            if ($file.Length -ge $minSizeBytes) {
                $results += [PSCustomObject]@{
                    Type = "File"
                    Name = $file.Name
                    Path = $file.FullName
                    SizeBytes = $file.Length
                    SizeFormatted = Format-FileSize -SizeInBytes $file.Length
                }
            }
        }

        # Sort by size descending and take top items
        $topItems = $results | Sort-Object SizeBytes -Descending | Select-Object -First $TopCount

        return $topItems

    } catch {
        Write-Output "$($Script:Colors.Red)Error accessing path: $Path - $($_.Exception.Message)$($Script:Colors.Reset)"
        return @()
    }
}

function Show-LargestItemsTable {
    param(
        [array]$Items,
        [string]$Path,
        [int]$Depth
    )

    if ($Items.Count -eq 0) {
        Write-Output "$($Script:Colors.Yellow)No items found meeting the minimum size criteria.$($Script:Colors.Reset)"
        return
    }

    $indent = "  " * $Depth
    Write-Output ""
    Write-Output "$($Script:Colors.Bold)$($Script:Colors.Green)$indent=== LEVEL $($Depth + 1): $Path ===$($Script:Colors.Reset)"
    Write-Output ""

    # Create table
    $tableData = $Items | Select-Object Type, Name, SizeFormatted | Sort-Object @{Expression={$_.Type}; Descending=$true}, @{Expression={[long]($_.SizeFormatted -replace '[^0-9.]', '')}; Descending=$true}

    $tableData | Format-Table -Property @(
        @{Label="Type"; Expression={$_.Type}; Width=8},
        @{Label="Name"; Expression={$_.Name}; Width=50},
        @{Label="Size"; Expression={$_.SizeFormatted}; Width=12; Alignment="Right"}
    ) -AutoSize | Out-String | Write-Output
}

function Find-LargestFoldersRecursive {
    param(
        [string]$Path,
        [int]$CurrentDepth,
        [int]$MaxDepth,
        [int]$TopCount,
        [double]$MinSizeGB
    )

    if ($CurrentDepth -ge $MaxDepth) {
        Write-Output "$($Script:Colors.Yellow)Maximum depth ($MaxDepth) reached.$($Script:Colors.Reset)"
        return
    }

    # Get largest items at current level
    $largestItems = Get-LargestItems -Path $Path -TopCount $TopCount -MinSizeGB $MinSizeGB

    if ($largestItems.Count -eq 0) {
        Write-Output "$($Script:Colors.Yellow)No items found at current level meeting size criteria.$($Script:Colors.Reset)"
        return
    }

    # Show table for current level
    Show-LargestItemsTable -Items $largestItems -Path $Path -Depth $CurrentDepth

    # Find the largest folder (not file) to drill down into
    $largestFolder = $largestItems | Where-Object { $_.Type -eq "Folder" } | Select-Object -First 1

    if ($largestFolder) {
        Write-Output "$($Script:Colors.Cyan)Drilling down into largest folder: $($largestFolder.Name) ($($largestFolder.SizeFormatted))$($Script:Colors.Reset)"

        # Recursively analyze the largest folder
        Find-LargestFoldersRecursive -Path $largestFolder.Path -CurrentDepth ($CurrentDepth + 1) -MaxDepth $MaxDepth -TopCount $TopCount -MinSizeGB $MinSizeGB
    } else {
        Write-Output "$($Script:Colors.Green)No more folders to drill down into at this level.$($Script:Colors.Reset)"
    }
}

function Start-AdvancedTranscript {
    param([string]$LogPath)

    try {
        $computerName = $env:COMPUTERNAME
        $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $logFileName = "Find-LargestFolders_${computerName}_${timestamp}.log"
        $fullLogPath = Join-Path -Path $LogPath -ChildPath $logFileName

        # Ensure log directory exists
        if (-not (Test-Path -Path $LogPath)) {
            New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
        }

        # Create header
        $headerText = @"
===============================================
FIND LARGEST FOLDERS ANALYZER v1.0.0
===============================================
Log started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC')
Computer: $computerName
PowerShell Version: $($PSVersionTable.PSVersion)
Process ID: $PID
===============================================
"@
        Set-Content -Path $fullLogPath -Value $headerText -Encoding UTF8

        Start-Transcript -Path $fullLogPath -Append -Force | Out-Null
        Write-Output "$($Script:Colors.Green)Transcript started: $fullLogPath$($Script:Colors.Reset)"
        return $fullLogPath
    } catch {
        Write-Warning "Could not start transcript: $($_.Exception.Message)"
        return $null
    }
}

function Stop-AdvancedTranscript {
    try {
        Stop-Transcript | Out-Null
        Write-Output "$($Script:Colors.Green)Transcript stopped successfully$($Script:Colors.Reset)"
    } catch {
        Write-Warning "Error stopping transcript: $($_.Exception.Message)"
    }
}

# Main execution
try {
    # WhatIf support
    if ($PSCmdlet.ShouldProcess($StartPath, "Analyze largest folders")) {

        # Start transcript logging
        $transcriptPath = Start-AdvancedTranscript -LogPath $PSScriptRoot

        # Normalize the start path
        $StartPath = $StartPath.TrimEnd('\')

        Write-Output "$($Script:Colors.Bold)$($Script:Colors.Green)Find Largest Folders Analyzer v1.0.0$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.Green)===============================================$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Start Path: $($Script:Colors.Cyan)$StartPath$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Max Depth: $($Script:Colors.Cyan)$MaxDepth$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Top Count: $($Script:Colors.Cyan)$TopCount$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Min Size: $($Script:Colors.Cyan)$MinSizeGB GB$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Started: $($Script:Colors.Cyan)$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')$($Script:Colors.Reset)"
        Write-Output ""

        # Validate start path
        if (-not (Test-Path -Path $StartPath -PathType Container)) {
            throw "Start path does not exist or is not accessible: $StartPath"
        }

        # Start the recursive analysis
        Find-LargestFoldersRecursive -Path $StartPath -CurrentDepth 0 -MaxDepth $MaxDepth -TopCount $TopCount -MinSizeGB $MinSizeGB

        # Final summary
        $endTime = Get-Date
        $duration = $endTime - $Script:StartTime

        Write-Output ""
        Write-Output "$($Script:Colors.Bold)$($Script:Colors.Green)Analysis Complete$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.Green)==================$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Folders Processed: $($Script:Colors.Cyan)$Script:ProcessedFolders$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Total Duration: $($Script:Colors.Cyan)$($duration.ToString('hh\:mm\:ss'))$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Completed: $($Script:Colors.Cyan)$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')$($Script:Colors.Reset)"

        if ($transcriptPath) {
            Write-Output "$($Script:Colors.White)Log File: $($Script:Colors.Cyan)$transcriptPath$($Script:Colors.Reset)"
        }

        Write-Output ""
    }
} catch {
    Write-Output "$($Script:Colors.Red)Script execution failed: $($_.Exception.Message)$($Script:Colors.Reset)"
    Write-Output "$($Script:Colors.Red)$($_.ScriptStackTrace)$($Script:Colors.Reset)"
} finally {
    # Stop transcript
    Stop-AdvancedTranscript
}
