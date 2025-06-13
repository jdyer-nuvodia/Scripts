# =============================================================================
# Script: Analyze-FolderUsage.ps1
# Created: 2025-06-13 20:57:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-06-13 23:29:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.5.0
# Additional Info: Fixed hierarchical display, improved space accounting, integrated safely scanned folders display
# =============================================================================

<#
.SYNOPSIS
Performs ultra-fast recursive folder size analysis with parallel processing and advanced features.

.DESCRIPTION
This script recursively scans directories starting from a specified path and calculates comprehensive folder statistics including:
- Folder sizes with GB formatting
- File and subfolder counts
- Largest file identification per directory
- OneDrive/cloud storage placeholder detection
- NTFS junction point and symbolic link handling
- Parallel processing with configurable thread limits
- Administrative privilege detection with graceful degradation
- Comprehensive error handling for access-denied scenarios
- Multi-level progress reporting with ANSI color coding
- Advanced transcript logging with cleanup

The script uses PowerShell runspaces for maximum performance and includes memory management with garbage collection.
Supports Windows PowerShell 5.1 and later with modern .NET integration.

.PARAMETER StartPath
The root directory path to begin analysis. Default is "C:\"
Type: String

.PARAMETER MaxDepth
Maximum recursion depth for directory scanning. Default is 10 levels.
Type: Integer

.PARAMETER Top
Number of largest folders to display in results. Default is 10.
Type: Integer

.PARAMETER MaxThreads
Maximum number of concurrent threads for parallel processing. Default is 10.
Type: Integer

.PARAMETER WhatIf
Shows what would be analyzed without performing the actual scan.
Type: Switch

.EXAMPLE
.\Analyze-FolderUsage.ps1
Analyzes the C:\ drive with default settings (max depth 10, top 3 folders, 10 threads).

.EXAMPLE
.\Analyze-FolderUsage.ps1 -StartPath "D:\Data" -MaxDepth 5 -Top 5 -MaxThreads 20
Analyzes D:\Data with custom depth limit of 5 levels, showing top 5 largest folders using 20 threads.

.EXAMPLE
.\Analyze-FolderUsage.ps1 -StartPath "C:\Users" -WhatIf
Shows what would be analyzed in the C:\Users directory without performing the scan.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $false, Position = 0)]
    [ValidateScript({
        if (Test-Path -Path $_ -PathType Container) {
            $true
        } else {
            throw "The specified path '$_' does not exist or is not a directory."
        }
    })]
    [string]$StartPath = "C:\",

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 50)]
    [int]$MaxDepth = 10,    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 100)]
    [int]$Top = 3,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 50)]
    [int]$MaxThreads = 10
)

# ANSI Color Codes for Enhanced Console Output
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
    Underline  = "`e[4m"
}

# Script Variables for Script Operation
$Script:ErrorTracker = @{}
$Script:IsAdmin = $false
$Script:ProcessedFolders = 0
$Script:TotalFolders = 0
$Script:StartTime = Get-Date

# Advanced Transcript Management Functions
function Start-AdvancedTranscript {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$LogPath
    )    try {
        # Validate and normalize the log path
        if ([string]::IsNullOrWhiteSpace($LogPath)) {
            $LogPath = $PWD.Path
        }

        $computerName = $env:COMPUTERNAME
        $timestamp = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")
        $logFileName = "Analyze-FolderUsage_${computerName}_${timestamp}.log"
        $fullLogPath = Join-Path -Path $LogPath -ChildPath $logFileName

        # Ensure log directory exists
        if (-not (Test-Path -Path $LogPath)) {
            if ($PSCmdlet.ShouldProcess($LogPath, "Create log directory")) {
                New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
            }
        }        if ($PSCmdlet.ShouldProcess($fullLogPath, "Start transcript")) {
            Start-Transcript -Path $fullLogPath -Force | Out-Null
            Write-Output "$($Script:Colors.Green)Transcript started: $fullLogPath$($Script:Colors.Reset)"
        }
        return $fullLogPath
    }
    catch {
        Write-Output "$($Script:Colors.Yellow)Warning: Could not start transcript: $($_.Exception.Message). Continuing without logging.$($Script:Colors.Reset)"
        return $null
    }
}

function Stop-AdvancedTranscript {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$LogPath
    )

    try {
        if ($PSCmdlet.ShouldProcess("Transcript", "Stop transcript")) {
            # First try normal stop
            Stop-Transcript -ErrorAction Stop
            Write-Output "$($Script:Colors.Green)Transcript stopped successfully$($Script:Colors.Reset)"
            if ($LogPath -and (Test-Path -Path $LogPath)) {
                Write-Output "$($Script:Colors.Green)Transcript saved: $LogPath$($Script:Colors.Reset)"
            }
        }
    }
    catch {
        # Multiple fallback methods for transcript cleanup
        Write-Output "$($Script:Colors.Yellow)Warning: Normal transcript stop failed. Attempting cleanup methods...$($Script:Colors.Reset)"

        try {
            # Method 1: Force stop with error suppression
            Stop-Transcript -ErrorAction SilentlyContinue
            Write-Output "$($Script:Colors.Green)Transcript stopped using fallback method 1$($Script:Colors.Reset)"
        }
        catch {
            try {
                # Method 2: Direct registry cleanup (Windows PowerShell specific)
                if ($PSVersionTable.PSVersion.Major -le 5) {
                    $regPath = "HKCU:\Software\Microsoft\PowerShell\1\PowerShellEngine"
                    if (Test-Path -Path $regPath) {
                        Remove-ItemProperty -Path $regPath -Name "TranscriptPath" -ErrorAction SilentlyContinue
                        Write-Output "$($Script:Colors.Green)Transcript stopped using registry cleanup$($Script:Colors.Reset)"
                    }
                }
            }
            catch {
                try {
                    # Method 3: Clear transcript variable directly
                    $ExecutionContext.InvokeCommand.InvokeScript('$null = Stop-Transcript') 2>$null
                    Write-Output "$($Script:Colors.Green)Transcript stopped using execution context$($Script:Colors.Reset)"
                }
                catch {
                    Write-Output "$($Script:Colors.Red)Warning: Could not properly stop transcript. File may remain locked.$($Script:Colors.Reset)"
                }
            }
        }

        # Final attempt: Force garbage collection to release file handles
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
    }
}

function Test-AdminPrivilege {
    <#
    .SYNOPSIS
    Detects if the current session has administrative privileges.
    #>
    try {
        $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function Write-ProgressUpdate {
    param(
        [string]$Activity,
        [string]$Status,
        [int]$PercentComplete = -1,
        [string]$Color = "Cyan",
        [switch]$Completed
    )

    $colorCode = $Script:Colors[$Color]
    $resetCode = $Script:Colors.Reset

    if ($Completed) {
        Write-Progress -Activity $Activity -Completed
    } elseif ($PercentComplete -ge 0) {
        Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete
    } else {
        Write-Progress -Activity $Activity -Status $Status
    }

    if (-not $Completed) {
        Write-Output "${colorCode}${Activity}: ${Status}${resetCode}"
    }
}

function Test-CloudPlaceholder {
    param(
        [System.IO.FileInfo]$FileInfo
    )

    try {
        # More specific OneDrive/cloud storage detection
        # Only check for actual cloud placeholder files, not all reparse points

        # Check for OneDrive specific file extensions and patterns
        if ($FileInfo.Extension -eq ".odlocal" -or $FileInfo.Name -like "*.onedrive") {
            return $true
        }

        # Check for cloud storage attributes (FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS)
        # This is more specific to actual cloud placeholders than ReparsePoint
        if (($FileInfo.Attributes.value__ -band 0x400000) -eq 0x400000) {
            return $true
        }

        # Check if it's in a known OneDrive path structure
        if ($FileInfo.FullName -like "*OneDrive*" -and ($FileInfo.Attributes -band [System.IO.FileAttributes]::ReparsePoint)) {
            return $true
        }

        return $false
    }
    catch {
        return $false
    }
}

function Test-ReparsePoint {
    param(
        [string]$Path
    )

    try {
        $item = Get-Item -Path $Path -Force -ErrorAction Stop
        return ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -eq [System.IO.FileAttributes]::ReparsePoint
    }
    catch {
        return $false
    }
}

function Get-FolderStatistic {
    param(
        [string]$FolderPath,
        [int]$CurrentDepth,
        [int]$MaxDepth
    )

    $stats = [PSCustomObject]@{
        Path = $FolderPath
        SizeBytes = 0
        FileCount = 0
        SubfolderCount = 0
        LargestFile = $null
        LargestFileSize = 0
        IsAccessible = $true
        HasCloudFiles = $false
        Error = $null
    }

    try {
        # Safety check for reparse points
        if (Test-ReparsePoint -Path $FolderPath) {
            Write-Output "$($Script:Colors.Yellow)Skipping reparse point: $FolderPath$($Script:Colors.Reset)"
            return $stats
        }

        # Get directory items with comprehensive error handling
        $items = Get-ChildItem -Path $FolderPath -Force -ErrorAction Stop

        foreach ($item in $items) {
            if ($item.PSIsContainer) {
                $stats.SubfolderCount++

                # Recursive processing within depth limits
                if ($CurrentDepth -lt $MaxDepth) {
                    $subStats = Get-FolderStatistic -FolderPath $item.FullName -CurrentDepth ($CurrentDepth + 1) -MaxDepth $MaxDepth
                    $stats.SizeBytes += $subStats.SizeBytes
                    $stats.FileCount += $subStats.FileCount
                    $stats.SubfolderCount += $subStats.SubfolderCount

                    if ($subStats.LargestFileSize -gt $stats.LargestFileSize) {
                        $stats.LargestFile = $subStats.LargestFile
                        $stats.LargestFileSize = $subStats.LargestFileSize
                    }

                    if ($subStats.HasCloudFiles) {
                        $stats.HasCloudFiles = $true
                    }
                }
            }
            else {
                $stats.FileCount++

                # Handle cloud placeholder files
                if (Test-CloudPlaceholder -FileInfo $item) {
                    $stats.HasCloudFiles = $true
                    # For cloud files, use the actual size if available, otherwise minimal size
                    $fileSize = if ($item.Length -gt 0) { $item.Length } else { 1KB }
                } else {
                    $fileSize = $item.Length
                }

                $stats.SizeBytes += $fileSize

                # Track largest file
                if ($fileSize -gt $stats.LargestFileSize) {
                    $stats.LargestFile = $item.FullName
                    $stats.LargestFileSize = $fileSize
                }
            }
        }

        $Script:ProcessedFolders++
    }
    catch [System.UnauthorizedAccessException] {
        $stats.IsAccessible = $false
        $stats.Error = "Access Denied"
        $Script:ErrorTracker[$FolderPath] = "Access Denied"
        Write-Output "$($Script:Colors.Red)Access denied: $FolderPath$($Script:Colors.Reset)"
    }
    catch [System.IO.DirectoryNotFoundException] {
        $stats.IsAccessible = $false
        $stats.Error = "Directory Not Found"
        $Script:ErrorTracker[$FolderPath] = "Directory Not Found"
        Write-Output "$($Script:Colors.Red)Directory not found: $FolderPath$($Script:Colors.Reset)"
    }
    catch {
        $stats.IsAccessible = $false
        $stats.Error = $_.Exception.Message
        $Script:ErrorTracker[$FolderPath] = $_.Exception.Message
        Write-Output "$($Script:Colors.Red)Error processing $FolderPath`: $($_.Exception.Message)$($Script:Colors.Reset)"
    }

    return $stats
}

function Get-ParallelFolderStatistic {
    param(
        [string[]]$FolderPaths,
        [int]$MaxDepth,
        [int]$MaxThreads
    )

    $runspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads)
    $runspacePool.Open()

    $jobs = @()
    $results = @()

    try {
        # Create scriptblock for parallel execution
        $scriptBlock = {
            param($FolderPath, $MaxDepth)            # Import required functions into runspace
            function Test-CloudPlaceholder {
                param([System.IO.FileInfo]$FileInfo)
                try {
                    # More specific OneDrive/cloud storage detection
                    # Only check for actual cloud placeholder files, not all reparse points

                    # Check for OneDrive specific file extensions and patterns
                    if ($FileInfo.Extension -eq ".odlocal" -or $FileInfo.Name -like "*.onedrive") {
                        return $true
                    }

                    # Check for cloud storage attributes (FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS)
                    # This is more specific to actual cloud placeholders than ReparsePoint
                    if (($FileInfo.Attributes.value__ -band 0x400000) -eq 0x400000) {
                        return $true
                    }

                    # Check if it's in a known OneDrive path structure
                    if ($FileInfo.FullName -like "*OneDrive*" -and ($FileInfo.Attributes -band [System.IO.FileAttributes]::ReparsePoint)) {
                        return $true
                    }

                    return $false
                } catch { return $false }
            }

            function Test-ReparsePoint {
                param([string]$Path)
                try {
                    $item = Get-Item -Path $Path -Force -ErrorAction Stop
                    return ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -eq [System.IO.FileAttributes]::ReparsePoint
                } catch { return $false }
            }

            function Get-FolderStatisticInternal {
                param([string]$FolderPath, [int]$CurrentDepth, [int]$MaxDepth)

                $stats = [PSCustomObject]@{
                    Path = $FolderPath
                    SizeBytes = 0
                    FileCount = 0
                    SubfolderCount = 0
                    LargestFile = $null
                    LargestFileSize = 0
                    IsAccessible = $true
                    HasCloudFiles = $false
                    Error = $null
                }

                try {
                    if (Test-ReparsePoint -Path $FolderPath) { return $stats }

                    $items = Get-ChildItem -Path $FolderPath -Force -ErrorAction Stop

                    foreach ($item in $items) {
                        if ($item.PSIsContainer) {
                            $stats.SubfolderCount++
                            if ($CurrentDepth -lt $MaxDepth) {
                                $subStats = Get-FolderStatisticInternal -FolderPath $item.FullName -CurrentDepth ($CurrentDepth + 1) -MaxDepth $MaxDepth
                                $stats.SizeBytes += $subStats.SizeBytes
                                $stats.FileCount += $subStats.FileCount
                                $stats.SubfolderCount += $subStats.SubfolderCount

                                if ($subStats.LargestFileSize -gt $stats.LargestFileSize) {
                                    $stats.LargestFile = $subStats.LargestFile
                                    $stats.LargestFileSize = $subStats.LargestFileSize
                                }

                                if ($subStats.HasCloudFiles) { $stats.HasCloudFiles = $true }
                            }
                        } else {
                            $stats.FileCount++

                            if (Test-CloudPlaceholder -FileInfo $item) {
                                $stats.HasCloudFiles = $true
                                $fileSize = if ($item.Length -gt 0) { $item.Length } else { 1KB }
                            } else {
                                $fileSize = $item.Length
                            }

                            $stats.SizeBytes += $fileSize

                            if ($fileSize -gt $stats.LargestFileSize) {
                                $stats.LargestFile = $item.FullName
                                $stats.LargestFileSize = $fileSize
                            }
                        }
                    }
                } catch {
                    $stats.IsAccessible = $false
                    $stats.Error = $_.Exception.Message
                }

                return $stats
            }

            return Get-FolderStatisticInternal -FolderPath $FolderPath -CurrentDepth 0 -MaxDepth $MaxDepth
        }

        # Launch parallel jobs
        foreach ($folder in $FolderPaths) {
            $powershell = [powershell]::Create()
            $powershell.RunspacePool = $runspacePool
            $powershell.AddScript($scriptBlock).AddParameter("FolderPath", $folder).AddParameter("MaxDepth", $MaxDepth) | Out-Null

            $jobs += [PSCustomObject]@{
                PowerShell = $powershell
                Handle = $powershell.BeginInvoke()
                Path = $folder
            }
        }        # Collect results with progress tracking and timeout
        $completed = 0
        $timeout = (Get-Date).AddMinutes(10)  # 10-minute timeout to handle large directories like Windows
        $lastProgress = Get-Date
        $progressCheckInterval = 500  # 0.5 second

        while ($jobs | Where-Object { -not $_.Handle.IsCompleted }) {
            $completedNow = ($jobs | Where-Object { $_.Handle.IsCompleted }).Count
            if ($completedNow -gt $completed) {
                $completed = $completedNow
                $percentComplete = [math]::Round(($completed / $jobs.Count) * 100, 0)
                Write-ProgressUpdate -Activity "Parallel Processing" -Status "Completed $completed of $($jobs.Count) folders" -PercentComplete $percentComplete
                $lastProgress = Get-Date
            }

            # Check for timeout or stuck jobs
            if ((Get-Date) -gt $timeout) {
                Write-Output "$($Script:Colors.Yellow)Warning: Processing timeout reached. Forcing completion of remaining jobs.$($Script:Colors.Reset)"
                break
            }

            # Check if no progress for 30 seconds (reduced from 1 minute)
            if ((Get-Date).Subtract($lastProgress).TotalSeconds -gt 30) {
                Write-Output "$($Script:Colors.Yellow)Warning: No progress for 30 seconds. Checking for stuck jobs.$($Script:Colors.Reset)"
                $stuckJobs = $jobs | Where-Object { -not $_.Handle.IsCompleted }
                if ($stuckJobs.Count -le 3) {
                    foreach ($stuckJob in $stuckJobs) {
                        Write-Output "$($Script:Colors.Yellow)  Stuck job for path: $($stuckJob.Path)$($Script:Colors.Reset)"
                    }
                }
                else {
                    Write-Output "$($Script:Colors.Yellow)  $($stuckJobs.Count) jobs appear stuck$($Script:Colors.Reset)"
                }
                $lastProgress = Get-Date
            }

            Start-Sleep -Milliseconds $progressCheckInterval
        }

        # Gather all results with improved error handling
        foreach ($job in $jobs) {
            try {                if ($job.Handle.IsCompleted) {
                    $result = $job.PowerShell.EndInvoke($job.Handle)
                    if ($result) {
                        $results += $result
                        $Script:ProcessedFolders++  # Count successfully processed folders
                    }
                }else {
                    Write-Output "$($Script:Colors.Yellow)Warning: Job for $($job.Path) did not complete. Skipping.$($Script:Colors.Reset)"
                    # Try to stop the incomplete job
                    try {
                        $job.PowerShell.Stop()
                    } catch {
                        Write-Verbose "Unable to stop PowerShell job: $($_.Exception.Message)"
                    }
                }
            }
            catch {
                Write-Output "$($Script:Colors.Red)Error in parallel job for $($job.Path): $($_.Exception.Message)$($Script:Colors.Reset)"
            }
            finally {
                try {
                    $job.PowerShell.Dispose()
                } catch {
                    Write-Verbose "Unable to dispose PowerShell job: $($_.Exception.Message)"
                }
            }
        }
    }
    finally {
        # Cleanup runspace pool with timeout
        try {
            $runspacePool.Close()
            $runspacePool.Dispose()
        }
        catch {
            Write-Output "$($Script:Colors.Yellow)Warning: Error during runspace cleanup: $($_.Exception.Message)$($Script:Colors.Reset)"
        }

        # Force garbage collection for memory management
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
    }

    return $results
}

function Format-FileSize {
    param(
        [long]$SizeInBytes
    )

    $sizes = @("B", "KB", "MB", "GB", "TB")
    $index = 0
    $size = [double]$SizeInBytes

    while ($size -ge 1024 -and $index -lt $sizes.Length - 1) {
        $size = $size / 1024
        $index++
    }

    return "{0:N2} {1}" -f $size, $sizes[$index]
}

function Show-HierarchicalResult{
    param(
        [array]$Results,
        [string]$StartPath,
        [int]$Top,
        [array]$SafeResults = @()
    )

    Write-Output "`n$($Script:Colors.Bold)$($Script:Colors.Underline)HIERARCHICAL FOLDER USAGE ANALYSIS$($Script:Colors.Reset)`n"

    # First, show the StartPath itself
    Write-Output "$($Script:Colors.Bold)$($Script:Colors.Cyan)LEVEL 0: ROOT ANALYSIS ($StartPath)$($Script:Colors.Reset)"
    $rootResult = $Results | Where-Object { $_.Path -eq $StartPath -and $_.IsAccessible }
    if ($rootResult) {
        Show-SingleTable -Results @($rootResult) -Title "Root Directory"
    } else {
        # Calculate root stats from all results
        $accessibleResults = $Results | Where-Object { $_.IsAccessible }
        $totalSize = ($accessibleResults | Measure-Object -Property SizeBytes -Sum).Sum
        $totalFiles = ($accessibleResults | Measure-Object -Property FileCount -Sum).Sum
        $totalFolders = ($accessibleResults | Measure-Object -Property SubfolderCount -Sum).Sum

        # Add safe results to totals
        if ($SafeResults.Count -gt 0) {
            $safeSize = ($SafeResults | Measure-Object -Property SizeBytes -Sum).Sum
            $safeFiles = ($SafeResults | Measure-Object -Property FileCount -Sum).Sum
            $safeFolders = ($SafeResults | Measure-Object -Property SubfolderCount -Sum).Sum
            $totalSize += $safeSize
            $totalFiles += $safeFiles
            $totalFolders += $safeFolders
        }

        Write-Output "$($Script:Colors.White)Root Path: $($Script:Colors.Green)$StartPath$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Total Size: $($Script:Colors.Green)$(Format-FileSize -SizeInBytes $totalSize)$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Total Files: $($Script:Colors.Green)$($totalFiles.ToString('N0'))$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Total Folders: $($Script:Colors.Green)$($totalFolders.ToString('N0'))$($Script:Colors.Reset)"
    }

    # Level 1: Show top-level subfolders
    Write-Output "`n$($Script:Colors.Bold)$($Script:Colors.Cyan)LEVEL 1: TOP SUBFOLDERS OF $StartPath$($Script:Colors.Reset)"
    $level1Folders = $Results | Where-Object {
        $_.IsAccessible -and
        $_.Path -ne $StartPath -and
        (Split-Path -Path $_.Path -Parent) -eq $StartPath
    } | Sort-Object SizeBytes -Descending | Select-Object -First $Top

    if ($level1Folders.Count -gt 0) {
        Show-SingleTable -Results $level1Folders -Title "Level 1 Subfolders"
        $largestLevel1 = $level1Folders[0]

        # Level 2: Show subfolders of the largest Level 1 folder
        if ($largestLevel1) {
            Write-Output "`n$($Script:Colors.Bold)$($Script:Colors.Cyan)LEVEL 2: TOP SUBFOLDERS OF $($largestLevel1.Path)$($Script:Colors.Reset)"
            $level2Folders = $Results | Where-Object {
                $_.IsAccessible -and
                $_.Path -ne $largestLevel1.Path -and
                (Split-Path -Path $_.Path -Parent) -eq $largestLevel1.Path
            } | Sort-Object SizeBytes -Descending | Select-Object -First $Top

            if ($level2Folders.Count -gt 0) {
                Show-SingleTable -Results $level2Folders -Title "Level 2 Subfolders"
                $largestLevel2 = $level2Folders[0]

                # Level 3: Show subfolders of the largest Level 2 folder
                if ($largestLevel2) {
                    Write-Output "`n$($Script:Colors.Bold)$($Script:Colors.Cyan)LEVEL 3: TOP SUBFOLDERS OF $($largestLevel2.Path)$($Script:Colors.Reset)"
                    $level3Folders = $Results | Where-Object {
                        $_.IsAccessible -and
                        $_.Path -ne $largestLevel2.Path -and
                        (Split-Path -Path $_.Path -Parent) -eq $largestLevel2.Path
                    } | Sort-Object SizeBytes -Descending | Select-Object -First $Top

                    if ($level3Folders.Count -gt 0) {
                        Show-SingleTable -Results $level3Folders -Title "Level 3 Subfolders"
                    } else {
                        Write-Output "$($Script:Colors.Yellow)No accessible subfolders found at this level.$($Script:Colors.Reset)"
                    }
                }
            } else {
                Write-Output "$($Script:Colors.Yellow)No accessible subfolders found at this level.$($Script:Colors.Reset)"
            }
        }
    } else {
        Write-Output "$($Script:Colors.Yellow)No accessible top-level subfolders found.$($Script:Colors.Reset)"
    }

    # Show safely scanned problematic directories
    if ($SafeResults.Count -gt 0) {
        Write-Output "`n$($Script:Colors.Bold)$($Script:Colors.Cyan)SAFELY SCANNED PROBLEMATIC DIRECTORIES$($Script:Colors.Reset)"
        Show-SingleTable -Results $SafeResults -Title "Safely Scanned Directories"
    }

    # Show summary
    Show-HierarchicalSummary -Results $Results -SafeResults $SafeResults
}

function Show-SingleTable {
    param(
        [array]$Results,
        [string]$Title
    )

    if ($Results.Count -eq 0) {
        Write-Output "$($Script:Colors.Yellow)No data to display for $Title.$($Script:Colors.Reset)"
        return
    }

    # Calculate maximum width for alignment with null safety
    $pathLengths = $Results | Where-Object { $null -ne $_.Path } | ForEach-Object { $_.Path.Length }
    $sizeLengths = $Results | Where-Object { $null -ne $_.SizeBytes } | ForEach-Object { (Format-FileSize -SizeInBytes $_.SizeBytes).Length }

    $maxPathLength = if ($pathLengths) { ($pathLengths | Measure-Object -Maximum).Maximum } else { 20 }
    $maxSizeLength = if ($sizeLengths) { ($sizeLengths | Measure-Object -Maximum).Maximum } else { 10 }

    # Ensure minimum widths
    $maxPathLength = [Math]::Max($maxPathLength, 15)
    $maxSizeLength = [Math]::Max($maxSizeLength, 8)

    # Header
    $headerFormat = "{0,-$maxPathLength} {1,$maxSizeLength} {2,8} {3,10} {4,10} {5,-60}"
    Write-Output ($headerFormat -f "FOLDER PATH", "SIZE", "FILES", "SUBFOLDERS", "CLOUD", "LARGEST FILE (SIZE)")
    Write-Output ($headerFormat -f ("-" * $maxPathLength), ("-" * $maxSizeLength), "--------", "----------", "----------", ("-" * 60))

    # Data rows with conditional coloring
    foreach ($folder in $Results) {
        $sizeFormatted = Format-FileSize -SizeInBytes $folder.SizeBytes
        $cloudIndicator = if ($folder.HasCloudFiles) { "Yes" } else { "No" }

        $largestFile = if ($folder.LargestFile) {
            $fileName = Split-Path -Path $folder.LargestFile -Leaf
            $fileSize = Format-FileSize -SizeInBytes $folder.LargestFileSize
            $displayName = if ($fileName.Length -gt 35) {
                $fileName.Substring(0, 32) + "..."
            } else {
                $fileName
            }
            "$displayName ($fileSize)"
        } else {
            "N/A"
        }

        # Color coding based on size
        $color = if ($folder.SizeBytes -gt 10GB) {
            $Script:Colors.Red
        } elseif ($folder.SizeBytes -gt 1GB) {
            $Script:Colors.Yellow
        } else {
            $Script:Colors.Green
        }

        $rowData = $headerFormat -f $folder.Path, $sizeFormatted, $folder.FileCount, $folder.SubfolderCount, $cloudIndicator, $largestFile
        Write-Output "${color}${rowData}$($Script:Colors.Reset)"
    }
}

function Show-HierarchicalSummary {
    param(
        [array]$Results,
        [array]$SafeResults = @()
    )

    # Calculate totals from accessible results
    $accessibleResults = $Results | Where-Object { $_.IsAccessible }
    $totalSize = ($accessibleResults | Measure-Object -Property SizeBytes -Sum).Sum
    $totalFiles = ($accessibleResults | Measure-Object -Property FileCount -Sum).Sum
    $totalFolders = ($accessibleResults | Measure-Object -Property SubfolderCount -Sum).Sum
    $inaccessibleCount = ($Results | Where-Object { -not $_.IsAccessible }).Count

    # Add safe results to totals
    if ($SafeResults.Count -gt 0) {
        $safeSize = ($SafeResults | Measure-Object -Property SizeBytes -Sum).Sum
        $safeFiles = ($SafeResults | Measure-Object -Property FileCount -Sum).Sum
        $safeFolders = ($SafeResults | Measure-Object -Property SubfolderCount -Sum).Sum
        $totalSize += $safeSize
        $totalFiles += $safeFiles
        $totalFolders += $safeFolders
    }

    Write-Output "`n$($Script:Colors.Bold)COMPREHENSIVE SUMMARY STATISTICS$($Script:Colors.Reset)"
    Write-Output "$($Script:Colors.White)Total Size Analyzed: $($Script:Colors.Green)$(Format-FileSize -SizeInBytes $totalSize)$($Script:Colors.Reset)"
    Write-Output "$($Script:Colors.White)Total Files: $($Script:Colors.Green)$($totalFiles.ToString('N0'))$($Script:Colors.Reset)"
    Write-Output "$($Script:Colors.White)Total Folders: $($Script:Colors.Green)$($totalFolders.ToString('N0'))$($Script:Colors.Reset)"
    Write-Output "$($Script:Colors.White)Processed Folders: $($Script:Colors.Green)$($Script:ProcessedFolders.ToString('N0'))$($Script:Colors.Reset)"

    if ($SafeResults.Count -gt 0) {
        Write-Output "$($Script:Colors.White)Safely Scanned Directories: $($Script:Colors.Cyan)$($SafeResults.Count.ToString('N0'))$($Script:Colors.Reset)"
    }

    if ($inaccessibleCount -gt 0) {
        Write-Output "$($Script:Colors.White)Inaccessible Folders: $($Script:Colors.Red)$($inaccessibleCount.ToString('N0'))$($Script:Colors.Reset)"
    }

    # Performance metrics
    $elapsedTime = (Get-Date) - $Script:StartTime
    Write-Output "$($Script:Colors.White)Analysis Duration: $($Script:Colors.Cyan)$($elapsedTime.ToString('hh\:mm\:ss'))$($Script:Colors.Reset)"
    Write-Output "$($Script:Colors.White)Processing Rate: $($Script:Colors.Cyan)$([math]::Round($Script:ProcessedFolders / $elapsedTime.TotalSeconds, 2)) folders/second$($Script:Colors.Reset)"
}

function Show-ErrorSummary {
    if ($Script:ErrorTracker.Count -gt 0) {
        Write-Output "`n$($Script:Colors.Bold)ERROR SUMMARY$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.Red)The following $($Script:ErrorTracker.Count) folders could not be accessed:$($Script:Colors.Reset)"

        $Script:ErrorTracker.GetEnumerator() | Sort-Object Name | ForEach-Object {
            Write-Output "$($Script:Colors.DarkGray)  $($_.Key): $($_.Value)$($Script:Colors.Reset)"
        }

        if (-not $Script:IsAdmin) {
            Write-Output "`n$($Script:Colors.Yellow)Note: Running with administrative privileges may provide access to additional folders.$($Script:Colors.Reset)"
        }
    }
}

function Test-TranscriptRunning {
    <#
    .SYNOPSIS
    Tests if a PowerShell transcript is currently running.
    #>
    try {
        # Try to get transcript status - this will fail if no transcript is running
        $null = Get-Variable -Name "PSTranscript*" -ErrorAction Stop 2>$null
        return $true
    }
    catch {
        # Alternative method: try to stop transcript
        try {
            $originalErrorAction = $ErrorActionPreference
            $ErrorActionPreference = 'Stop'
            Stop-Transcript 2>$null
            # If we get here, transcript was running
            return $true
        }
        catch {
            # No transcript was running
            return $false
        }
        finally {
            $ErrorActionPreference = $originalErrorAction
        }
    }
}

function Get-ProblematicDirectorySize {
    <#
    .SYNOPSIS
    Safely gets the size of problematic directories without deep recursion to avoid hangs.
    #>
    param(
        [string]$DirectoryPath,
        [string]$DirectoryName
    )    $result = [PSCustomObject]@{
        Path = $DirectoryPath
        Name = $DirectoryName
        SizeBytes = 0
        FileCount = 0
        SubfolderCount = 0
        LargestFile = $null
        LargestFileSize = 0
        IsAccessible = $false
        HasCloudFiles = $false
        Error = "Problematic Directory - Size estimate only"
    }

    try {
        if (Test-Path -Path $DirectoryPath -PathType Container) {
            # For system directories, use shallow scanning only (no recursion)
            $items = Get-ChildItem -Path $DirectoryPath -Force -ErrorAction SilentlyContinue

            if ($items) {
                foreach ($item in $items) {
                    if ($item.PSIsContainer) {
                        $result.SubfolderCount++
                        # For subdirectories in problematic dirs, just count them without recursing
                        try {
                            $subItems = Get-ChildItem -Path $item.FullName -File -Force -ErrorAction SilentlyContinue
                            if ($subItems) {
                                $result.FileCount += $subItems.Count
                                $subItemsSum = ($subItems | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                                $result.SizeBytes += $subItemsSum

                                # Track largest file in subdirectories
                                $largestSubItem = $subItems | Sort-Object Length -Descending | Select-Object -First 1
                                if ($largestSubItem -and $largestSubItem.Length -gt $result.LargestFileSize) {
                                    $result.LargestFile = $largestSubItem.FullName
                                    $result.LargestFileSize = $largestSubItem.Length
                                }
                            }                        }
                        catch {
                            Write-Debug "Skipping inaccessible subdirectory: $($item.FullName)"
                        }
                    }
                    else {
                        $result.FileCount++
                        $fileSize = $item.Length
                        $result.SizeBytes += $fileSize

                        # Track largest file at root level
                        if ($fileSize -gt $result.LargestFileSize) {
                            $result.LargestFile = $item.FullName
                            $result.LargestFileSize = $fileSize
                        }
                    }
                }
                $result.IsAccessible = $true
                $result.Error = "Shallow scan only (problematic directory)"
            }
        }
    }
    catch {
        $result.Error = "Access Denied: $($_.Exception.Message)"
    }

    return $result
}

# MAIN SCRIPT EXECUTION
function Main {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$StartPath,
        [int]$MaxDepth,
        [int]$Top,
        [int]$MaxThreads
    )    $scriptPath = if ($MyInvocation.MyCommand.Path) {
        Split-Path -Parent $MyInvocation.MyCommand.Path
    } else {
        Split-Path -Parent $PSCommandPath
    }
    $logPath = Start-AdvancedTranscript -LogPath $scriptPath

    try {
        # Initialize and display header        Write-Output "$($Script:Colors.Bold)$($Script:Colors.Cyan)===============================================$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.Bold)$($Script:Colors.Cyan)    ULTRA-FAST FOLDER USAGE ANALYZER v1.5.0    $($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.Bold)$($Script:Colors.Cyan)===============================================$($Script:Colors.Reset)"

        # Administrative privilege check
        $Script:IsAdmin = Test-AdminPrivilege
        $adminStatus = if ($Script:IsAdmin) {
            "$($Script:Colors.Green)Administrator$($Script:Colors.Reset)"
        } else {
            "$($Script:Colors.Yellow)Standard User$($Script:Colors.Reset)"
        }
        Write-Output "$($Script:Colors.White)Running as: $adminStatus$($Script:Colors.Reset)"

        # Display configuration
        Write-Output "$($Script:Colors.White)Start Path: $($Script:Colors.Cyan)$StartPath$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Max Depth: $($Script:Colors.Cyan)$MaxDepth$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Top Folders: $($Script:Colors.Cyan)$Top$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Max Threads: $($Script:Colors.Cyan)$MaxThreads$($Script:Colors.Reset)"
        Write-Output ""

        if ($PSCmdlet.ShouldProcess($StartPath, "Analyze folder usage")) {            # Get top-level directories for parallel processing
            Write-ProgressUpdate -Activity "Initialization" -Status "Scanning top-level directories..." -Color "Cyan"

            # Known problematic directories that can cause hangs
            $problematicDirs = @(
                "System Volume Information",
                "`$Recycle.Bin",
                "Recovery",
                "Documents and Settings",                "`$WinREAgent",
                "hiberfil.sys",
                "pagefile.sys",                "swapfile.sys"
            )

            $topLevelDirs = @()
            $problematicDirsFound = @()

            try {
                $items = Get-ChildItem -Path $StartPath -Directory -Force -ErrorAction Stop
                  # Separate normal directories from problematic ones
                foreach ($item in $items) {
                    $dirName = $item.Name
                    # Use exact matching instead of wildcard matching to prevent false positives
                    $isProblematic = $problematicDirs -contains $dirName

                    if ($isProblematic) {
                        $problematicDirsFound += $item
                    } else {
                        $topLevelDirs += $item.FullName
                    }
                }

                $Script:TotalFolders = $topLevelDirs.Count

                Write-Output "$($Script:Colors.Green)Found $($topLevelDirs.Count) top-level directories to analyze$($Script:Colors.Reset)"
                if ($problematicDirsFound.Count -gt 0) {
                    Write-Output "$($Script:Colors.Yellow)Found $($problematicDirsFound.Count) problematic directories - will analyze safely$($Script:Colors.Reset)"
                }
            }
            catch {
                Write-Output "$($Script:Colors.Red)Error accessing start path: $($_.Exception.Message)$($Script:Colors.Reset)"
                return
            }

            if ($topLevelDirs.Count -eq 0) {
                Write-Output "$($Script:Colors.Yellow)No directories found to analyze.$($Script:Colors.Reset)"
                return
            }            # Perform parallel analysis
            Write-ProgressUpdate -Activity "Analysis" -Status "Starting parallel folder analysis..." -Color "Cyan"
            $results = Get-ParallelFolderStatistic -FolderPaths $topLevelDirs -MaxDepth $MaxDepth -MaxThreads $MaxThreads

            # Analyze the StartPath directory itself for files at the root level
            Write-ProgressUpdate -Activity "Analysis" -Status "Analyzing root directory files..." -Color "Cyan"
            try {
                $rootFiles = Get-ChildItem -Path $StartPath -File -Force -ErrorAction Stop
                if ($rootFiles.Count -gt 0) {
                    $rootStats = [PSCustomObject]@{
                        Path = $StartPath
                        SizeBytes = ($rootFiles | Measure-Object -Property Length -Sum).Sum
                        FileCount = $rootFiles.Count
                        SubfolderCount = $topLevelDirs.Count + $problematicDirsFound.Count
                        LargestFile = ($rootFiles | Sort-Object Length -Descending | Select-Object -First 1).FullName
                        LargestFileSize = ($rootFiles | Sort-Object Length -Descending | Select-Object -First 1).Length
                        IsAccessible = $true
                        HasCloudFiles = ($rootFiles | Where-Object { Test-CloudPlaceholder -FileInfo $_ }).Count -gt 0
                        Error = $null
                    }
                    $results += $rootStats
                    $Script:ProcessedFolders++
                }
            }
            catch {
                Write-Output "$($Script:Colors.Yellow)Warning: Could not analyze root directory files: $($_.Exception.Message)$($Script:Colors.Reset)"
            }            # Safely analyze problematic directories
            $safeResults = @()
            if ($problematicDirsFound.Count -gt 0) {
                Write-ProgressUpdate -Activity "Analysis" -Status "Analyzing problematic directories safely..." -Color "Yellow"
                foreach ($problematicDir in $problematicDirsFound) {
                    Write-Output "$($Script:Colors.DarkGray)Safely scanning: $($problematicDir.FullName)$($Script:Colors.Reset)"
                    $problematicResult = Get-ProblematicDirectorySize -DirectoryPath $problematicDir.FullName -DirectoryName $problematicDir.Name
                    $safeResults += $problematicResult
                    $Script:ProcessedFolders++  # Count problematic directories as processed
                }
            }

            # Display results using hierarchical view
            Show-HierarchicalResult -Results $results -StartPath $StartPath -Top $Top -SafeResults $safeResults
            Show-ErrorSummary

            Write-Output "`n$($Script:Colors.Green)Analysis completed successfully!$($Script:Colors.Reset)"
        }
        else {
            Write-Output "$($Script:Colors.Yellow)WhatIf: Would analyze folder usage starting from '$StartPath'$($Script:Colors.Reset)"
            Write-Output "$($Script:Colors.Yellow)WhatIf: Would scan up to $MaxDepth levels deep$($Script:Colors.Reset)"
            Write-Output "$($Script:Colors.Yellow)WhatIf: Would display top $Top largest folders$($Script:Colors.Reset)"
            Write-Output "$($Script:Colors.Yellow)WhatIf: Would use $MaxThreads parallel threads$($Script:Colors.Reset)"        }
    }
    catch {
        Write-Output "$($Script:Colors.Red)Fatal error: $($_.Exception.Message)$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.Red)Stack trace: $($_.ScriptStackTrace)$($Script:Colors.Reset)"
    }
    finally {
        # Cleanup and finalization - ALWAYS executed
        Write-ProgressUpdate -Activity "Analysis" -Status "Complete" -Completed

        # Enhanced transcript cleanup with multiple fallback methods
        if (Test-TranscriptRunning) {
            Write-Output "$($Script:Colors.Cyan)Stopping active transcript...$($Script:Colors.Reset)"
            try {
                if ($logPath -and ($logPath -is [string])) {
                    Stop-AdvancedTranscript -LogPath ([string]$logPath)
                } else {
                    # Force stop any transcript even without path
                    Stop-AdvancedTranscript -LogPath ""
                }
            }
            catch {
                # Ultimate fallback - force stop without our function
                try {
                    Stop-Transcript -ErrorAction SilentlyContinue
                    Write-Output "$($Script:Colors.Yellow)Transcript stopped using emergency fallback$($Script:Colors.Reset)"
                }
                catch {
                    Write-Output "$($Script:Colors.Red)Warning: Could not stop transcript properly$($Script:Colors.Reset)"
                }
            }
        } else {
            Write-Output "$($Script:Colors.Green)No active transcript to clean up$($Script:Colors.Reset)"
        }

        # Final memory cleanup
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# Execute main function
if ($MyInvocation.InvocationName -ne '.') {
    Main -StartPath $StartPath -MaxDepth $MaxDepth -Top $Top -MaxThreads $MaxThreads -WhatIf:$WhatIfPreference
}
