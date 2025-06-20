# =============================================================================
# Script: Analyze-FolderUsage.ps1
# Created: 2025-06-13 20:57:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-06-20 21:47:00 UTC
# Updated By: jdyer-nuvodia
# Version: 5.4.1
# Additional Info: Fixed parallel analysis progress bar to show completed jobs out of total work items (launched + queued) for accurate progress reporting.
# =============================================================================

<#
.SYNOPSIS
Performs ultra-fast recursive folder size analysis with parallel processing and advanced features.
The script is primarily designed to be run against the root of a drive. It can be used on any directory, but is optimized for large-scale analysis starting from a drive root.
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
- Multi-level progress reporting with version-aware color coding (ANSI in PS 7+, plain text in PS 5.1)
- Advanced transcript logging with cleanup
- Drive information display for the target drive
The script uses PowerShell runspaces for maximum performance and includes memory management with garbage collection.
Supports Windows PowerShell 5.1 and later with modern .NET integration and version-aware color compatibility.
Drive information is automatically displayed for the drive containing the StartPath to provide context on available space.
.PARAMETER StartPath
The root directory path to begin analysis. Default is "C:\"
Type: String
.PARAMETER MaxDepth
Maximum recursion depth for directory scanning. Default is 15 levels.
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
.PARAMETER Debug
Uses the built-in PowerShell Debug parameter. Enables detailed debug logging for troubleshooting. Shows all major processing steps, directory counts, and calculation details.
Type: Switch (Built-in PowerShell parameter)
.PARAMETER Verbose
Uses the built-in PowerShell Verbose parameter. Enables detailed progress reporting and status updates during execution.
Type: Switch (Built-in PowerShell parameter)
.EXAMPLE
.\Analyze-FolderUsage.ps1
Analyzes the C:\ drive with default settings (max depth 15, top 10 folders, 10 threads).
.EXAMPLE
.\Analyze-FolderUsage.ps1 -StartPath "D:\Data" -MaxDepth 5 -Top 5 -MaxThreads 20
Analyzes D:\Data with custom depth limit of 5 levels, showing top 5 largest folders using 20 threads.
.EXAMPLE
.\Analyze-FolderUsage.ps1 -StartPath "C:\Users" -Debug -Verbose
Analyzes C:\Users with detailed debug logging and verbose progress reporting using the built-in Verbose parameter.
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
    [int]$MaxDepth = 15,
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 100)]
    [int]$Top = 10,
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 50)]
    [int]$MaxThreads = 10
)
# PowerShell Version-Aware Color System
# PowerShell 5.1 does not support ANSI escape sequences, while PowerShell 7+ does
# This ensures compatibility across versions without using Write-Host
if ($PSVersionTable.PSVersion.Major -ge 7) {
    # PowerShell 7+ supports ANSI escape sequences
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
} else {
    # PowerShell 5.1 and earlier - use empty strings to disable colors
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
        Underline  = ""
    }
}
# Script Variables for Script Operation
$Script:ErrorTracker = @{}
$Script:IsAdmin = $false
$Script:ProcessedFolders = 0
$Script:TotalFolders = 0
$Script:InaccessibleFolderCount = 0
$Script:StartTime = Get-Date
$Script:MaxDepthReached = $false
# Centralized logging variables
$Script:CentralLogPath = $null
$Script:LogMutex = $null
$Script:EnableCentralLogging = $false
# Debug and Verbose Logging Functions
function Initialize-CentralLogging {
    <#
    .SYNOPSIS
    Initializes centralized logging for debug messages only when -Debug is specified.
    #>
    [CmdletBinding()]
    param()
    # Only initialize central logging if Debug preference is enabled
    if ($DebugPreference -eq 'SilentlyContinue') {
        Write-Verbose "Debug mode not enabled - skipping central debug log initialization"
        $Script:EnableCentralLogging = $false
        return $null
    }
    try {
        # Create central log file path
        $computerName = $env:COMPUTERNAME
        $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $logFileName = "Analyze-FolderUsage_CENTRAL_${computerName}_${timestamp}.log"
        $Script:CentralLogPath = Join-Path -Path $PSScriptRoot -ChildPath $logFileName
        # Initialize mutex for thread safety
        $Script:LogMutex = New-Object System.Threading.Mutex($false, "AnalyzeFolderUsageCentralLog")
        # Enable central logging
        $Script:EnableCentralLogging = $true
        # Create initial log entry
        $headerText = @"
===============================================
CENTRAL DEBUG LOG - FOLDER USAGE ANALYZER v3.0.0
===============================================
Log started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC')
Computer: $computerName
PowerShell Version: $($PSVersionTable.PSVersion)
Process ID: $PID
===============================================
"@
        Set-Content -Path $Script:CentralLogPath -Value $headerText -Encoding UTF8 -ErrorAction Stop
        Write-Output "$($Script:Colors.Cyan)Central debug logging initialized: $Script:CentralLogPath$($Script:Colors.Reset)"
        return $Script:CentralLogPath
    } catch {
        Write-Warning "Failed to initialize central logging: $($_.Exception.Message)"
        $Script:EnableCentralLogging = $false
        return $null
    }
}
function Write-CentralLog {
    <#
    .SYNOPSIS
    Thread-safe centralized logging function that captures DEBUG output only when -Debug is specified.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$Category = "INFO",
        [Parameter(Mandatory = $false)]
        [string]$Source = "MAIN"
    )
    # Only log to central log if debug mode is enabled and central logging is initialized
    if (-not $Script:EnableCentralLogging -or [string]::IsNullOrEmpty($Script:CentralLogPath) -or $DebugPreference -eq 'SilentlyContinue') {
        return
    }
    try {
        # Create timestamp
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
        $logEntry = "[$timestamp] [$Source] [$Category] $Message"
        # Write to log file only - console output should come from main script logic
        # Logging functions should not output to streams that can be captured by functions
        # Thread-safe file writing
        $acquired = $false
        try {
            if ($Script:LogMutex) {
                $acquired = $Script:LogMutex.WaitOne(5000)
            }
            if ($acquired -or -not $Script:LogMutex) {
                # Write to log file
                Add-Content -Path $Script:CentralLogPath -Value $logEntry -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        } catch {
            # Silently continue if logging fails to avoid disrupting main functionality
            Write-Verbose "Central logging failed: $($_.Exception.Message)"
        }
        finally {
            if ($acquired -and $Script:LogMutex) {
                $Script:LogMutex.ReleaseMutex()
            }
        }
    } catch {
        # Silently continue if logging fails
        Write-Verbose "Central logging initialization failed: $($_.Exception.Message)"
    }
}
function Write-DebugInfo {
    [CmdletBinding()]
    param(
        [string]$Message,
        [string]$Category = "DEBUG"
    )
    if ($DebugPreference -ne 'SilentlyContinue') {
        $timestamp = Get-Date -Format "HH:mm:ss.fff"
        $formattedMessage = "[$timestamp] [$Category] $Message"
        Write-Output "$($Script:Colors.Magenta)$formattedMessage$($Script:Colors.Reset)"
        # Also write to central log
        Write-CentralLog -Message $Message -Category $Category -Source "MAIN"
    }
}
function Write-VerboseInfo {
    [CmdletBinding()]
    param(
        [string]$Message,
        [string]$Category = "VERBOSE"
    )
    # Use PowerShell's built-in Write-Verbose which will output to transcript when -Verbose is used
    if ($VerbosePreference -ne 'SilentlyContinue') {
        $timestamp = Get-Date -Format "HH:mm:ss.fff"
        $formattedMessage = "[$timestamp] [$Category] $Message"
        Write-Verbose $formattedMessage
    }
    # Also show verbose info when debug is enabled
    if ($DebugPreference -ne 'SilentlyContinue') {
        $timestamp = Get-Date -Format "HH:mm:ss.fff"
        Write-Output "$($Script:Colors.DarkGray)[$timestamp] [$Category] $Message$($Script:Colors.Reset)"
    }
}
function Write-DetailedProgress {
    [CmdletBinding()
    ]
    param(
        [string]$Activity,
        [string]$Status,
        [string]$Detail = "",
        [string]$Color = "Cyan",
        [int]$PercentComplete = -1
    )
    Write-ProgressUpdate -Activity $Activity -Status $Status -Color $Color -PercentComplete $PercentComplete
    if ($VerbosePreference -ne 'SilentlyContinue' -or $DebugPreference -ne 'SilentlyContinue') {
        $message = if ($Detail) { "$Status - $Detail" } else { $Status }
        Write-VerboseInfo -Message $message -Category "PROGRESS"
    }
}
# Advanced Transcript Management Functions
function Start-AdvancedTranscript {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$LogPath
    )
    try {
        # Validate and normalize the log path
        if ([string]::IsNullOrWhiteSpace($LogPath)) {
            $LogPath = $PWD.Path
        }
        $computerName = $env:COMPUTERNAME
        $timestamp = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")
        $logFileName = "Analyze-FolderUsage_TRANSCRIPT_${computerName}_${timestamp}.log"
        $fullLogPath = Join-Path -Path $LogPath -ChildPath $logFileName
        # Ensure log directory exists
        if (-not (Test-Path -Path $LogPath)) {
            if ($PSCmdlet.ShouldProcess($LogPath, "Create log directory")) {
                New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
            }
        }
        # Create initial log file with header
        $headerText = @"
===============================================
ULTRA-FAST FOLDER USAGE ANALYZER v5.4.0 TRANSCRIPT LOG
===============================================
Log started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC')
Computer: $computerName
PowerShell Version: $($PSVersionTable.PSVersion)
Process ID: $PID
===============================================
"@
        Set-Content -Path $fullLogPath -Value $headerText -Encoding UTF8 -ErrorAction SilentlyContinue
        if ($PSCmdlet.ShouldProcess($fullLogPath, "Start transcript")) {
            Start-Transcript -Path $fullLogPath -Append -Force | Out-Null
            Write-Output "$($Script:Colors.Green)Advanced transcript started: $fullLogPath$($Script:Colors.Reset)"
            Write-Verbose "Transcript logging initialized for all verbose output"
        }
        return $fullLogPath
    } catch {
        Write-Output "$($Script:Colors.Yellow)Warning: Could not start transcript: $($_.Exception.Message). Continuing without logging.$($Script:Colors.Reset)"
        return $null
    }
}
function Clear-CentralLogging {
    <#
    .SYNOPSIS    Cleans up central logging resources and finalizes the log file.
    #>
    [CmdletBinding()]
    param()
    try {
        if ($Script:EnableCentralLogging -and $Script:CentralLogPath) {
            # Write final log entry
            $finalEntry = @"
===============================================
LOG SESSION COMPLETED
===============================================
End time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC')
Total duration: $((Get-Date).Subtract($Script:StartTime).ToString())
===============================================
"@
Add-Content -Path $Script:CentralLogPath -Value $finalEntry -Encoding UTF8 -ErrorAction SilentlyContinue
            Write-Output "$($Script:Colors.Cyan)Central debug log finalized: $Script:CentralLogPath$($Script:Colors.Reset)"
        }
        # Clean up mutex
        if ($Script:LogMutex) {
            try {
                $Script:LogMutex.Dispose()
            } catch {
                Write-Error "Failed to dispose logging mutex: $($_.Exception.Message)" -ErrorAction SilentlyContinue
            }
            $Script:LogMutex = $null
        }
        # Disable logging
        $Script:EnableCentralLogging = $false
    } catch {
        Write-Verbose "Error during central logging cleanup: $($_.Exception.Message)"
    }
}
function Stop-AdvancedTranscript {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$LogPath
    )
    try {
        if ($PSCmdlet.ShouldProcess("Transcript", "Stop transcript")) {
            # Write final log entry
            Write-Verbose "Stopping transcript"
            # First try normal stop
            Stop-Transcript -ErrorAction Stop
            Write-Output "$($Script:Colors.Green)Transcript stopped successfully$($Script:Colors.Reset)"
            if ($LogPath -and (Test-Path -Path $LogPath)) {
                Write-Output "$($Script:Colors.Green)Transcript saved: $LogPath$($Script:Colors.Reset)"
            }
        }
    } catch {
        # Multiple fallback methods for transcript cleanup
        Write-Output "$($Script:Colors.Yellow)Warning: Normal transcript stop failed. Attempting cleanup methods...$($Script:Colors.Reset)"
        try {
            # Method 1: Force stop with error suppression
            Stop-Transcript -ErrorAction SilentlyContinue
            Write-Output "$($Script:Colors.Green)Transcript stopped using fallback method 1$($Script:Colors.Reset)"
        } catch {
            try {
                # Method 2: Direct registry cleanup (Windows PowerShell specific)
                if ($PSVersionTable.PSVersion.Major -le 5) {
                    $regPath = "HKCU:\Software\Microsoft\PowerShell\1\PowerShellEngine"
                    if (Test-Path -Path $regPath) {
                        Remove-ItemProperty -Path $regPath -Name "TranscriptPath" -ErrorAction SilentlyContinue
                        Write-Output "$($Script:Colors.Green)Transcript stopped using registry cleanup$($Script:Colors.Reset)"
                    }
                }
            } catch {
                try {
                    # Method 3: Clear transcript variable directly
                    $ExecutionContext.InvokeCommand.InvokeScript('$null = Stop-Transcript') 2>$null
                    Write-Output "$($Script:Colors.Green)Transcript stopped using execution context$($Script:Colors.Reset)"
                } catch {
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
    } catch {
        return $false
    }
}
function Show-DriveInfo {
    <#
    .SYNOPSIS
    Displays detailed drive information for a specified drive letter.
    .DESCRIPTION
    Retrieves and displays comprehensive drive information including:
    - Drive letter and label
    - File system type
    - Total, used, and free space
    - Health status
    Uses PowerShell's Get-Volume cmdlet.
    .PARAMETER DriveLetter
    The drive letter to display information for (without colon).
    .EXAMPLE
    Show-DriveInfo -DriveLetter "C"
    Displays information for the C: drive
    #>
    param (
        [Parameter(Mandatory=$true)]
        [ValidatePattern('^[A-Za-z]$')]
        [string]$DriveLetter
    )
    try {
        $volume = Get-Volume -DriveLetter $DriveLetter -ErrorAction Stop
        Write-Output "`n$($Script:Colors.Bold)$($Script:Colors.Green)Drive Volume Details for ${DriveLetter}:$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.Green)------------------------$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Drive Letter: $($Script:Colors.Cyan)$($volume.DriveLetter):$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Drive Label: $($Script:Colors.Cyan)$($volume.FileSystemLabel)$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)File System: $($Script:Colors.Cyan)$($volume.FileSystem)$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Drive Type: $($Script:Colors.Cyan)$($volume.DriveType)$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Size: $($Script:Colors.Cyan)$([math]::Round($volume.Size/1GB, 2)) GB$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Free Space: $($Script:Colors.Cyan)$([math]::Round($volume.SizeRemaining/1GB, 2)) GB$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Used Space: $($Script:Colors.Cyan)$([math]::Round(($volume.Size - $volume.SizeRemaining)/1GB, 2)) GB$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Free Space %: $($Script:Colors.Cyan)$([math]::Round(($volume.SizeRemaining/$volume.Size) * 100, 2))%$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Health Status: $($Script:Colors.Cyan)$($volume.HealthStatus)$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Operational Status: $($Script:Colors.Cyan)$($volume.OperationalStatus)$($Script:Colors.Reset)"
        Write-Output ""
    } catch {
        Write-Output "$($Script:Colors.Red)Error retrieving drive information for ${DriveLetter}: $($_.Exception.Message)$($Script:Colors.Reset)"
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
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$FileInfo
    )
    try {
        # Enhanced OneDrive/cloud storage detection
        # Validate input to prevent null reference exceptions
        if ($null -eq $FileInfo) {
            Write-DebugInfo -Message "Null FileInfo passed to Test-CloudPlaceholder" -Category "CLOUD_CHECK"
            return $false
        }
        # Check for OneDrive specific file extensions and patterns
        if ($FileInfo.Extension -eq ".odlocal" -or $FileInfo.Name -like "*.onedrive") {
            return $true
        }
        # Check if it's in a known OneDrive path structure
        if ($FileInfo.FullName -like "*OneDrive*") {
            # Check for cloud storage attributes (FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS = 0x400000)
            if (($FileInfo.Attributes.value__ -band 0x400000) -eq 0x400000) {
                return $true
            }
            # Check for reparse point in OneDrive paths (common for cloud files)
            if (($FileInfo.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -eq [System.IO.FileAttributes]::ReparsePoint) {
                return $true
            }
            # Check for specific OneDrive cloud file attributes
            # OneDrive files often have Offline + ReparsePoint attributes
            if (($FileInfo.Attributes -band [System.IO.FileAttributes]::Offline) -eq [System.IO.FileAttributes]::Offline) {
                return $true
            }
        }

        # Additional check for cloud storage attributes outside OneDrive paths
        if (($FileInfo.Attributes.value__ -band 0x400000) -eq 0x400000) {
            return $true
        }
        return $false
    } catch {
        Write-DebugInfo -Message "Error in Test-CloudPlaceholder: $($_.Exception.Message)" -Category "CLOUD_CHECK"
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
    } catch {
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
        MaxDepthReached = $false
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
                    # Propagate MaxDepthReached flag
                    if ($subStats.MaxDepthReached) {
                        $stats.MaxDepthReached = $true
                        $Script:MaxDepthReached = $true
                    }
                } else {
                    # Max depth reached - set flag
                    $stats.MaxDepthReached = $true
                    $Script:MaxDepthReached = $true
                }
            }
            else {
                $stats.FileCount++
                # Handle cloud placeholder files
                if (Test-CloudPlaceholder -FileInfo $item) {
                    $stats.HasCloudFiles = $true
                    # For cloud placeholder files, use a minimal size estimation
                    # OneDrive placeholders typically show full size but aren't locally stored
                    # Use actual size or 1KB, whichever is smaller
                    $fileSize = [Math]::Min($item.Length, 1KB)
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
    } catch [System.UnauthorizedAccessException] {
        $stats.IsAccessible = $false
        $stats.Error = "Access Denied"
        $Script:ErrorTracker[$FolderPath] = "Access Denied"
        Write-DetailedProgress -Activity "Access Error" -Status "Access denied: $FolderPath" -Color "Red"
    } catch [System.IO.DirectoryNotFoundException] {
        $stats.IsAccessible = $false
        $stats.Error = "Directory Not Found"
        $Script:ErrorTracker[$FolderPath] = "Directory Not Found"
        Write-DetailedProgress -Activity "Directory Error" -Status "Directory not found: $FolderPath" -Color "Red"
    } catch {
        $stats.IsAccessible = $false
        $stats.Error = $_.Exception.Message
        $Script:ErrorTracker[$FolderPath] = $_.Exception.Message
        Write-DetailedProgress -Activity "Processing Error" -Status "Error processing $FolderPath`: $($_.Exception.Message)" -Color "Red"
    }
    return $stats
}

function Get-ParallelFolderStatistic {
    param(
        [string[]]$FolderPaths,
        [int]$MaxDepth,
        [int]$MaxThreads
    )
    # Enhanced monitoring and job management variables
    $monitoringStartTime = Get-Date
    $peakActiveJobs = 0
    $jobLaunchTimes = @{}
    $jobCompletionTimes = @{}
      # PERFORMANCE OPTIMIZATION: Use separate collections to avoid O(n^2) filtering
    $activeJobs = [System.Collections.ArrayList]::new()
    $completedJobs = [System.Collections.ArrayList]::new()
    $allResults = [System.Collections.ArrayList]::new()

    # Initialize runspace pool
    $runspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads)
    $runspacePool.Open()
    # WORK QUEUE SYSTEM - This is the key improvement
    $workQueue = [System.Collections.Concurrent.ConcurrentQueue[PSCustomObject]]::new()
    $processedPaths = [System.Collections.Concurrent.ConcurrentDictionary[string, bool]]::new()

    try {
        Write-CentralLog -Message "=== PARALLEL EXECUTION WITH WORK-STEALING QUEUE ===" -Category "PARALLEL_MONITOR" -Source "MAIN"
        Write-CentralLog -Message "Runspace pool configured with MaxThreads: $MaxThreads" -Category "PARALLEL_MONITOR" -Source "MAIN"
        Write-CentralLog -Message "Initial directories to process: $($FolderPaths.Count)" -Category "PARALLEL_MONITOR" -Source "MAIN"
        Write-CentralLog -Message "Work queue system: ENABLED for maximum thread utilization" -Category "PARALLEL_MONITOR" -Source "MAIN"

        # Initialize work queue with initial directories
        foreach ($folder in $FolderPaths) {
            $workItem = [PSCustomObject]@{
                Path = $folder
                Depth = 0
                Priority = 0
            }
            $workQueue.Enqueue($workItem)
            Write-CentralLog -Message "Queued initial work item: $folder" -Category "WORK_QUEUE" -Source "MAIN"
        }

        Write-CentralLog -Message "Work queue initialized with $($workQueue.Count) items" -Category "WORK_QUEUE" -Source "MAIN"
        # Create enhanced scriptblock with work queue integration
        $scriptBlock = {
            param(
                [string]$FolderPath,
                [int]$MaxDepth,
                [int]$CurrentDepth,
                [bool]$EnableDebug,
                [string]$CentralLogPath
            )
            # Strict output control - capture all unwanted output
            $ErrorActionPreference = 'Continue'
            $DebugPreference = 'SilentlyContinue'
            $VerbosePreference = 'SilentlyContinue'
            $InformationPreference = 'SilentlyContinue'
            $WarningPreference = 'SilentlyContinue'
            $ProgressPreference = 'SilentlyContinue'

            # Debug flag passed from main script
            $DebugEnabled = $EnableDebug
            # Make CentralLogPath available to nested functions
            $script:CentralLogPath = $CentralLogPath
            # Debug output function for runspace
            function Write-RunspaceDebug {
                param([string]$Message, [string]$Category = "RUNSPACE")
                if ($DebugEnabled -and -not [string]::IsNullOrEmpty($script:CentralLogPath)) {
                    try {
                        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
                        $logEntry = "[$timestamp] [RUNSPACE:$Category] $Message"
                        Add-Content -Path $script:CentralLogPath -Value $logEntry -Encoding UTF8 -ErrorAction SilentlyContinue
                    } catch {
                        Write-Error -Message "Failed to write to central log: $($_.Exception.Message)" -ErrorAction SilentlyContinue
                    }
                }
            }
            # Helper functions within scriptblock
            function Test-CloudPlaceholder {
                param([System.IO.FileInfo]$FileInfo)
                try {
                    if ($FileInfo.FullName -like "*OneDrive*") {
                        return ($FileInfo.Attributes.value__ -band 0x400000) -eq 0x400000 -or $FileInfo.Length -eq 0
                    }
                    if (($FileInfo.Attributes.value__ -band 0x400000) -eq 0x400000) {
                        return $true
                    }
                    return $false
                } catch {
                    return $false
                }
            }
            function Test-ReparsePoint {
                param([string]$Path)
                try {
                    $item = Get-Item -Path $Path -Force -ErrorAction Stop
                    return ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -eq [System.IO.FileAttributes]::ReparsePoint
                } catch {
                    return $false
                }
            }

            # OPTIMIZED SINGLE-DIRECTORY SCAN - No deep recursion, returns discovered subdirectories
            function Get-FolderStatisticSingle {
                param(
                    [string]$FolderPath,
                    [int]$CurrentDepth,
                    [int]$MaxDepth
                )
                Write-RunspaceDebug -Message "Scanning single directory: $FolderPath (Depth: $CurrentDepth/$MaxDepth)" -Category "FOLDER_SCAN"

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
                    MaxDepthReached = $false
                    DiscoveredSubdirectories = @()
                }

                try {
                    if (Test-ReparsePoint -Path $FolderPath) {
                        Write-RunspaceDebug -Message "Skipping reparse point: $FolderPath" -Category "FOLDER_SCAN"
                        $stats.IsAccessible = $false
                        $stats.Error = "Reparse Point"
                        return $stats
                    }

                    $items = Get-ChildItem -Path $FolderPath -Force -ErrorAction Stop
                    $itemCount = 0

                    foreach ($item in $items) {
                        $itemCount++
                        if ($item.PSIsContainer) {
                            # Directory processing
                            $stats.SubfolderCount++

                            # Add to discovered subdirectories if we can go deeper
                            if ($CurrentDepth -lt $MaxDepth) {
                                $stats.DiscoveredSubdirectories += $item.FullName
                            } else {
                                $stats.MaxDepthReached = $true
                                Write-RunspaceDebug -Message "Max depth reached, discovered but not queuing: $($item.FullName)" -Category "FOLDER_SCAN"
                            }
                        } else {
                            # File processing
                            $stats.FileCount++
                            if (Test-CloudPlaceholder -FileInfo $item) {
                                $stats.HasCloudFiles = $true
                                $fileSize = [Math]::Min($item.Length, 1KB)
                            } else {
                                $fileSize = $item.Length
                            }
                            $stats.SizeBytes += $fileSize
                            if ($fileSize -gt $stats.LargestFileSize) {
                                $stats.LargestFile = $item.FullName
                                $stats.LargestFileSize = $fileSize
                            }
                        }
                        # Throttle for large directories
                        if ($itemCount % 1000 -eq 0) {
                            Start-Sleep -Milliseconds 5
                        }
                    }

                    Write-RunspaceDebug -Message "SUCCESS: Scanned $FolderPath - found $($stats.FileCount) files, $($stats.SubfolderCount) folders, discovered $($stats.DiscoveredSubdirectories.Count) subdirectories" -Category "FOLDER_SCAN"

                } catch [System.UnauthorizedAccessException] {
                    $stats.IsAccessible = $false
                    $stats.Error = "Access Denied"
                    Write-RunspaceDebug -Message "ACCESS DENIED: $FolderPath" -Category "FOLDER_ERROR"
                } catch [System.IO.DirectoryNotFoundException] {
                    $stats.IsAccessible = $false
                    $stats.Error = "Directory Not Found"
                    Write-RunspaceDebug -Message "DIRECTORY NOT FOUND: $FolderPath" -Category "FOLDER_ERROR"
                } catch {
                    $stats.IsAccessible = $false
                    $stats.Error = $_.Exception.Message
                    Write-RunspaceDebug -Message "ERROR: Failed to access $FolderPath - $($_.Exception.Message)" -Category "FOLDER_ERROR"
                }

                return $stats
            }

            # Main execution within runspace
            try {
                Write-RunspaceDebug -Message "Starting optimized folder analysis for: $FolderPath at depth $CurrentDepth" -Category "RUNSPACE_START"
                $result = Get-FolderStatisticSingle -FolderPath $FolderPath -CurrentDepth $CurrentDepth -MaxDepth $MaxDepth
                Write-RunspaceDebug -Message "Completed folder analysis for: $FolderPath - discovered $($result.DiscoveredSubdirectories.Count) subdirectories" -Category "RUNSPACE_COMPLETE"
                return $result
            } catch {
                Write-RunspaceDebug -Message "CRITICAL ERROR in runspace for $FolderPath - $($_.Exception.Message)" -Category "RUNSPACE_ERROR"
                # Return minimal error result
                return [PSCustomObject]@{
                    Path = $FolderPath
                    SizeBytes = 0
                    FileCount = 0
                    SubfolderCount = 0
                    LargestFile = $null
                    LargestFileSize = 0
                    IsAccessible = $false
                    HasCloudFiles = $false
                    Error = "Runspace execution error: $($_.Exception.Message)"
                    MaxDepthReached = $false
                    DiscoveredSubdirectories = @()
                }
            }
        }
        # DYNAMIC JOB MANAGEMENT WITH WORK-STEALING QUEUE
        Write-CentralLog -Message "Starting dynamic job management system..." -Category "PARALLEL_MONITOR" -Source "MAIN"
        $individualJobTimeout = 180
        $completedJobsCount = 0
        $totalJobsLaunched = 0
        $timeout = (Get-Date).AddMinutes(30)
        $lastProgress = Get-Date
        $lastMonitoringReport = Get-Date
        $progressCheckInterval = 500
        $monitoringReportInterval = 5000
        Write-CentralLog -Message "Work-stealing queue system active with $($workQueue.Count) initial items" -Category "PARALLEL_MONITOR" -Source "MAIN"
        # MAIN WORK-STEALING EXECUTION LOOP - OPTIMIZED FOR PERFORMANCE AND THREAD SATURATION
        while (($workQueue.Count -gt 0 -or $activeJobs.Count -gt 0) -and (Get-Date) -lt $timeout) {
            $currentTime = Get-Date

            # PERFORMANCE: Direct access to active jobs count - no filtering needed!
            $currentActiveCount = $activeJobs.Count
            $availableThreads = $MaxThreads - $currentActiveCount
            # ENHANCED JOB LAUNCHING - Respect MaxThreads while maintaining aggressive saturation
            if ($availableThreads -gt 0 -and $workQueue.Count -gt 0) {
                # AGGRESSIVE BUT CONTROLLED: Always respect MaxThreads limit for system stability
                # Launch jobs to fill ALL available threads immediately for maximum utilization
                # The key to performance is instant job replacement, not exceeding thread limits
                $jobsToLaunch = [Math]::Min($availableThreads, $workQueue.Count)

                # Enhanced logging based on queue size to show saturation strategy
                if ($workQueue.Count -gt ($MaxThreads * 10)) {
                    Write-CentralLog -Message "LARGE QUEUE MODE: Queue has $($workQueue.Count) items, launching $jobsToLaunch jobs to saturate all $availableThreads available threads" -Category "THREAD_SATURATION" -Source "MAIN"
                } elseif ($workQueue.Count -gt ($MaxThreads * 3)) {
                    Write-CentralLog -Message "ACTIVE QUEUE MODE: Queue has $($workQueue.Count) items, launching $jobsToLaunch jobs to fill $availableThreads threads" -Category "THREAD_SATURATION" -Source "MAIN"
                } else {
                    Write-CentralLog -Message "SATURATION: Launching $jobsToLaunch jobs to fill $availableThreads available threads (Queue: $($workQueue.Count))" -Category "THREAD_SATURATION" -Source "MAIN"
                }

                $jobsLaunchedThisIteration = 0
                for ($i = 0; $i -lt $jobsToLaunch; $i++) {
                    $workItem = $null
                    if ($workQueue.TryDequeue([ref]$workItem)) {
                        # Skip if already processed
                        if ($processedPaths.ContainsKey($workItem.Path)) {
                            continue
                        }

                        # Mark as processed
                        $processedPaths[$workItem.Path] = $true

                        # Launch new job immediately
                        $jobLaunchTime = Get-Date
                        $powershell = [powershell]::Create()
                        $powershell.RunspacePool = $runspacePool
                        $powershell.AddScript($scriptBlock).AddParameter("FolderPath", $workItem.Path).AddParameter("MaxDepth", $MaxDepth).AddParameter("CurrentDepth", $workItem.Depth).AddParameter("EnableDebug", $DebugPreference -eq 'Continue').AddParameter("CentralLogPath", $Script:CentralLogPath) | Out-Null
                        $handle = $powershell.BeginInvoke()
                        $totalJobsLaunched++
                        $jobsLaunchedThisIteration++
                        $newJob = [PSCustomObject]@{
                            PowerShell = $powershell
                            Handle = $handle
                            Path = $workItem.Path
                            Depth = $workItem.Depth
                            StartTime = $jobLaunchTime
                            LaunchOrder = $totalJobsLaunched
                        }
                        [void]$activeJobs.Add($newJob)

                        $jobLaunchTimes[$workItem.Path] = $jobLaunchTime

                        # Only log every 10th job to reduce overhead
                        if ($totalJobsLaunched % 10 -eq 0) {
                            Write-CentralLog -Message "Job #$totalJobsLaunched launched for: $($workItem.Path) (Depth: $($workItem.Depth), Queue: $($workQueue.Count) remaining)" -Category "WORK_QUEUE" -Source "MAIN"
                        }
                    }
                }

                if ($jobsLaunchedThisIteration -gt 0) {
                    Write-CentralLog -Message "BATCH LAUNCH: Launched $jobsLaunchedThisIteration jobs in this iteration" -Category "THREAD_SATURATION" -Source "MAIN"
                }
            }
            # Update active count after launching new jobs - PERFORMANCE: Direct count, no filtering needed!
            $currentActiveCount = $activeJobs.Count

            # Track peak active jobs for utilization analysis
            if ($currentActiveCount -gt $peakActiveJobs) {
                $peakActiveJobs = $currentActiveCount
                Write-CentralLog -Message "NEW PEAK: $peakActiveJobs concurrent active jobs" -Category "PARALLEL_MONITOR" -Source "MAIN"
            }
            # Detailed monitoring report with saturation analysis
            if ($currentTime.Subtract($lastMonitoringReport).TotalMilliseconds -gt $monitoringReportInterval) {
                $utilizationPercent = [math]::Round(($currentActiveCount / $MaxThreads) * 100, 1)
                $elapsedMinutes = [math]::Round($currentTime.Subtract($monitoringStartTime).TotalMinutes, 2)

                Write-CentralLog -Message "=== THREAD SATURATION REPORT @ $elapsedMinutes minutes ===" -Category "THREAD_SATURATION" -Source "MAIN"
                Write-CentralLog -Message "Active jobs: $currentActiveCount/$MaxThreads ($utilizationPercent% utilization)" -Category "THREAD_SATURATION" -Source "MAIN"
                Write-CentralLog -Message "Work queue: $($workQueue.Count) items remaining" -Category "THREAD_SATURATION" -Source "MAIN"
                Write-CentralLog -Message "Total jobs launched: $totalJobsLaunched, Completed: $completedJobsCount" -Category "THREAD_SATURATION" -Source "MAIN"
                Write-CentralLog -Message "Peak concurrent jobs: $peakActiveJobs" -Category "THREAD_SATURATION" -Source "MAIN"
                # Saturation efficiency warning - only warn if we're significantly under-utilizing available threads
                if ($utilizationPercent -lt 80 -and $workQueue.Count -gt 0 -and $availableThreads -gt 2) {
                    Write-CentralLog -Message "SATURATION WARNING: Only $utilizationPercent% utilization with $($workQueue.Count) work items available and $availableThreads threads free!" -Category "THREAD_SATURATION" -Source "MAIN"
                }

                # Show longest running jobs
                if ($activeJobs.Count -gt 0) {
                    $longestRunning = $activeJobs | Sort-Object { $currentTime.Subtract($_.StartTime).TotalSeconds } -Descending | Select-Object -First 3
                    Write-CentralLog -Message "Longest running jobs:" -Category "THREAD_SATURATION" -Source "MAIN"
                    foreach ($longJob in $longestRunning) {
                        $runtime = [math]::Round($currentTime.Subtract($longJob.StartTime).TotalSeconds, 1)
                        Write-CentralLog -Message "  Job #$($longJob.LaunchOrder): $($longJob.Path) (Depth: $($longJob.Depth), Runtime: $runtime seconds)" -Category "THREAD_SATURATION" -Source "MAIN"
                    }
                }
                $lastMonitoringReport = $currentTime
            }
            # HARVEST COMPLETED JOBS - OPTIMIZED USING DIRECT COLLECTION ACCESS
            $jobsCompletedThisIteration = 0
            for ($i = $activeJobs.Count - 1; $i -ge 0; $i--) {
                $job = $activeJobs[$i]

                if ($job.Handle.IsCompleted) {
                    $jobCompletionTime = Get-Date
                    $jobDuration = $jobCompletionTime.Subtract($job.StartTime)
                    $jobCompletionTimes[$job.Path] = $jobCompletionTime
                    $completedJobsCount++
                    $jobsCompletedThisIteration++
                    try {
                        $jobResult = $job.PowerShell.EndInvoke($job.Handle)
                        # Only log every 10th completion to reduce overhead
                        if ($completedJobsCount % 10 -eq 0) {
                            $currentUtilization = [math]::Round(($activeJobs.Count / $MaxThreads) * 100, 1)
                            Write-CentralLog -Message "Job #$($job.LaunchOrder) completed: $($job.Path) (Depth: $($job.Depth)) in $([math]::Round($jobDuration.TotalSeconds, 1))s [Util: $currentUtilization%]" -Category "PARALLEL_MONITOR" -Source "MAIN"
                        }

                        # Handle PSDataCollection results from runspace - extract the actual object
                        $actualResult = $null
                        if ($jobResult -is [System.Management.Automation.PSDataCollection[System.Management.Automation.PSObject]]) {
                            if ($jobResult.Count -gt 0) {
                                $actualResult = $jobResult[0]
                            }
                        } else {
                            $actualResult = $jobResult
                        }

                        # Validate the extracted result
                        if ($actualResult -and $actualResult.PSObject.Properties['Path']) {
                            [void]$allResults.Add($actualResult)

                            # ADD DISCOVERED SUBDIRECTORIES TO WORK QUEUE IMMEDIATELY
                            if ($actualResult.DiscoveredSubdirectories -and $actualResult.DiscoveredSubdirectories.Count -gt 0) {
                                $newWorkItems = 0
                                foreach ($subDir in $actualResult.DiscoveredSubdirectories) {
                                    if (-not $processedPaths.ContainsKey($subDir)) {
                                        $workItem = [PSCustomObject]@{
                                            Path = $subDir
                                            Depth = $job.Depth + 1
                                            Priority = $job.Depth + 1
                                        }
                                        $workQueue.Enqueue($workItem)
                                        $newWorkItems++
                                    }
                                }
                                # Only log significant additions to reduce overhead
                                if ($newWorkItems -gt 5) {
                                    Write-CentralLog -Message "Added $newWorkItems new subdirectories to work queue from $($job.Path)" -Category "WORK_QUEUE" -Source "MAIN"
                                }
                            }
                        }

                    } catch {
                        Write-CentralLog -Message "ERROR: Job #$($job.LaunchOrder) failed for $($job.Path): $($_.Exception.Message)" -Category "PARALLEL_MONITOR" -Source "MAIN"
                    } finally {
                        # CLEAN UP RESOURCES AND MOVE JOB BETWEEN COLLECTIONS
                        $job.PowerShell.Dispose()
                        [void]$completedJobs.Add($job)
                        $activeJobs.RemoveAt($i)
                    }
                }
            }

            # CHECK FOR JOB TIMEOUTS
            for ($i = $activeJobs.Count - 1; $i -ge 0; $i--) {
                $job = $activeJobs[$i]
                $elapsed = (Get-Date).Subtract($job.StartTime).TotalSeconds
                if ($elapsed -gt $individualJobTimeout) {
                    Write-CentralLog -Message "TIMEOUT: Job #$($job.LaunchOrder) for $($job.Path) exceeded $individualJobTimeout seconds" -Category "PARALLEL_MONITOR" -Source "MAIN"
                    try {
                        $job.PowerShell.Stop()
                        $job.PowerShell.Dispose()
                    } catch {
                        Write-CentralLog -Message "Error stopping timed-out job: $($_.Exception.Message)" -Category "PARALLEL_MONITOR" -Source "MAIN"
                    }
                    [void]$completedJobs.Add($job)
                    $activeJobs.RemoveAt($i)
                    $completedJobsCount++
                }
            }
            # Regular progress reporting with saturation metrics
            if ($currentTime.Subtract($lastProgress).TotalMilliseconds -gt $progressCheckInterval) {
                # FIXED: Calculate total work items (jobs launched + jobs still in queue) for accurate progress
                $totalWorkItems = $totalJobsLaunched + $workQueue.Count
                $percentComplete = if ($totalWorkItems -gt 0) { [math]::Round(($completedJobsCount / $totalWorkItems) * 100, 1) } else { 0 }
                $currentUtilization = [math]::Round(($currentActiveCount / $MaxThreads) * 100, 1)
                Write-DetailedProgress -Activity "Parallel Analysis" -Status "Active: $currentActiveCount/$MaxThreads ($currentUtilization%), Queue: $($workQueue.Count), Completed: $completedJobsCount/$totalWorkItems ($percentComplete%)" -PercentComplete $percentComplete -Color "Cyan"

                # Log saturation metrics every progress update
                if ($jobsCompletedThisIteration -gt 0) {
                    Write-CentralLog -Message "SATURATION METRICS: Completed $jobsCompletedThisIteration jobs this iteration, Current active: $currentActiveCount/$MaxThreads" -Category "THREAD_SATURATION" -Source "MAIN"
                }

                $lastProgress = $currentTime
            }
            # ULTRA-LOW LATENCY - Adaptive sleep based on queue size for maximum responsiveness
            # Use burst mode (no sleep) when queue is large, micro-sleep when queue is small
            if ($workQueue.Count -gt ($MaxThreads * 2)) {
                # BURST MODE: No sleep when queue is large to maintain saturation
                # This ensures we launch jobs as fast as possible when work is abundant
            } elseif ($workQueue.Count -gt 0) {
                # Minimal sleep when work is available but queue is smaller
                Start-Sleep -Milliseconds 2
            } else {
                # Slightly longer sleep only when no work is available
                Start-Sleep -Milliseconds 10
            }
        }

        # Final monitoring report
        $totalDuration = (Get-Date).Subtract($monitoringStartTime)
        Write-CentralLog -Message "=== FINAL WORK QUEUE EXECUTION REPORT ===" -Category "PARALLEL_MONITOR" -Source "MAIN"
        Write-CentralLog -Message "Total execution time: $([math]::Round($totalDuration.TotalMinutes, 2)) minutes" -Category "PARALLEL_MONITOR" -Source "MAIN"
        Write-CentralLog -Message "Total jobs launched: $totalJobsLaunched (dynamic work queue)" -Category "PARALLEL_MONITOR" -Source "MAIN"
        Write-CentralLog -Message "Peak concurrent jobs: $peakActiveJobs/$MaxThreads ($([math]::Round(($peakActiveJobs / $MaxThreads) * 100, 1))% peak utilization)" -Category "PARALLEL_MONITOR" -Source "MAIN"
        Write-CentralLog -Message "Work queue final state: $($workQueue.Count) items remaining" -Category "PARALLEL_MONITOR" -Source "MAIN"

        # Calculate average job duration safely
        if ($jobCompletionTimes.Count -gt 0 -and $jobLaunchTimes.Count -gt 0) {
            $durations = @()
            foreach ($completedPath in $jobCompletionTimes.Keys) {
                if ($jobLaunchTimes.ContainsKey($completedPath)) {
                    $duration = $jobCompletionTimes[$completedPath].Subtract($jobLaunchTimes[$completedPath]).TotalSeconds
                    $durations += $duration
                }
            }
            if ($durations.Count -gt 0) {
                $avgDuration = ($durations | Measure-Object -Average).Average
                Write-CentralLog -Message "Average job duration: $([math]::Round($avgDuration, 1)) seconds" -Category "PARALLEL_MONITOR" -Source "MAIN"
            }
        }

        # Enhanced efficiency analysis for work queue system
        if ($peakActiveJobs -eq $MaxThreads) {
            Write-CentralLog -Message "EFFICIENCY EXCELLENT: Achieved 100% thread utilization with work-stealing queue" -Category "PARALLEL_MONITOR" -Source "MAIN"
        } elseif ($peakActiveJobs -ge $MaxThreads * 0.9) {
            Write-CentralLog -Message "EFFICIENCY VERY GOOD: Achieved $([math]::Round(($peakActiveJobs / $MaxThreads) * 100, 1))% peak utilization" -Category "PARALLEL_MONITOR" -Source "MAIN"
        } elseif ($peakActiveJobs -ge $MaxThreads * 0.7) {
            Write-CentralLog -Message "EFFICIENCY GOOD: Achieved $([math]::Round(($peakActiveJobs / $MaxThreads) * 100, 1))% peak utilization" -Category "PARALLEL_MONITOR" -Source "MAIN"
        } else {
            Write-CentralLog -Message "EFFICIENCY WARNING: Only achieved $([math]::Round(($peakActiveJobs / $MaxThreads) * 100, 1))% peak utilization - workload may be too small" -Category "PARALLEL_MONITOR" -Source "MAIN"
        }
        # Handle incomplete jobs - Clean up any remaining active jobs
        if ($activeJobs.Count -gt 0) {
            Write-CentralLog -Message "WARNING: $($activeJobs.Count) jobs did not complete within timeout" -Category "PARALLEL_MONITOR" -Source "MAIN"
            foreach ($incompleteJob in $activeJobs) {
                try {
                    $incompleteJob.PowerShell.Stop()
                    $incompleteJob.PowerShell.Dispose()
                } catch {
                    Write-CentralLog -Message "Error cleaning up incomplete job: $($_.Exception.Message)" -Category "PARALLEL_MONITOR" -Source "MAIN"
                }
            }
        }

        Write-CentralLog -Message "Work-stealing parallel execution completed. Collected $($allResults.Count) total results from $totalJobsLaunched jobs" -Category "PARALLEL_MONITOR" -Source "MAIN"
        return $allResults

    } catch {
        Write-Output "$($Script:Colors.Red)Error in parallel processing: $($_.Exception.Message)$($Script:Colors.Reset)"
        Write-CentralLog -Message "CRITICAL ERROR in parallel processing: $($_.Exception | Out-String)" -Category "PARALLEL_MONITOR" -Source "MAIN"
        return @()
    } finally {
        if ($runspacePool) {
            $runspacePool.Close()
            $runspacePool.Dispose()
        }
    }
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
    # First, show the StartPath itself - this should show ONLY the direct files in the root, not subdirectories
    Write-Output "$($Script:Colors.Bold)$($Script:Colors.Cyan)LEVEL 0: ROOT ANALYSIS ($StartPath)$($Script:Colors.Reset)"
    Write-DebugInfo -Message "Showing hierarchical results - determining root display" -Category "HIERARCHY"
    # The root result should only show direct files in the root directory (not subdirectories)
    $rootResult = $Results | Where-Object { $_.Path -eq $StartPath -and $_.IsAccessible }
    if ($rootResult) {
        Write-DebugInfo -Message "Found root result with size: $(Format-FileSize -SizeInBytes $rootResult.SizeBytes)" -Category "HIERARCHY"
        Write-DebugInfo -Message "Root result files: $($rootResult.FileCount), folders: $($rootResult.SubfolderCount)" -Category "HIERARCHY"
        Show-SingleTable -Results @($rootResult) -Title "Root Directory"
    } else {
        Write-DebugInfo -Message "No root result found in results array" -Category "HIERARCHY"
        Write-Output "$($Script:Colors.White)Root Path: $($Script:Colors.Green)$StartPath$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)No root analysis available$($Script:Colors.Reset)"
    }
    # Level 1: Show top-level subfolders
    Write-Output "`n$($Script:Colors.Bold)$($Script:Colors.Cyan)LEVEL 1: TOP SUBFOLDERS OF $StartPath$($Script:Colors.Reset)"
    Write-DebugInfo -Message "Searching for Level 1 folders in results array" -Category "HIERARCHY"
    Write-DebugInfo -Message "Total results count: $($Results.Count)" -Category "HIERARCHY"
    Write-DebugInfo -Message "StartPath for comparison: '$StartPath'" -Category "HIERARCHY"
    # Normalize the StartPath to ensure consistent comparison
    $normalizedStartPath = $StartPath.TrimEnd('\')
    if ($normalizedStartPath -eq 'C:') { $normalizedStartPath = 'C:\' }
    Write-DebugInfo -Message "Normalized StartPath: '$normalizedStartPath'" -Category "HIERARCHY"
    # Debug: Show sample paths from results for troubleshooting
    if ($DebugPreference -ne 'SilentlyContinue' -and $Results.Count -gt 0) {
        Write-DebugInfo -Message "Sample paths from results (first 10):" -Category "HIERARCHY"
        $sampleResults = $Results | Select-Object -First 10
        foreach ($sample in $sampleResults) {
            $parentPath = Split-Path -Path $sample.Path -Parent
            Write-DebugInfo -Message "  Path: '$($sample.Path)' -> Parent: '$parentPath' (Accessible: $($sample.IsAccessible))" -Category "HIERARCHY"
        }
    }
    $allLevel1Candidates = $Results | Where-Object {
        $_.IsAccessible -and
        $_.Path -ne $normalizedStartPath -and
        $_.Path -ne $StartPath -and
        (
            (Split-Path -Path $_.Path -Parent) -eq $normalizedStartPath -or
            # Handle root drive comparison (C: vs C:\)
            ((Split-Path -Path $_.Path -Parent) -eq $normalizedStartPath.TrimEnd('\') -and $normalizedStartPath -like '*:\')
        )
    }
    Write-DebugInfo -Message "Level 1 candidates found: $($allLevel1Candidates.Count)" -Category "HIERARCHY"
    if ($DebugPreference -ne 'SilentlyContinue' -and $allLevel1Candidates.Count -gt 0) {
        Write-DebugInfo -Message "All Level 1 candidates:" -Category "HIERARCHY"
        foreach ($candidate in $allLevel1Candidates | Sort-Object SizeBytes -Descending) {
            $sizeFormatted = Format-FileSize -SizeInBytes $candidate.SizeBytes
            Write-DebugInfo -Message "  $($candidate.Path): $sizeFormatted (Accessible: $($candidate.IsAccessible))" -Category "HIERARCHY"
        }
    }
    $level1Folders = $allLevel1Candidates | Sort-Object SizeBytes -Descending | Select-Object -First $Top
    if ($level1Folders.Count -gt 0) {
        Write-DebugInfo -Message "Displaying top $($level1Folders.Count) Level 1 folders" -Category "HIERARCHY"
        Show-SingleTable -Results $level1Folders -Title "Level 1 Subfolders"
        # Truly dynamic hierarchical display - show all available levels without artificial limits
        # The script already has timeout mechanisms, so let it show whatever was actually analyzed
        $currentLevelFolders = $level1Folders
        $currentLevel = 1
        while ($currentLevelFolders.Count -gt 0) {
            $largestCurrent = $currentLevelFolders[0]
            # Find subfolders of the largest folder at current level
            $nextLevelPath = $largestCurrent.Path
            $nextLevel = $currentLevel + 1
            Write-Output "`n$($Script:Colors.Bold)$($Script:Colors.Cyan)LEVEL $nextLevel`: TOP SUBFOLDERS OF $nextLevelPath$($Script:Colors.Reset)"
            Write-DebugInfo -Message "Searching for Level $nextLevel subfolders of: $nextLevelPath" -Category "HIERARCHY"
            $nextLevelFolders = $Results | Where-Object {
                $_.IsAccessible -and
                $_.Path -ne $nextLevelPath -and
                (Split-Path -Path $_.Path -Parent) -eq $nextLevelPath
            } | Sort-Object SizeBytes -Descending | Select-Object -First $Top
            Write-DebugInfo -Message "Found $($nextLevelFolders.Count) subfolders for Level $nextLevel" -Category "HIERARCHY"
            if ($nextLevelFolders.Count -gt 0) {
                Show-SingleTable -Results $nextLevelFolders -Title "Level $nextLevel Subfolders"
                $currentLevelFolders = $nextLevelFolders
                $currentLevel = $nextLevel
            } else {
                Write-Output "$($Script:Colors.Yellow)No accessible subfolders found at this level.$($Script:Colors.Reset)"
                break
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
    # Show summary with corrected totals
    Write-DebugInfo -Message "CRITICAL DEBUG: Checking for stored cumulative totals..." -Category "SUMMARY_TOTALS"
    Write-DebugInfo -Message "  Results count: $($Results.Count)" -Category "SUMMARY_TOTALS"
    if ($Results.Count -gt 0) {
        Write-DebugInfo -Message "  First result path: $($Results[0].Path)" -Category "SUMMARY_TOTALS"
        Write-DebugInfo -Message "  First result properties: $($Results[0].PSObject.Properties.Name -join ', ')" -Category "SUMMARY_TOTALS"
    }
    $rootResult = $Results | Where-Object { $_.Path -eq $Results[0].Path -and $_.PSObject.Properties.Name -contains "TotalCumulativeSize" } | Select-Object -First 1
    Write-DebugInfo -Message "  Root result found: $(if ($rootResult) { 'YES' } else { 'NO' })" -Category "SUMMARY_TOTALS"
    if ($rootResult -and $rootResult.PSObject.Properties.Name -contains "TotalCumulativeSize") {
        # Use the stored total cumulative values for accurate summary (0 is a valid value)
        $summaryResults = $Results | Where-Object { $_.Path -ne $Results[0].Path -or (-not $_.PSObject.Properties.Name -contains "TotalCumulativeSize") }
        Write-DebugInfo -Message "Using stored cumulative totals: Size=$(Format-FileSize -SizeInBytes $rootResult.TotalCumulativeSize), Files=$($rootResult.TotalCumulativeFiles)" -Category "SUMMARY"
        Show-HierarchicalSummary -Results $summaryResults -SafeResults $SafeResults -TotalSize $rootResult.TotalCumulativeSize -TotalFiles $rootResult.TotalCumulativeFiles
    } else {
        # Fallback to original method if no stored totals
        Write-DebugInfo -Message "No stored cumulative totals found - using fallback calculation method" -Category "SUMMARY"
        Show-HierarchicalSummary -Results $Results -SafeResults $SafeResults
    }
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
        $color = if ($folder.SizeBytes -gt 30GB) {
            $Script:Colors.Red
        } elseif ($folder.SizeBytes -gt 5GB) {
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
        [array]$SafeResults = @(),
        [long]$TotalSize = -1,
        [long]$TotalFiles = -1
    )
    Write-DebugInfo -Message "CRITICAL DEBUG: Show-HierarchicalSummary called with:" -Category "SUMMARY_DEBUG"
    Write-DebugInfo -Message "  TotalSize parameter: $TotalSize" -Category "SUMMARY_DEBUG"
    Write-DebugInfo -Message "  TotalFiles parameter: $TotalFiles" -Category "SUMMARY_DEBUG"
    Write-DebugInfo -Message "  Results count: $($Results.Count)" -Category "SUMMARY_DEBUG"
    Write-DebugInfo -Message "  SafeResults count: $($SafeResults.Count)" -Category "SUMMARY_DEBUG"
    # Initialize variables at function scope to ensure they're accessible throughout
    # CRITICAL FIX: Use different variable names to avoid overwriting parameters!
    $calculatedSize = 0
    $calculatedFiles = 0
    $totalFolders = 0
    $inaccessibleCount = 0
    # Use provided totals if available, otherwise calculate from top-level results to prevent double-counting
    if ($TotalSize -ge 0 -and $TotalFiles -ge 0) {
        Write-DebugInfo -Message "USING PROVIDED TOTALS PATH" -Category "SUMMARY_DEBUG"
        Write-DebugInfo -Message "PRE-ASSIGNMENT: TotalSize=$TotalSize, TotalFiles=$TotalFiles" -Category "SUMMARY_DEBUG"
        $calculatedSize = $TotalSize
        $calculatedFiles = $TotalFiles
        Write-DebugInfo -Message "POST-ASSIGNMENT: calculatedSize=$calculatedSize, calculatedFiles=$calculatedFiles" -Category "SUMMARY_DEBUG"
        Write-DebugInfo -Message "Using provided accurate totals: Size=$(Format-FileSize -SizeInBytes $calculatedSize), Files=$calculatedFiles" -Category "SUMMARY"
        # When using provided totals, SafeResults are already included, so don't add them again
        # Calculate folder counts normally
        $totalFolders = ($Results | Where-Object { $_.IsAccessible } | Measure-Object -Property SubfolderCount -Sum).Sum
        if ($SafeResults.Count -gt 0) {
            $safeFolders = ($SafeResults | Measure-Object -Property SubfolderCount -Sum).Sum
            $totalFolders += $safeFolders
        }
    } else {
        Write-DebugInfo -Message "USING CALCULATED TOTALS PATH" -Category "SUMMARY_DEBUG"
        # CRITICAL FIX: Calculate totals from TOP-LEVEL accessible results only to prevent double-counting
        # The original logic was summing ALL results which included subdirectories counted multiple times
        $topLevelAccessibleResults = $Results | Where-Object { $_.IsAccessible -and $_.Path -notlike "*\*\*" }
        # If no clear top-level distinction, take only direct children of StartPath
        if ($topLevelAccessibleResults.Count -eq 0) {
            # Fallback: group by parent and only count the parent-level directories
            $groupedByParent = $Results | Where-Object { $_.IsAccessible } | Group-Object { Split-Path $_.Path -Parent }
            $topLevelAccessibleResults = $groupedByParent | ForEach-Object { $_.Group | Sort-Object SizeBytes -Descending | Select-Object -First 1 }
        }
        # Safe calculation with null protection
        $sizeMeasure = $topLevelAccessibleResults | Measure-Object -Property SizeBytes -Sum
        $filesMeasure = $topLevelAccessibleResults | Measure-Object -Property FileCount -Sum
        $calculatedSize = if ($sizeMeasure.Sum) { $sizeMeasure.Sum } else { 0 }
        $calculatedFiles = if ($filesMeasure.Sum) { $filesMeasure.Sum } else { 0 }
        Write-DebugInfo -Message "Calculated from top-level results: Size=$(Format-FileSize -SizeInBytes $calculatedSize), Files=$calculatedFiles" -Category "SUMMARY"
        # Calculate folder counts and add SafeResults when calculating from scratch
        $folderMeasure = ($Results | Where-Object { $_.IsAccessible } | Measure-Object -Property SubfolderCount -Sum)
        $totalFolders = if ($folderMeasure.Sum) { $folderMeasure.Sum } else { 0 }
        # Add safe results to totals only when calculating from scratch
        if ($SafeResults.Count -gt 0) {
            $safeSizeMeasure = ($SafeResults | Measure-Object -Property SizeBytes -Sum)
            $safeFilesMeasure = ($SafeResults | Measure-Object -Property FileCount -Sum)
            $safeFoldersMeasure = ($SafeResults | Measure-Object -Property SubfolderCount -Sum)
            $safeSize = if ($safeSizeMeasure.Sum) { $safeSizeMeasure.Sum } else { 0 }
            $safeFiles = if ($safeFilesMeasure.Sum) { $safeFilesMeasure.Sum } else { 0 }
            $safeFolders = if ($safeFoldersMeasure.Sum) { $safeFoldersMeasure.Sum } else { 0 }
            $calculatedSize += $safeSize
            $calculatedFiles += $safeFiles
            $totalFolders += $safeFolders
        }
    }
    $inaccessibleCount = ($Results | Where-Object { -not $_.IsAccessible }).Count
    # Debug logging for accessibility breakdown with fix details
    Write-DebugInfo -Message "FIXED: Summary calculation method determined:" -Category "ACCESSIBILITY"
    Write-DebugInfo -Message "  Using provided totals: $(if ($TotalSize -ge 0) { 'YES' } else { 'NO' })" -Category "ACCESSIBILITY"
    Write-DebugInfo -Message "  Total results: $($Results.Count)" -Category "ACCESSIBILITY"
    Write-DebugInfo -Message "  Inaccessible results: $inaccessibleCount" -Category "ACCESSIBILITY"
    Write-DebugInfo -Message "  Final calculated size: $(Format-FileSize -SizeInBytes $calculatedSize)" -Category "ACCESSIBILITY"
    Write-DebugInfo -Message "  Error tracker entries: $($Script:ErrorTracker.Count)" -Category "ACCESSIBILITY"
    if ($DebugPreference -eq 'Continue') {
        $inaccessibleResults = $Results | Where-Object { -not $_.IsAccessible }
        if ($inaccessibleResults.Count -gt 0) {
            Write-DebugInfo -Message "Sample inaccessible directories:" -Category "ACCESSIBILITY"
            $sampleErrors = $inaccessibleResults | Select-Object -First 10
            foreach ($errorResult in $sampleErrors) {
                Write-DebugInfo -Message "  Path: '$($errorResult.Path)' | Error: '$($errorResult.Error)' | Type: $($errorResult.GetType().Name) | IsAccessible: $($errorResult.IsAccessible)" -Category "ACCESSIBILITY"
            }
        }
        # Show comparison only if we calculated from top-level results
        if ($TotalSize -lt 0) {
            $allAccessibleResults = $Results | Where-Object { $_.IsAccessible }
            Write-DebugInfo -Message "Double-counting prevention comparison:" -Category "ACCESSIBILITY"
            Write-DebugInfo -Message "  Top-level only total: $(Format-FileSize -SizeInBytes $calculatedSize)" -Category "ACCESSIBILITY"
            Write-DebugInfo -Message "  All results total (old buggy method): $(Format-FileSize -SizeInBytes (($allAccessibleResults | Measure-Object -Property SizeBytes -Sum).Sum))" -Category "ACCESSIBILITY"
        }
        if ($inaccessibleResults.Count -gt 10) {
                Write-DebugInfo -Message "  ... and $($inaccessibleResults.Count - 10) more" -Category "ACCESSIBILITY"
            }
            # Additional analysis - check for null/empty objects
            $nullPaths = $inaccessibleResults | Where-Object { [string]::IsNullOrEmpty($_.Path) }
            Write-DebugInfo -Message "Results with null/empty paths: $($nullPaths.Count)" -Category "ACCESSIBILITY"
            # Check for different types of objects in results
            $resultTypes = $Results | Group-Object { $_.GetType().Name } | Select-Object Name, Count
            Write-DebugInfo -Message "Result types in collection:" -Category "ACCESSIBILITY"
            foreach ($type in $resultTypes) {
                Write-DebugInfo -Message "  $($type.Name): $($type.Count)" -Category "ACCESSIBILITY"        }
    }
    # Ensure all calculation variables are safe
    if (-not $calculatedSize) { $calculatedSize = 0 }
    if (-not $calculatedFiles) { $calculatedFiles = 0 }
    if (-not $totalFolders) { $totalFolders = 0 }

    # Calculate inaccessible count
    $inaccessibleCount = ($Results | Where-Object { -not $_.IsAccessible }).Count
    Write-Output "`n$($Script:Colors.Bold)COMPREHENSIVE SUMMARY STATISTICS$($Script:Colors.Reset)"
    Write-Output "$($Script:Colors.White)Total Size Analyzed: $($Script:Colors.Green)$(Format-FileSize -SizeInBytes $calculatedSize)$($Script:Colors.Reset)"
    Write-Output "$($Script:Colors.White)Total Files: $($Script:Colors.Green)$($calculatedFiles.ToString('N0'))$($Script:Colors.Reset)"
    Write-Output "$($Script:Colors.White)Total Folders: $($Script:Colors.Green)$($totalFolders.ToString('N0'))$($Script:Colors.Reset)"
    Write-Output "$($Script:Colors.White)Processed Folders: $($Script:Colors.Green)$($Script:ProcessedFolders.ToString('N0'))$($Script:Colors.Reset)"
    if ($SafeResults.Count -gt 0) {
        Write-Output "$($Script:Colors.White)Safely Scanned Directories: $($Script:Colors.Cyan)$($SafeResults.Count.ToString('N0'))$($Script:Colors.Reset)"
    }
    if ($inaccessibleCount -gt 0) {
        Write-Output "$($Script:Colors.White)Inaccessible Folders: $($Script:Colors.Red)$($inaccessibleCount.ToString('N0'))$($Script:Colors.Reset)"
    }
    # Check for timed-out directories
    $timedOutDirectories = $Results | Where-Object { $_.Error -and $_.Error -like "*timeout*" }
    if ($timedOutDirectories.Count -gt 0) {
        Write-Output "$($Script:Colors.White)Timed-Out Directories: $($Script:Colors.Yellow)$($timedOutDirectories.Count.ToString('N0'))$($Script:Colors.Reset)"
        Write-DebugInfo -Message "Directories that timed out:" -Category "TIMEOUT_SUMMARY"
        foreach ($dir in $timedOutDirectories) {
            Write-DebugInfo -Message "  $($dir.Path) - $(Format-FileSize -SizeInBytes $dir.SizeBytes) (partial)" -Category "TIMEOUT_SUMMARY"
        }
    }
    # Performance metrics
    $elapsedTime = (Get-Date) - $Script:StartTime
    Write-Output "$($Script:Colors.White)Analysis Duration: $($Script:Colors.Cyan)$($elapsedTime.ToString('hh\:mm\:ss'))$($Script:Colors.Reset)"
    Write-Output "$($Script:Colors.White)Processing Rate: $($Script:Colors.Cyan)$([math]::Round($Script:ProcessedFolders / $elapsedTime.TotalSeconds, 2)) folders/second$($Script:Colors.Reset)"
    # Drive discrepancy analysis with enhanced diagnostics
    try {
        # Try to get the drive information for comparison
        $startPathRoot = [System.IO.Path]::GetPathRoot($Results[0].Path)
        if ($startPathRoot) {
            $driveLetter = $startPathRoot.TrimEnd('\').TrimEnd(':')
            if ($driveLetter -and $driveLetter.Length -eq 1) {
                $volume = Get-Volume -DriveLetter $driveLetter -ErrorAction SilentlyContinue
                if ($volume) {
                    $driveUsedSpace = $volume.Size - $volume.SizeRemaining
                    $discrepancy = $driveUsedSpace - $calculatedSize
                    $discrepancyPercent = if ($driveUsedSpace -gt 0) { [math]::Round(($discrepancy / $driveUsedSpace) * 100, 1) } else { 0 }
                    Write-Output "`n$($Script:Colors.Bold)DRIVE USAGE COMPARISON$($Script:Colors.Reset)"
                    Write-Output "$($Script:Colors.White)Drive Used Space: $($Script:Colors.Cyan)$(Format-FileSize -SizeInBytes $driveUsedSpace)$($Script:Colors.Reset)"
                    Write-Output "$($Script:Colors.White)Script Calculated: $($Script:Colors.Cyan)$(Format-FileSize -SizeInBytes $calculatedSize)$($Script:Colors.Reset)"
                    if ($discrepancy -gt 0) {
                        Write-Output "$($Script:Colors.White)Unaccounted Space: $($Script:Colors.Yellow)$(Format-FileSize -SizeInBytes $discrepancy) ($discrepancyPercent%)$($Script:Colors.Reset)"
                        # Get NTFS overhead estimation
                        $ntfsOverhead = Get-NTFSOverhead -DriveLetter $driveLetter
                        if ($ntfsOverhead.TotalOverhead -gt 0) {
                            Write-Output "`n$($Script:Colors.Bold)NTFS SYSTEM OVERHEAD ANALYSIS$($Script:Colors.Reset)"
                            Write-Output "$($Script:Colors.White)MFT Size: $($Script:Colors.Cyan)$(Format-FileSize -SizeInBytes $ntfsOverhead.MFTSize)$($Script:Colors.Reset)"
                            # Show additional NTFS information if available
                            if ($ntfsOverhead.TotalReservedClusters -gt 0 -and $ntfsOverhead.BytesPerCluster -gt 0) {
                                $reservedSize = $ntfsOverhead.TotalReservedClusters * $ntfsOverhead.BytesPerCluster
                                Write-Output "$($Script:Colors.White)Reserved Clusters: $($Script:Colors.Cyan)$($ntfsOverhead.TotalReservedClusters.ToString('N0')) ($(Format-FileSize -SizeInBytes $reservedSize))$($Script:Colors.Reset)"
                            }
                            if ($ntfsOverhead.MFTZoneSize -gt 0) {
                                Write-Output "$($Script:Colors.White)MFT Zone Size: $($Script:Colors.Cyan)$(Format-FileSize -SizeInBytes $ntfsOverhead.MFTZoneSize)$($Script:Colors.Reset)"
                            }
                            if ($ntfsOverhead.StorageReservedClusters -gt 0 -and $ntfsOverhead.BytesPerCluster -gt 0) {
                                $storageReservedSize = $ntfsOverhead.StorageReservedClusters * $ntfsOverhead.BytesPerCluster
                                Write-Output "$($Script:Colors.White)Storage Reserved: $($Script:Colors.Cyan)$($ntfsOverhead.StorageReservedClusters.ToString('N0')) clusters ($(Format-FileSize -SizeInBytes $storageReservedSize))$($Script:Colors.Reset)"
                            }
                            Write-Output "$($Script:Colors.White)Total NTFS Overhead: $($Script:Colors.Cyan)$(Format-FileSize -SizeInBytes $ntfsOverhead.TotalOverhead)$($Script:Colors.Reset)"
                            Write-Output "$($Script:Colors.DarkGray)Method: $($ntfsOverhead.EstimationMethod)$($Script:Colors.Reset)"
                            # Show additional technical details if available
                            if ($ntfsOverhead.RawNTFSInfo.Count -gt 0) {
                                if ($ntfsOverhead.RawNTFSInfo.ContainsKey('NTFSVersion')) {
                                    Write-Output "$($Script:Colors.DarkGray)NTFS Version: $($ntfsOverhead.RawNTFSInfo['NTFSVersion'])$($Script:Colors.Reset)"
                                }
                                if ($ntfsOverhead.BytesPerCluster -gt 0) {
                                    Write-Output "$($Script:Colors.DarkGray)Cluster Size: $(Format-FileSize -SizeInBytes $ntfsOverhead.BytesPerCluster)$($Script:Colors.Reset)"
                                }
                                if ($ntfsOverhead.RawNTFSInfo.ContainsKey('BytesPerFileRecord')) {
                                    Write-Output "$($Script:Colors.DarkGray)File Record Size: $($ntfsOverhead.RawNTFSInfo['BytesPerFileRecord']) bytes$($Script:Colors.Reset)"
                                }
                            }
                            $discrepancy -= $ntfsOverhead.TotalOverhead
                            $discrepancyPercent = if ($driveUsedSpace -gt 0) { [math]::Round(($discrepancy / $driveUsedSpace) * 100, 1) } else { 0 }
                        }
                        # Estimate inaccessible directories if we have any
                        if ($inaccessibleCount -gt 0) {
                            $inaccessiblePaths = $Script:ErrorTracker.Keys | Where-Object { $_ -like "$($driveLetter):*" }
                            if ($inaccessiblePaths.Count -gt 0) {
                                $inaccessibleEstimate = Get-InaccessibleDirectoryEstimate -InaccessiblePaths $inaccessiblePaths
                                Write-Output "`n$($Script:Colors.Bold)INACCESSIBLE DIRECTORY ANALYSIS$($Script:Colors.Reset)"
                                Write-Output "$($Script:Colors.White)Estimated Size of Inaccessible Dirs: $($Script:Colors.Yellow)$(Format-FileSize -SizeInBytes $inaccessibleEstimate.TotalEstimatedSize)$($Script:Colors.Reset)"
                                Write-Output "$($Script:Colors.DarkGray)Methods Used: $($inaccessibleEstimate.MethodsUsed -join ', ')$($Script:Colors.Reset)"
                                # Show top 5 largest estimated directories
                                $topInaccessible = $inaccessibleEstimate.Details | Sort-Object EstimatedSize -Descending | Select-Object -First 5
                                foreach ($dir in $topInaccessible) {
                                    Write-Output "$($Script:Colors.DarkGray)  $(Split-Path $dir.Path -Leaf): $(Format-FileSize -SizeInBytes $dir.EstimatedSize) ($($dir.Method))$($Script:Colors.Reset)"
                                }
                                $discrepancy -= $inaccessibleEstimate.TotalEstimatedSize
                                $discrepancyPercent = if ($driveUsedSpace -gt 0) { [math]::Round(($discrepancy / $driveUsedSpace) * 100, 1) } else { 0 }
                            }
                        }
                        Write-Output "`n$($Script:Colors.Bold)REMAINING UNACCOUNTED SPACE$($Script:Colors.Reset)"
                        if ($discrepancy -gt 0) {
                            Write-Output "$($Script:Colors.White)After Estimates: $($Script:Colors.Yellow)$(Format-FileSize -SizeInBytes $discrepancy) ($discrepancyPercent%)$($Script:Colors.Reset)"
                            Write-Output "$($Script:Colors.DarkGray)Remaining factors may include:$($Script:Colors.Reset)"
                            # Only show deep directory message if max depth was actually reached
                            if ($Script:MaxDepthReached) {
                                Write-Output "$($Script:Colors.DarkGray)  - Deep directory structures exceeding MaxDepth$($Script:Colors.Reset)"
                            }
                            Write-Output "$($Script:Colors.DarkGray)  - Additional system files and hidden data$($Script:Colors.Reset)"
                            Write-Output "$($Script:Colors.DarkGray)  - File system reserved clusters and bad sectors$($Script:Colors.Reset)"
                            # Only suggest increasing MaxDepth if it was actually reached and discrepancy is significant
                            if ($discrepancyPercent -gt 10 -and $Script:MaxDepthReached) {
                                Write-Output "$($Script:Colors.Yellow)Consider increasing MaxDepth parameter for more complete analysis$($Script:Colors.Reset)"
                            }
                        } else {
                            Write-Output "$($Script:Colors.Green)Excellent accounting! All space accounted for within estimates.$($Script:Colors.Reset)"
                        }
                    } else {
                        Write-Output "$($Script:Colors.Green)Excellent match between calculated and actual drive usage!$($Script:Colors.Reset)"
                    }
                }
            }
        }
    } catch {
        Write-Verbose "Could not perform drive comparison: $($_.Exception.Message)"
    }
}
function Show-ErrorSummary {
    # Enhanced debugging information
    Write-DebugInfo -Message "=== ERROR SUMMARY ANALYSIS ===" -Category "ERROR_SUMMARY"
    Write-DebugInfo -Message "Error tracker contains $($Script:ErrorTracker.Count) entries" -Category "ERROR_SUMMARY"
    Write-DebugInfo -Message "Total inaccessible folders reported during processing: $Script:InaccessibleFolderCount" -Category "ERROR_SUMMARY"
    if ($DebugPreference -eq 'Continue') {
        Write-DebugInfo -Message "Detailed error tracker contents:" -Category "ERROR_SUMMARY"
        $Script:ErrorTracker.GetEnumerator() | Sort-Object Name | ForEach-Object {
            Write-DebugInfo -Message "  $($_.Key) -> $($_.Value)" -Category "ERROR_SUMMARY"
        }
    }
    if ($Script:ErrorTracker.Count -gt 0) {
        Write-Output "`n$($Script:Colors.Bold)ERROR SUMMARY$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.Red)The following $($Script:ErrorTracker.Count) folders could not be accessed:$($Script:Colors.Reset)"
        $Script:ErrorTracker.GetEnumerator() | Sort-Object Name | ForEach-Object {
            Write-Output "$($Script:Colors.DarkGray)  $($_.Key): $($_.Value)$($Script:Colors.Reset)"
        }
        if (-not $Script:IsAdmin) {
            Write-Output "`n$($Script:Colors.Yellow)Note: Running with administrative privileges may provide access to additional folders.$($Script:Colors.Reset)"
        }
        # Additional debugging for discrepancies
        if ($Script:InaccessibleFolderCount -gt $Script:ErrorTracker.Count) {
            Write-Output "`n$($Script:Colors.Yellow)Debug Note: $($Script:InaccessibleFolderCount - $Script:ErrorTracker.Count) additional inaccessible folders were detected but not captured in error summary.$($Script:Colors.Reset)"
        }
    } else {
        Write-DebugInfo -Message "No errors to display in summary" -Category "ERROR_SUMMARY"
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
    } catch {
        # Alternative method: try to stop transcript
        try {
            $originalErrorAction = $ErrorActionPreference
            $ErrorActionPreference = 'Stop'
            Stop-Transcript 2>$null
            # If we get here, transcript was running
            return $true
        } catch {
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
    )
    $result = [PSCustomObject]@{
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
                            }
                        }
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
    } catch {
        $result.Error = "Access Denied: $($_.Exception.Message)"
    }
    return $result
}
function Get-NTFSOverhead {
    <#
    .SYNOPSIS
    Retrieves comprehensive NTFS file system overhead information using fsutil fsinfo ntfsinfo.
    .DESCRIPTION
    Uses fsutil fsinfo ntfsinfo to extract detailed NTFS metadata including MFT size, reserved clusters,
    and other file system overhead that contributes to used space on the drive but is not accounted
    for in standard file enumeration.
    .PARAMETER DriveLetter
    The drive letter (without colon) to analyze for NTFS overhead information.
    .EXAMPLE
    Get-NTFSOverhead -DriveLetter "C"
    Returns NTFS overhead information for the C: drive.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$DriveLetter
    )
    try {
        $overhead = [PSCustomObject]@{
            MFTSize = 0
            TotalReservedClusters = 0
            StorageReservedClusters = 0
            MFTZoneSize = 0
            BytesPerCluster = 0
            TotalOverhead = 0
            EstimationMethod = "Unknown"
            RawNTFSInfo = @{}
        }
        # Execute fsutil fsinfo ntfsinfo to get comprehensive NTFS information
        try {
            Write-DebugInfo -Message "Executing fsutil fsinfo ntfsinfo ${DriveLetter}:" -Category "NTFS"
            $fsutilOutput = & fsutil fsinfo ntfsinfo "${DriveLetter}:" 2>$null
            if ($fsutilOutput -and $fsutilOutput.Count -gt 0) {
                Write-DebugInfo -Message "Successfully retrieved fsutil output with $($fsutilOutput.Count) lines" -Category "NTFS"
                foreach ($line in $fsutilOutput) {
                    $line = $line.Trim()
                    # Parse MFT Valid Data Length (actual MFT size in use)
                    if ($line -match "Mft Valid Data Length\s*:\s*(.+)") {
                        $mftSizeText = $matches[1].Trim()
                        Write-DebugInfo -Message "Found MFT Valid Data Length: '$mftSizeText'" -Category "NTFS"
                        # Parse size with unit (e.g., "1.01 GB")
                        if ($mftSizeText -match "(\d+\.?\d*)\s*(GB|MB|KB|B)") {
                            $mftValue = [double]$matches[1]
                            $mftUnit = $matches[2]
                            switch ($mftUnit) {
                                "GB" { $overhead.MFTSize = [int64]($mftValue * 1GB) }
                                "MB" { $overhead.MFTSize = [int64]($mftValue * 1MB) }
                                "KB" { $overhead.MFTSize = [int64]($mftValue * 1KB) }
                                "B"  { $overhead.MFTSize = [int64]$mftValue }
                            }
                            Write-DebugInfo -Message "Parsed MFT size: $($overhead.MFTSize) bytes" -Category "NTFS"
                        }
                    }
                    # Parse Total Reserved Clusters
                    elseif ($line -match "Total Reserved Clusters\s*:\s*([0-9,]+)\s*\(\s*(.+?)\s*\)") {
                        $reservedClustersText = $matches[1] -replace ',', ''
                        $reservedSizeText = $matches[2].Trim()
                        Write-DebugInfo -Message "Found Total Reserved Clusters: '$reservedClustersText' ($reservedSizeText)" -Category "NTFS"
                        $overhead.TotalReservedClusters = [int64]$reservedClustersText
                        # Parse the size in parentheses
                        if ($reservedSizeText -match "(\d+\.?\d*)\s*(GB|MB|KB|B)") {
                            $reservedValue = [double]$matches[1]
                            $reservedUnit = $matches[2]
                            switch ($reservedUnit) {
                                "GB" { $overhead.RawNTFSInfo['TotalReservedSize'] = [int64]($reservedValue * 1GB) }
                                "MB" { $overhead.RawNTFSInfo['TotalReservedSize'] = [int64]($reservedValue * 1MB) }
                                "KB" { $overhead.RawNTFSInfo['TotalReservedSize'] = [int64]($reservedValue * 1KB) }
                                "B"  { $overhead.RawNTFSInfo['TotalReservedSize'] = [int64]$reservedValue }
                            }
                        }
                    }
                    # Parse Reserved For Storage Reserve
                    elseif ($line -match "Reserved For Storage Reserve\s*:\s*([0-9,]+)\s*\(\s*(.+?)\s*\)") {
                        $storageReservedText = $matches[1] -replace ',', ''
                        $storageSizeText = $matches[2].Trim()
                        Write-DebugInfo -Message "Found Storage Reserved: '$storageReservedText' ($storageSizeText)" -Category "NTFS"
                        $overhead.StorageReservedClusters = [int64]$storageReservedText
                        # Parse the size in parentheses
                        if ($storageSizeText -match "(\d+\.?\d*)\s*(GB|MB|KB|B)") {
                            $storageValue = [double]$matches[1]
                            $storageUnit = $matches[2]
                            switch ($storageUnit) {
                                "GB" { $overhead.RawNTFSInfo['StorageReservedSize'] = [int64]($storageValue * 1GB) }
                                "MB" { $overhead.RawNTFSInfo['StorageReservedSize'] = [int64]($storageValue * 1MB) }
                                "KB" { $overhead.RawNTFSInfo['StorageReservedSize'] = [int64]($storageValue * 1KB) }
                                "B"  { $overhead.RawNTFSInfo['StorageReservedSize'] = [int64]$storageValue }
                            }
                        }
                    }
                    # Parse MFT Zone Size
                    elseif ($line -match "MFT Zone Size\s*:\s*(.+)") {
                        $mftZoneSizeText = $matches[1].Trim()
                        Write-DebugInfo -Message "Found MFT Zone Size: '$mftZoneSizeText'" -Category "NTFS"
                        if ($mftZoneSizeText -match "(\d+\.?\d*)\s*(GB|MB|KB|B)") {
                            $zoneValue = [double]$matches[1]
                            $zoneUnit = $matches[2]
                            switch ($zoneUnit) {
                                "GB" { $overhead.MFTZoneSize = [int64]($zoneValue * 1GB) }
                                "MB" { $overhead.MFTZoneSize = [int64]($zoneValue * 1MB) }
                                "KB" { $overhead.MFTZoneSize = [int64]($zoneValue * 1KB) }
                                "B"  { $overhead.MFTZoneSize = [int64]$zoneValue }
                            }
                        }
                    }
                    # Parse Bytes Per Cluster
                    elseif ($line -match "Bytes Per Cluster\s*:\s*([0-9,]+)") {
                        $overhead.BytesPerCluster = [int]($matches[1] -replace ',', '')
                        Write-DebugInfo -Message "Found Bytes Per Cluster: $($overhead.BytesPerCluster)" -Category "NTFS"
                    }
                    # Store additional useful information
                    elseif ($line -match "Total Sectors\s*:\s*([0-9,]+)\s*\(\s*(.+?)\s*\)") {
                        $overhead.RawNTFSInfo['TotalSectors'] = $matches[1] -replace ',', ''
                        $overhead.RawNTFSInfo['TotalSize'] = $matches[2].Trim()
                    }
                    elseif ($line -match "Free Clusters\s*:\s*([0-9,]+)\s*\(\s*(.+?)\s*\)") {
                        $overhead.RawNTFSInfo['FreeClusters'] = $matches[1] -replace ',', ''
                        $overhead.RawNTFSInfo['FreeSize'] = $matches[2].Trim()
                    }
                    elseif ($line -match "NTFS Version\s*:\s*(.+)") {
                        $overhead.RawNTFSInfo['NTFSVersion'] = $matches[1].Trim()
                    }
                    elseif ($line -match "Bytes Per FileRecord Segment\s*:\s*([0-9,]+)") {
                        $overhead.RawNTFSInfo['BytesPerFileRecord'] = $matches[1] -replace ',', ''
                    }
                }
                # Calculate total overhead from parsed values
                $calculatedOverhead = $overhead.MFTSize
                # Add reserved cluster space if we have cluster size information
                if ($overhead.BytesPerCluster -gt 0) {
                    if ($overhead.TotalReservedClusters -gt 0) {
                        $reservedSpace = $overhead.TotalReservedClusters * $overhead.BytesPerCluster
                        $calculatedOverhead += $reservedSpace
                        Write-DebugInfo -Message "Added reserved clusters overhead: $(Format-FileSize -SizeInBytes $reservedSpace)" -Category "NTFS"
                    }
                    # Add MFT Zone if it is not already included in MFT size
                    if ($overhead.MFTZoneSize -gt 0 -and $overhead.MFTZoneSize -gt $overhead.MFTSize) {
                        $mftZoneOverhead = $overhead.MFTZoneSize - $overhead.MFTSize
                        $calculatedOverhead += $mftZoneOverhead
                        Write-DebugInfo -Message "Added MFT Zone overhead: $(Format-FileSize -SizeInBytes $mftZoneOverhead)" -Category "NTFS"
                    }
                }
                $overhead.TotalOverhead = $calculatedOverhead
                $overhead.EstimationMethod = "fsutil fsinfo ntfsinfo"
                Write-DebugInfo -Message "Successfully calculated total NTFS overhead: $(Format-FileSize -SizeInBytes $overhead.TotalOverhead)" -Category "NTFS"
            }
            else {
                Write-DebugInfo -Message "fsutil command returned no output" -Category "NTFS"
                throw "fsutil fsinfo ntfsinfo returned no output"
            }
        } catch {
            Write-DebugInfo -Message "Failed to execute fsutil or parse output: $($_.Exception.Message)" -Category "NTFS"
            Write-Verbose "Could not get NTFS info via fsutil: $($_.Exception.Message)"
            # Fallback: Estimate based on drive size (typical NTFS overhead is 1-3% of drive size)
            try {
                $volume = Get-Volume -DriveLetter $DriveLetter -ErrorAction Stop
                $estimatedOverhead = [math]::Round($volume.Size * 0.02, 0)
                $overhead.TotalOverhead = $estimatedOverhead
                $overhead.EstimationMethod = "Percentage estimate (2%) - fsutil failed"
                Write-DebugInfo -Message "Using fallback percentage estimate: $(Format-FileSize -SizeInBytes $overhead.TotalOverhead)" -Category "NTFS"
            } catch {
                $overhead.EstimationMethod = "Unable to estimate - both fsutil and volume query failed"
                Write-DebugInfo -Message "All estimation methods failed" -Category "NTFS"
            }
        }
        return $overhead
    } catch {
        Write-DebugInfo -Message "Critical error in Get-NTFSOverhead: $($_.Exception.Message)" -Category "NTFS"
        return [PSCustomObject]@{
            MFTSize = 0
            TotalReservedClusters = 0
            StorageReservedClusters = 0
            MFTZoneSize = 0
            BytesPerCluster = 0
            TotalOverhead = 0
            EstimationMethod = "Error: $($_.Exception.Message)"
            RawNTFSInfo = @{}
        }
    }
}
function Get-InaccessibleDirectoryEstimate {
    <#
    .SYNOPSIS
    Attempts to estimate the size of inaccessible directories using alternative methods.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [array]$InaccessiblePaths
    )
    $estimate = [PSCustomObject]@{
        TotalEstimatedSize = 0
        MethodsUsed = @()
        Details = @()
    }
    foreach ($path in $InaccessiblePaths) {
        $pathEstimate = 0
        $method = "None"
        try {
            # Method 1: Try to get directory size using Get-ChildItem with minimal access
            try {
                $items = Get-ChildItem -Path $path -Force -ErrorAction Stop | Select-Object -First 1
                if ($items) {
                    # If we can list at least one item, estimate based on accessible parent directory patterns
                    $parentSize = try { (Get-ChildItem -Path (Split-Path $path -Parent) -Directory -Force | Where-Object { $_.Name -ne (Split-Path $path -Leaf) } | Measure-Object -Property Length -Sum).Sum } catch { 0 }
                    $pathEstimate = [math]::Max(1MB, $parentSize * 0.1)
                    $method = "Parent directory pattern"
                }
            } catch {
                Write-Verbose "Could not analyze parent directory pattern for $path"
            }
            # Method 2: Try using WMI/CIM to get folder information
            if ($pathEstimate -eq 0) {
                try {
                    $folderPath = $path.Replace('\', '\\')
                    $folder = Get-CimInstance -ClassName Win32_Directory -Filter "Name='$folderPath'" -ErrorAction Stop 2>$null
                    if ($folder -and $folder.FileSize) {
                        $pathEstimate = $folder.FileSize
                        $method = "WMI Directory query"
                    }
                } catch {
                    Write-Verbose "Could not query WMI for directory $path"
                }
            }
            # Method 3: Check if it's a known system directory and use typical sizes
            if ($pathEstimate -eq 0) {
                $dirName = Split-Path $path -Leaf
                switch -Regex ($dirName) {
                    "^Users$" {
                        $pathEstimate = 50GB
                        $method = "Known directory type estimate (Users)"
                    }
                    "^Windows$" {
                        $pathEstimate = 25GB
                        $method = "Known directory type estimate (Windows)"
                    }
                    "^WinSxS$" {
                        $pathEstimate = 15GB
                        $method = "Known directory type estimate (WinSxS)"
                    }
                    "^System32$" {
                        $pathEstimate = 5GB
                        $method = "Known directory type estimate (System32)"
                    }
                    "^Temp" {
                        $pathEstimate = 2GB
                        $method = "Known directory type estimate (Temp)"
                    }
                    default {
                        $pathEstimate = 100MB
                        $method = "Default estimate"
                    }
                }
            }
            $estimate.TotalEstimatedSize += $pathEstimate
            $estimate.Details += [PSCustomObject]@{
                Path = $path
                EstimatedSize = $pathEstimate
                Method = $method
            }
            if ($method -notin $estimate.MethodsUsed) {
                $estimate.MethodsUsed += $method
            }
        } catch {
            # Even if we can't estimate, add a minimal placeholder
            $estimate.TotalEstimatedSize += 50MB
            $estimate.Details += [PSCustomObject]@{
                Path = $path
                EstimatedSize = 50MB
                Method = "Fallback estimate"
            }
        }
    }
    return $estimate
}
function Get-OptimizedDirectoryList {
    <#
    .SYNOPSIS
    Optimizes directory list for better thread utilization using fast, lightweight job distribution.
    .DESCRIPTION
    This function implements a fast job distribution system that minimizes upfront directory scanning.
    Instead of doing expensive hierarchical analysis, it uses dynamic work queue discovery during parallel execution.

    PERFORMANCE OPTIMIZED VERSION:
    - Minimal upfront directory enumeration to avoid delays
    - Simple level-1 expansion only when absolutely necessary
    - Relies on dynamic work queue for deep directory discovery
    - Fast startup with optimal thread utilization
    .PARAMETER TopLevelDirs
    Array of top-level directory paths to potentially expand.
    .PARAMETER MaxThreads
    Maximum number of threads available for parallel processing.
    .PARAMETER MaxDepth
    Maximum depth for directory scanning - used for optimization strategy selection.
    .EXAMPLE
    $optimizedDirs = Get-OptimizedDirectoryList -TopLevelDirs $topLevelDirs -MaxThreads 25 -MaxDepth 50
    Returns an optimized list of directories that maximizes thread utilization with minimal overhead.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$TopLevelDirs,
        [Parameter(Mandatory = $true)]
        [int]$MaxThreads,
        [Parameter(Mandatory = $true)]
        [int]$MaxDepth
    )
    Write-CentralLog -Message "=== FAST JOB DISTRIBUTION ANALYSIS ===" -Category "JOB_DISTRIBUTION" -Source "MAIN"
    Write-CentralLog -Message "Initial directory count: $($TopLevelDirs.Count), Target threads: $MaxThreads, MaxDepth: $MaxDepth" -Category "JOB_DISTRIBUTION" -Source "MAIN"
    Write-CentralLog -Message "Current utilization: $([math]::Round(($TopLevelDirs.Count / $MaxThreads) * 100, 1))%" -Category "JOB_DISTRIBUTION" -Source "MAIN"
    # Fast decision logic - minimal upfront work
    $minJobsForOptimalUtilization = [math]::Max($MaxThreads, 10)

    if ($TopLevelDirs.Count -ge $minJobsForOptimalUtilization) {
        Write-CentralLog -Message "DECISION: Sufficient directories for optimal utilization ($($TopLevelDirs.Count) >= $minJobsForOptimalUtilization)" -Category "JOB_DISTRIBUTION" -Source "MAIN"
        Write-CentralLog -Message "Using original top-level directory list - dynamic work queue will handle deep discovery" -Category "JOB_DISTRIBUTION" -Source "MAIN"
        return $TopLevelDirs
    }
    # Only do minimal expansion if we have very few directories
    if ($TopLevelDirs.Count -lt ($MaxThreads / 2) -and $MaxDepth -gt 1) {
        Write-CentralLog -Message "Very few directories ($($TopLevelDirs.Count)) - performing fast level-1 expansion only" -Category "JOB_DISTRIBUTION" -Source "MAIN"

        $expandedDirs = @()
        $expandedDirs += $TopLevelDirs

        foreach ($topDir in $TopLevelDirs) {
            try {
                # Fast enumeration - just get immediate subdirectories
                $subDirs = Get-ChildItem -Path $topDir -Directory -Force -ErrorAction Stop | Select-Object -First 20
                foreach ($subDir in $subDirs) {
                    $expandedDirs += $subDir.FullName
                    if ($expandedDirs.Count -ge ($MaxThreads * 2)) {
                        break
                    }
                }
                if ($expandedDirs.Count -ge ($MaxThreads * 2)) {
                    break
                }
            } catch {
                Write-CentralLog -Message "Could not expand $topDir (continuing with others)" -Category "JOB_DISTRIBUTION" -Source "MAIN"
                continue
            }
        }

        if ($expandedDirs.Count -gt $TopLevelDirs.Count) {
            Write-CentralLog -Message "DECISION: Fast expansion successful ($($TopLevelDirs.Count) -> $($expandedDirs.Count) directories)" -Category "JOB_DISTRIBUTION" -Source "MAIN"
            return $expandedDirs | Select-Object -Unique
        }
    }

    Write-CentralLog -Message "DECISION: Using original directory list - dynamic work queue optimal for this scenario" -Category "JOB_DISTRIBUTION" -Source "MAIN"
    Write-CentralLog -Message "Final job distribution: $($TopLevelDirs.Count) directories for $MaxThreads threads" -Category "JOB_DISTRIBUTION" -Source "MAIN"
    return $TopLevelDirs
}

function Get-HierarchicalJobSplit {
    <#
    .SYNOPSIS
    Implements hierarchical job splitting for deep scan scenarios.
    .DESCRIPTION
    This function addresses the thread starvation problem in deep scans by intelligently
    splitting large directory trees into smaller, more manageable jobs that can be
    distributed across multiple threads.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$TopLevelDirs,
        [Parameter(Mandatory = $true)]
        [int]$MaxDepth,
        [Parameter(Mandatory = $true)]
        [int]$TargetJobCount
    )

    Write-CentralLog -Message "=== HIERARCHICAL JOB SPLITTING ===" -Category "JOB_DISTRIBUTION" -Source "MAIN"

    $allJobs = @()
    $allJobs += $TopLevelDirs
    $processedDirs = @{}
    $expansionLevel = 2
    $maxExpansionLevel = [math]::Min(4, [math]::Floor($MaxDepth / 3))

    Write-CentralLog -Message "Target job count: $TargetJobCount" -Category "JOB_DISTRIBUTION" -Source "MAIN"
    Write-CentralLog -Message "Max expansion level: $maxExpansionLevel" -Category "JOB_DISTRIBUTION" -Source "MAIN"
    while ($allJobs.Count -lt $TargetJobCount -and $expansionLevel -le $maxExpansionLevel) {
        Write-CentralLog -Message "Expansion level ${expansionLevel}: Current jobs: $($allJobs.Count), Target: $TargetJobCount" -Category "JOB_DISTRIBUTION" -Source "MAIN"

        $newJobs = @()
        $dirsToExpand = @()

        # Identify directories that haven't been expanded at this level
        foreach ($dir in $TopLevelDirs) {
            $expandKey = "$dir|$expansionLevel"
            if (-not $processedDirs.ContainsKey($expandKey)) {
                $dirsToExpand += $dir
            }
        }

        if ($dirsToExpand.Count -eq 0) {
            Write-CentralLog -Message "No more directories available for expansion at level $expansionLevel" -Category "JOB_DISTRIBUTION" -Source "MAIN"
            break
        }

        foreach ($baseDir in $dirsToExpand) {
            if ($allJobs.Count -ge $TargetJobCount) {
                break
            }

            try {
                $expandKey = "$baseDir|$expansionLevel"
                $processedDirs[$expandKey] = $true

                # Get subdirectories at the target expansion level
                $subDirs = Get-DirectoriesAtLevel -BasePath $baseDir -TargetLevel $expansionLevel -MaxDepth $MaxDepth

                Write-CentralLog -Message "Found $($subDirs.Count) directories at level $expansionLevel under $baseDir" -Category "JOB_DISTRIBUTION" -Source "MAIN"

                foreach ($subDir in $subDirs) {
                    if ($allJobs.Count -lt $TargetJobCount) {
                        # Quick complexity assessment
                        $complexity = Get-DirectoryComplexity -Path $subDir.FullName

                        if ($complexity.ShouldSplit -and $allJobs.Count -lt ($TargetJobCount - 5)) {
                            # For very complex directories, split them further
                            $splitDirs = Get-DirectoriesAtLevel -BasePath $subDir.FullName -TargetLevel 1 -MaxDepth ($MaxDepth - $expansionLevel)
                            if ($splitDirs.Count -gt 1 -and ($allJobs.Count + $splitDirs.Count) -le $TargetJobCount) {
                                foreach ($splitDir in $splitDirs) {
                                    $newJobs += $splitDir.FullName
                                }
                                Write-CentralLog -Message "Split complex directory $($subDir.FullName) into $($splitDirs.Count) jobs" -Category "JOB_DISTRIBUTION" -Source "MAIN"
                            } else {
                                $newJobs += $subDir.FullName
                            }
                        } else {
                            $newJobs += $subDir.FullName
                        }
                    }
                }
            } catch {
                Write-CentralLog -Message "Could not expand $baseDir at level $expansionLevel : $($_.Exception.Message)" -Category "JOB_DISTRIBUTION" -Source "MAIN"
            }
        }

        if ($newJobs.Count -gt 0) {
            # Replace original directories with expanded ones, avoiding duplicates
            $allJobs = @()
            $allJobs += $newJobs | Select-Object -Unique
            Write-CentralLog -Message "Level $expansionLevel expansion complete: $($allJobs.Count) total jobs" -Category "JOB_DISTRIBUTION" -Source "MAIN"
        }

        $expansionLevel++
    }

    Write-CentralLog -Message "Hierarchical job splitting complete: $($allJobs.Count) jobs created" -Category "JOB_DISTRIBUTION" -Source "MAIN"
    return $allJobs
}

function Get-DirectoriesAtLevel {
    <#
    .SYNOPSIS
    Gets directories at a specific depth level from a base path.
    .OUTPUTS
    System.Object[]
    #>
    [CmdletBinding()]
    [OutputType([System.Object[]])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$BasePath,
        [Parameter(Mandatory = $true)]
        [int]$TargetLevel,
        [Parameter(Mandatory = $true)]
        [int]$MaxDepth
    )

    if ($TargetLevel -gt $MaxDepth -or $TargetLevel -lt 1) {
        return @()
    }

    try {
        $currentLevel = @($BasePath)

        for ($level = 1; $level -le $TargetLevel; $level++) {
            $nextLevel = @()
            foreach ($path in $currentLevel) {
                try {
                    $subDirs = Get-ChildItem -Path $path -Directory -Force -ErrorAction Stop
                    foreach ($subDir in $subDirs) {
                        # Skip problematic directories
                        $skipPatterns = @(
                            "*`$RECYCLE.BIN*", "*System Volume Information*", "*Recovery*",
                            "*Documents and Settings*", "*`$WinREAgent*", "*.gradle*"
                        )

                        $shouldSkip = $false
                        foreach ($pattern in $skipPatterns) {
                            if ($subDir.FullName -like $pattern) {
                                $shouldSkip = $true
                                break
                            }
                        }

                        if (-not $shouldSkip) {
                            $nextLevel += $subDir
                        }
                    }
                } catch {
                    # Skip inaccessible directories
                    continue
                }
            }
            $currentLevel = $nextLevel | Select-Object -ExpandProperty FullName

            if ($currentLevel.Count -eq 0) {
                break
            }
        }

        return $currentLevel | ForEach-Object { [PSCustomObject]@{ FullName = $_ } }
    } catch {
        return @()
    }
}

function Get-DirectoryComplexity {
    <#
    .SYNOPSIS
    Performs a quick assessment of directory complexity to determine if it should be split.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        # Quick sample of subdirectories (don't enumerate all)
        $subDirSample = Get-ChildItem -Path $Path -Directory -Force -ErrorAction Stop | Select-Object -First 50
        $subDirCount = $subDirSample.Count

        # Heuristics for complexity
        $isComplex = $false
        $reason = "Standard directory"

        if ($subDirCount -gt 20) {
            $isComplex = $true
            $reason = "High subdirectory count ($subDirCount)"
        } elseif ($Path -like "*gradle*" -or $Path -like "*node_modules*" -or $Path -like "*\.git*") {
            $isComplex = $true
            $reason = "Known complex directory type"
        } elseif ($Path.Length -gt 180) {
            $isComplex = $true
            $reason = "Very deep path (${$Path.Length} chars)"
        }

        return [PSCustomObject]@{
            ShouldSplit = $isComplex
            Reason = $reason
            SubdirectoryCount = $subDirCount
        }
    } catch {
        return [PSCustomObject]@{
            ShouldSplit = $false
            Reason = "Inaccessible directory"
            SubdirectoryCount = 0
        }
    }
}

function Get-SimpleLevel2Expansion {
    <#
    .SYNOPSIS
    Implements simple level 2 expansion for shallow scan scenarios (original logic).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$TopLevelDirs,
        [Parameter(Mandatory = $true)]
        [int]$TargetJobCount
    )

    Write-CentralLog -Message "=== SIMPLE LEVEL 2 EXPANSION ===" -Category "JOB_DISTRIBUTION" -Source "MAIN"

    $level2Dirs = @()
    $expandedCount = 0
    $skippedCount = 0
    $errorCount = 0

    foreach ($topDir in $TopLevelDirs) {
        try {
            Write-CentralLog -Message "Analyzing subdirectories of: $topDir" -Category "JOB_DISTRIBUTION" -Source "MAIN"
            $subDirs = Get-ChildItem -Path $topDir -Directory -Force -ErrorAction Stop

            Write-CentralLog -Message "Found $($subDirs.Count) subdirectories in $topDir" -Category "JOB_DISTRIBUTION" -Source "MAIN"

            foreach ($subDir in $subDirs) {
                if ($level2Dirs.Count -lt $TargetJobCount) {
                    # Skip problematic directories
                    $skipPatterns = @(
                        "*`$RECYCLE.BIN*", "*System Volume Information*", "*Recovery*",
                        "*Documents and Settings*", "*`$WinREAgent*"
                    )

                    $shouldSkip = $false
                    foreach ($pattern in $skipPatterns) {
                        if ($subDir.FullName -like $pattern) {
                            $shouldSkip = $true
                            $skippedCount++
                            Write-CentralLog -Message "Skipping problematic directory: $($subDir.FullName)" -Category "JOB_DISTRIBUTION" -Source "MAIN"
                            break
                        }
                    }

                    if (-not $shouldSkip) {
                        $level2Dirs += $subDir.FullName
                        $expandedCount++
                        Write-CentralLog -Message "Added level 2 directory: $($subDir.FullName)" -Category "JOB_DISTRIBUTION" -Source "MAIN"
                    }
                } else {
                    Write-CentralLog -Message "Reached expansion limit ($TargetJobCount) - stopping expansion" -Category "JOB_DISTRIBUTION" -Source "MAIN"
                    break
                }
            }
        } catch {
            $errorCount++
            Write-CentralLog -Message "Could not expand $topDir : $($_.Exception.Message)" -Category "JOB_DISTRIBUTION" -Source "MAIN"
        }
    }

    Write-CentralLog -Message "Simple expansion complete: $($level2Dirs.Count) directories found" -Category "JOB_DISTRIBUTION" -Source "MAIN"
    Write-CentralLog -Message "Directories expanded: $expandedCount, Skipped: $skippedCount, Errors: $errorCount" -Category "JOB_DISTRIBUTION" -Source "MAIN"

    return $level2Dirs
}

function Optimize-HierarchicalSize {
    <#
    .SYNOPSIS
    Aggregates child directory sizes up to their parent directories to fix size calculation discrepancies.
    .DESCRIPTION
    The parallel processing system scans directories individually but doesn't aggregate child sizes upward.
    This function takes the flat results and properly calculates cumulative sizes by adding child directory
    sizes to their parent directories, ensuring accurate hierarchical size reporting.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.ArrayList]$Results
    )

    Write-DebugInfo -Message "Starting hierarchical size aggregation for $($Results.Count) results" -Category "SIZE_AGGREGATION"
    Write-CentralLog -Message "=== HIERARCHICAL SIZE AGGREGATION PROCESS ===" -Category "SIZE_AGGREGATION" -Source "MAIN"

    # Create a hashtable for fast path lookup
    $pathLookup = @{
    }
    $accessibleResults = @()

    # Populate lookup table with accessible results only
    foreach ($result in $Results) {
        if ($result.IsAccessible -and $result.Path) {
            $pathLookup[$result.Path] = $result
            $accessibleResults += $result
        }
    }

    Write-DebugInfo -Message "Built path lookup table with $($pathLookup.Count) accessible directories" -Category "SIZE_AGGREGATION"
    Write-CentralLog -Message "Path lookup table built: $($pathLookup.Count) accessible directories" -Category "SIZE_AGGREGATION" -Source "MAIN"

    # Sort by path depth (deepest first) to aggregate bottom-up
    $sortedResults = $accessibleResults | Sort-Object {
        ($_.Path -split '\\').Count
    } -Descending

    Write-DebugInfo -Message "Sorted results by depth for bottom-up aggregation" -Category "SIZE_AGGREGATION"

    $processedPaths = @{
    }
    $aggregationStats = @{
        DirectoriesProcessed = 0
        SizeAggregations = 0
        TotalSizeAdded = 0
    }

    # Process each directory from deepest to shallowest
    foreach ($result in $sortedResults) {
        if ($processedPaths.ContainsKey($result.Path)) {
            continue
        }

        $parentPath = Split-Path -Path $result.Path -Parent
        if ([string]::IsNullOrEmpty($parentPath) -or $parentPath -eq $result.Path) {
            # This is a root directory, skip
            $processedPaths[$result.Path] = $true
            continue
        }
        # Check if parent exists in our results
        if ($pathLookup.ContainsKey($parentPath)) {
            $parentResult = $pathLookup[$parentPath]

            # Add this directory's size to its parent
            $parentResult.SizeBytes += $result.SizeBytes
            $parentResult.FileCount += $result.FileCount

            # Track cloud files propagation
            if ($result.HasCloudFiles) {
                $parentResult.HasCloudFiles = $true
            }

            # Update largest file if necessary
            if ($result.LargestFileSize -gt $parentResult.LargestFileSize) {
                $parentResult.LargestFile = $result.LargestFile
                $parentResult.LargestFileSize = $result.LargestFileSize
            }

            $aggregationStats.SizeAggregations++
            $aggregationStats.TotalSizeAdded += $result.SizeBytes

            # Log significant aggregations for debugging
            if ($result.SizeBytes -gt 100MB) {
                Write-DebugInfo -Message "Aggregated $($result.Path) ($(Format-FileSize -SizeInBytes $result.SizeBytes)) into parent $parentPath" -Category "SIZE_AGGREGATION"
                Write-CentralLog -Message "Significant aggregation: $(Format-FileSize -SizeInBytes $result.SizeBytes) from $($result.Path) to $parentPath" -Category "SIZE_AGGREGATION" -Source "MAIN"
            }
        } else {
            Write-DebugInfo -Message "Parent path $parentPath not found in results for $($result.Path)" -Category "SIZE_AGGREGATION"
        }

        $processedPaths[$result.Path] = $true
        $aggregationStats.DirectoriesProcessed++

        # Progress reporting for large datasets
        if ($aggregationStats.DirectoriesProcessed % 1000 -eq 0) {
            Write-DebugInfo -Message "Processed $($aggregationStats.DirectoriesProcessed) directories for size aggregation" -Category "SIZE_AGGREGATION"
        }
    }

    Write-DebugInfo -Message "Hierarchical size aggregation completed:" -Category "SIZE_AGGREGATION"
    Write-DebugInfo -Message "  Directories processed: $($aggregationStats.DirectoriesProcessed)" -Category "SIZE_AGGREGATION"
    Write-DebugInfo -Message "  Size aggregations performed: $($aggregationStats.SizeAggregations)" -Category "SIZE_AGGREGATION"
    Write-DebugInfo -Message "  Total size aggregated: $(Format-FileSize -SizeInBytes $aggregationStats.TotalSizeAdded)" -Category "SIZE_AGGREGATION"

    Write-CentralLog -Message "=== HIERARCHICAL SIZE AGGREGATION COMPLETED ===" -Category "SIZE_AGGREGATION" -Source "MAIN"
    Write-CentralLog -Message "Directories processed: $($aggregationStats.DirectoriesProcessed)" -Category "SIZE_AGGREGATION" -Source "MAIN"
    Write-CentralLog -Message "Size aggregations: $($aggregationStats.SizeAggregations)" -Category "SIZE_AGGREGATION" -Source "MAIN"
    Write-CentralLog -Message "Total size aggregated: $(Format-FileSize -SizeInBytes $aggregationStats.TotalSizeAdded)" -Category "SIZE_AGGREGATION" -Source "MAIN"

    return $Results
}
# MAIN SCRIPT EXECUTION
function Main {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$StartPath,
        [int]$MaxDepth,
        [int]$Top,
        [int]$MaxThreads
    )
    $scriptPath = if ($MyInvocation.MyCommand.Path) {
        Split-Path -Parent $MyInvocation.MyCommand.Path
    } else {
        Split-Path -Parent $PSCommandPath
    }
    $logPath = Start-AdvancedTranscript -LogPath $scriptPath
    try {
        # Initialize and display header
        Write-Output "$($Script:Colors.Bold)$($Script:Colors.Cyan)===============================================$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.Bold)$($Script:Colors.Cyan)    ULTRA-FAST FOLDER USAGE ANALYZER v5.4.0    $($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.Bold)$($Script:Colors.Cyan)===============================================$($Script:Colors.Reset)"
        # Initialize central logging for debug messages ONLY if -Debug is specified
        if ($DebugPreference -ne 'SilentlyContinue') {
            Initialize-CentralLogging | Out-Null
            Write-Output "$($Script:Colors.Magenta)Debug mode enabled - Central debug logging active$($Script:Colors.Reset)"
        } else {
            Write-Verbose "Running in normal mode - Debug logging disabled, Verbose output goes to transcript"
        }
        # Administrative privilege check
        $Script:IsAdmin = Test-AdminPrivilege
        $adminStatus = if ($Script:IsAdmin) {
            "$($Script:Colors.Green)Administrator$($Script:Colors.Reset)"
        } else {
            "$($Script:Colors.Yellow)Standard User$($Script:Colors.Reset)"
        }
        Write-Output "$($Script:Colors.White)Running as: $adminStatus$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Start Path: $($Script:Colors.Cyan)$($PSBoundParameters['StartPath'])$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Max Depth: $($Script:Colors.Cyan)$($PSBoundParameters['MaxDepth'])$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Top Folders: $($Script:Colors.Cyan)$($PSBoundParameters['Top'])$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Max Threads: $($Script:Colors.Cyan)$($PSBoundParameters['MaxThreads'])$($Script:Colors.Reset)"
        # Extract and display drive information for the target drive
        try {
            $driveLetter = [System.IO.Path]::GetPathRoot($PSBoundParameters['StartPath']).TrimEnd('\').TrimEnd(':')
            if ($driveLetter -and $driveLetter.Length -eq 1) {
                Show-DriveInfo -DriveLetter $driveLetter
            }
        } catch {
            Write-Output "$($Script:Colors.Yellow)Warning: Could not extract drive information from path: $($PSBoundParameters['StartPath'])$($Script:Colors.Reset)"
        }
        Write-Output ""
        if ($PSCmdlet.ShouldProcess($PSBoundParameters['StartPath'], "Analyze folder usage")) {
            # Get top-level directories for parallel processing
            Write-ProgressUpdate -Activity "Initialization" -Status "Scanning top-level directories..." -Color "Cyan"
            # Known problematic directories that can cause hangs
            $problematicDirs = @(
                "System Volume Information",
                "`$Recycle.Bin",
                "Recovery",
                "Documents and Settings",
                "`$WinREAgent"
            )
            # System files that should be included in size calculation
            $systemFiles = @(
                "hiberfil.sys",
                "pagefile.sys",
                "swapfile.sys"
            )
            $topLevelDirs = @()
            $problematicDirsFound = @()
            $systemFilesFound = @()
            try {
                $items = Get-ChildItem -Path $PSBoundParameters['StartPath'] -Directory -Force -ErrorAction Stop
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
                # Check for system files at root level
                try {
                    $rootItems = Get-ChildItem -Path $PSBoundParameters['StartPath'] -File -Force -ErrorAction Stop
                    foreach ($file in $rootItems) {
                        if ($systemFiles -contains $file.Name) {
                            $systemFilesFound += $file
                        }
                    }
                } catch {
                    Write-Output "$($Script:Colors.Yellow)Warning: Could not check for system files: $($_.Exception.Message)$($Script:Colors.Reset)"
                }
                $Script:TotalFolders = $topLevelDirs.Count
                Write-Output "$($Script:Colors.Green)Found $($topLevelDirs.Count) top-level directories to analyze$($Script:Colors.Reset)"
                # Debug logging for directory discovery
                Write-DebugInfo -Message "Total directories found in $($PSBoundParameters['StartPath']): $($items.Count)" -Category "DISCOVERY"
                Write-DebugInfo -Message "Normal directories: $($topLevelDirs.Count)" -Category "DISCOVERY"
                Write-DebugInfo -Message "Problematic directories: $($problematicDirsFound.Count)" -Category "DISCOVERY"
                  if ($DebugPreference -ne 'SilentlyContinue') {
                    Write-DebugInfo -Message "Normal directories list:" -Category "DIRECTORY_LIST"
                    foreach ($dir in $topLevelDirs | Sort-Object) {
                        Write-DebugInfo -Message "  - $dir" -Category "DIRECTORY_LIST"
                    }
                }
                if ($problematicDirsFound.Count -gt 0) {
                    Write-Output "$($Script:Colors.Yellow)Found $($problematicDirsFound.Count) problematic directories - will analyze safely$($Script:Colors.Reset)"
                      if ($DebugPreference -ne 'SilentlyContinue') {
                        Write-DebugInfo -Message "Problematic directories list:" -Category "DIRECTORY_LIST"
                        foreach ($dir in $problematicDirsFound) {
                            Write-DebugInfo -Message "  - $($dir.FullName)" -Category "DIRECTORY_LIST"
                        }
                    }
                }
                if ($systemFilesFound.Count -gt 0) {
                    $totalSystemFileSize = ($systemFilesFound | Measure-Object -Property Length -Sum).Sum
                    Write-Output "$($Script:Colors.Yellow)Found $($systemFilesFound.Count) system files ($(Format-FileSize -SizeInBytes $totalSystemFileSize)) - will include in total$($Script:Colors.Reset)"
                }
            } catch {
                Write-Output "$($Script:Colors.Red)Error accessing start path: $($_.Exception.Message)$($Script:Colors.Reset)"
                return
            }
            if ($topLevelDirs.Count -eq 0) {
                Write-Output "$($Script:Colors.Yellow)No directories found to analyze.$($Script:Colors.Reset)"
                return
            }
            # Perform parallel analysis with dynamic job distribution optimization
            Write-DetailedProgress -Activity "Analysis" -Status "Optimizing job distribution for maximum thread utilization..." -Color "Cyan" -Detail "Initial: $($topLevelDirs.Count) directories, Available: $($PSBoundParameters['MaxThreads']) threads"

            # Use optimized directory list for better thread utilization
            $optimizedDirs = Get-OptimizedDirectoryList -TopLevelDirs $topLevelDirs -MaxThreads $PSBoundParameters['MaxThreads'] -MaxDepth $PSBoundParameters['MaxDepth']

            Write-DetailedProgress -Activity "Analysis" -Status "Starting parallel folder analysis..." -Color "Cyan" -Detail "Processing $($optimizedDirs.Count) directories with $($PSBoundParameters['MaxThreads']) threads, MaxDepth=$($PSBoundParameters['MaxDepth'])"
            Write-DebugInfo -Message "Starting parallel analysis with optimized parameters:" -Category "PARALLEL_START"
            Write-DebugInfo -Message "  Original directories: $($topLevelDirs.Count)" -Category "PARALLEL_START"
            Write-DebugInfo -Message "  Optimized directories: $($optimizedDirs.Count)" -Category "PARALLEL_START"
            Write-DebugInfo -Message "  Max Depth: $($PSBoundParameters['MaxDepth'])" -Category "PARALLEL_START"
            Write-DebugInfo -Message "  Max Threads: $($PSBoundParameters['MaxThreads'])" -Category "PARALLEL_START"
            Write-DebugInfo -Message "  Thread utilization: $([math]::Round(([math]::Min($optimizedDirs.Count, $PSBoundParameters['MaxThreads']) / $PSBoundParameters['MaxThreads']) * 100, 1))%" -Category "PARALLEL_START"
            $results = Get-ParallelFolderStatistic -FolderPaths $optimizedDirs -MaxDepth $PSBoundParameters['MaxDepth'] -MaxThreads $PSBoundParameters['MaxThreads']
            Write-DebugInfo -Message "Parallel analysis completed. Results count: $($results.Count)" -Category "PARALLEL_COMPLETE"

            # CRITICAL FIX: Create placeholder root entry BEFORE hierarchical aggregation to prevent size loss
            Write-ProgressUpdate -Activity "Analysis" -Status "Creating root placeholder for hierarchical aggregation..." -Color "Cyan"
            $rootPlaceholder = [PSCustomObject]@{
                Path = $PSBoundParameters['StartPath']
                SizeBytes = 0
                FileCount = 0
                SubfolderCount = 0
                LargestFile = $null
                LargestFileSize = 0
                IsAccessible = $true
                HasCloudFiles = $false
                Error = $null
            }
            $results = @($rootPlaceholder) + $results
            Write-DebugInfo -Message "Root placeholder created for hierarchical aggregation" -Category "ROOT_CALC"

            # CRITICAL FIX: Perform hierarchical size aggregation to correctly calculate cumulative directory sizes
            Write-ProgressUpdate -Activity "Analysis" -Status "Performing hierarchical size aggregation..." -Color "Cyan"
            $results = Optimize-HierarchicalSize -Results $results
            Write-DebugInfo -Message "Hierarchical size aggregation completed for $($results.Count) results" -Category "PARALLEL_COMPLETE"
            # Debug: Analyze the composition of results before cleanup
            if ($DebugPreference -eq 'Continue' -and $results.Count -gt 0) {
                $resultTypes = $results | Group-Object { $_.GetType().Name } | Select-Object Name, Count
                Write-DebugInfo -Message "BEFORE CLEANUP - Result types:" -Category "PARALLEL_COMPLETE"
                foreach ($type in $resultTypes) {
                    Write-DebugInfo -Message "  $($type.Name): $($type.Count)" -Category "PARALLEL_COMPLETE"
                }
                # Show sample of non-PSCustomObject results
                $invalidResults = $results | Where-Object { $_ -isnot [PSCustomObject] }
                if ($invalidResults.Count -gt 0) {
                    Write-DebugInfo -Message "Sample invalid results (first 5):" -Category "PARALLEL_COMPLETE"
                    for ($i = 0; $i -lt [Math]::Min($invalidResults.Count, 5); $i++) {
                        Write-DebugInfo -Message "  Invalid[$i]: Type=$($invalidResults[$i].GetType().Name), Value='$($invalidResults[$i])'" -Category "PARALLEL_COMPLETE"
                    }
                }
            }
            if ($DebugPreference -ne 'SilentlyContinue' -and $results.Count -gt 0) {
                $accessibleResults = $results | Where-Object { $_.IsAccessible }
                $inaccessibleResults = $results | Where-Object { -not $_.IsAccessible }
                Write-DebugInfo -Message "  Accessible directories: $($accessibleResults.Count)" -Category "PARALLEL_COMPLETE"
                Write-DebugInfo -Message "  Inaccessible directories: $($inaccessibleResults.Count)" -Category "PARALLEL_COMPLETE"
                if ($accessibleResults.Count -gt 0) {
                    Write-DebugInfo -Message "Top accessible directories by size:" -Category "PARALLEL_RESULTS"
                    $topResults = $accessibleResults | Sort-Object SizeBytes -Descending | Select-Object -First 5
                    foreach ($result in $topResults) {
                        $sizeFormatted = Format-FileSize -SizeInBytes $result.SizeBytes
                        Write-DebugInfo -Message "  $($result.Path): $sizeFormatted" -Category "PARALLEL_RESULTS"
                    }
                }
            }
            # Safely analyze problematic directories first (before root calculation)
            $safeResults = @()
            if ($problematicDirsFound.Count -gt 0) {
                Write-ProgressUpdate -Activity "Analysis" -Status "Analyzing problematic directories safely..." -Color "Yellow"
                foreach ($problematicDir in $problematicDirsFound) {
                    Write-Output "$($Script:Colors.DarkGray)Safely scanning: $($problematicDir.FullName)$($Script:Colors.Reset)"
                    $problematicResult = Get-ProblematicDirectorySize -DirectoryPath $problematicDir.FullName -DirectoryName $problematicDir.Name
                    $safeResults += $problematicResult
                    $Script:ProcessedFolders++
                }
            }
            # Analyze the StartPath directory itself for files at the root level
            Write-ProgressUpdate -Activity "Analysis" -Status "Analyzing root directory files..." -Color "Cyan"
            try {
                $rootFiles = Get-ChildItem -Path $PSBoundParameters['StartPath'] -File -Force -ErrorAction Stop
                # CRITICAL FIX: Calculate ONLY the size of TOP-LEVEL directories to prevent double-counting
                # The original logic was summing ALL accessible results, which included nested subdirectories
                # causing massive over-reporting (e.g., C:\Program Files would be counted, plus C:\Program Files\SubDir, etc.)

                # Find the root placeholder (first result) which now contains aggregated sizes from hierarchical processing
                $rootPlaceholder = $results | Where-Object { $_.Path -eq $PSBoundParameters['StartPath'] } | Select-Object -First 1
                $aggregatedRootSize = if ($rootPlaceholder) { $rootPlaceholder.SizeBytes } else { 0 }

                Write-DebugInfo -Message "FIXED: Root size calculation with hierarchical aggregation:" -Category "ROOT_CALC"
                Write-DebugInfo -Message "  Aggregated root size from hierarchy: $(Format-FileSize -SizeInBytes $aggregatedRootSize)" -Category "ROOT_CALC"

                $rootFilesSize = if ($rootFiles.Count -gt 0) { ($rootFiles | Measure-Object -Property Length -Sum).Sum } else { 0 }
                # Add problematic directories size (safely scanned) - these are separate from normal subdirectories
                $problematicDirsSize = if ($safeResults.Count -gt 0) { ($safeResults | Measure-Object -Property SizeBytes -Sum).Sum } else { 0 }

                # Total cumulative size includes aggregated subdirectories + root files + problematic directories
                $totalCumulativeSize = $aggregatedRootSize + $rootFilesSize + $problematicDirsSize

                # For the ROOT directory display, show aggregated subdirectories + direct files + problematic directories
                $rootDisplaySize = $aggregatedRootSize + $rootFilesSize + $problematicDirsSize
                # CRITICAL DEBUG: Check if results array is empty causing 0 calculation
                Write-DebugInfo -Message "CRITICAL DEBUG: Results array analysis:" -Category "ROOT_CALC"
                Write-DebugInfo -Message "  Results count: $($results.Count)" -Category "ROOT_CALC"
                Write-DebugInfo -Message "  Accessible results count: $(($results | Where-Object { $_.IsAccessible }).Count)" -Category "ROOT_CALC"
                if ($results.Count -eq 0) {
                    Write-DebugInfo -Message "  WARNING: Results array is empty - this explains 0.00 B calculation!" -Category "ROOT_CALC"
                } elseif (($results | Where-Object { $_.IsAccessible }).Count -eq 0) {
                    Write-DebugInfo -Message "  WARNING: No accessible results found - this explains 0.00 B calculation!" -Category "ROOT_CALC"
                }
                # Debug logging for root calculation with fix details
                Write-DebugInfo -Message "FIXED: Root calculation breakdown (with hierarchical aggregation):" -Category "ROOT_CALC"
                Write-DebugInfo -Message "  Aggregated subdirectory size: $(Format-FileSize -SizeInBytes $aggregatedRootSize)" -Category "ROOT_CALC"
                Write-DebugInfo -Message "  Root files size: $(Format-FileSize -SizeInBytes $rootFilesSize) ($($rootFiles.Count) files)" -Category "ROOT_CALC"
                Write-DebugInfo -Message "  Problematic dirs size: $(Format-FileSize -SizeInBytes $problematicDirsSize)" -Category "ROOT_CALC"
                Write-DebugInfo -Message "  Root display size: $(Format-FileSize -SizeInBytes $rootDisplaySize)" -Category "ROOT_CALC"
                Write-DebugInfo -Message "  TOTAL cumulative size (accurate): $(Format-FileSize -SizeInBytes $totalCumulativeSize)" -Category "ROOT_CALC"
                if ($DebugPreference -ne 'SilentlyContinue' -and $rootFiles.Count -gt 0) {
                    Write-DebugInfo -Message "Root files found:" -Category "ROOT_FILES"
                    $sortedRootFiles = $rootFiles | Sort-Object Length -Descending | Select-Object -First 10
                    foreach ($file in $sortedRootFiles) {
                        $fileSize = Format-FileSize -SizeInBytes $file.Length
                        Write-DebugInfo -Message "  $($file.Name): $fileSize" -Category "ROOT_FILES"
                    }
                }
                # Calculate cumulative file count: subdirectories + root files + problematic directories
                $subdirectoryTotalFiles = ($results | Where-Object { $_.IsAccessible } | Measure-Object -Property FileCount -Sum).Sum
                $problematicDirsFiles = if ($safeResults.Count -gt 0) { ($safeResults | Measure-Object -Property FileCount -Sum).Sum } else { 0 }
                $totalCumulativeFiles = $subdirectoryTotalFiles + $rootFiles.Count + $problematicDirsFiles
                Write-DebugInfo -Message "File count breakdown:" -Category "ROOT_CALC"
                Write-DebugInfo -Message "  Subdirectory files: $subdirectoryTotalFiles" -Category "ROOT_CALC"
                Write-DebugInfo -Message "  Root files: $($rootFiles.Count)" -Category "ROOT_CALC"
                Write-DebugInfo -Message "  Problematic dir files: $problematicDirsFiles" -Category "ROOT_CALC"
                Write-DebugInfo -Message "  TOTAL files: $totalCumulativeFiles" -Category "ROOT_CALC"
                # Determine largest file across all analyzed content (including problematic directories)
                $largestFileFromSubdirs = $results | Where-Object { $_.IsAccessible -and $_.LargestFileSize -gt 0 } | Sort-Object LargestFileSize -Descending | Select-Object -First 1
                $largestFileFromRoot = if ($rootFiles.Count -gt 0) { $rootFiles | Sort-Object Length -Descending | Select-Object -First 1 } else { $null }
                $largestFileFromProblematic = if ($safeResults.Count -gt 0) { $safeResults | Where-Object { $_.LargestFileSize -gt 0 } | Sort-Object LargestFileSize -Descending | Select-Object -First 1 } else { $null }
                $overallLargestFile = $null
                $overallLargestFileSize = 0
                # Compare all largest files to find the overall largest
                $candidates = @()
                if ($largestFileFromSubdirs) { $candidates += @{ File = $largestFileFromSubdirs.LargestFile; Size = $largestFileFromSubdirs.LargestFileSize } }
                if ($largestFileFromRoot) { $candidates += @{ File = $largestFileFromRoot.FullName; Size = $largestFileFromRoot.Length } }
                if ($largestFileFromProblematic) { $candidates += @{ File = $largestFileFromProblematic.LargestFile; Size = $largestFileFromProblematic.LargestFileSize } }
                if ($candidates.Count -gt 0) {
                    $winner = $candidates | Sort-Object Size -Descending | Select-Object -First 1
                    $overallLargestFile = $winner.File
                    $overallLargestFileSize = $winner.Size
                }
                # Check for cloud files in root, subdirectories, and problematic directories
                $hasCloudFilesInRoot = if ($rootFiles.Count -gt 0) { ($rootFiles | Where-Object { Test-CloudPlaceholder -FileInfo $_ }).Count -gt 0 } else { $false }
                $hasCloudFilesInSubdirs = ($results | Where-Object { $_.HasCloudFiles }).Count -gt 0
                $hasCloudFilesInProblematic = ($safeResults | Where-Object { $_.HasCloudFiles }).Count -gt 0
                $overallHasCloudFiles = $hasCloudFilesInRoot -or $hasCloudFilesInSubdirs -or $hasCloudFilesInProblematic
                # Count total subfolders including problematic directories
                $totalSubfolderCount = $topLevelDirs.Count + $problematicDirsFound.Count + $systemFilesFound.Count
                $rootStats = [PSCustomObject]@{
                    Path = $StartPath
                    SizeBytes = $rootDisplaySize
                    FileCount = $rootFiles.Count + $problematicDirsFiles
                    SubfolderCount = $totalSubfolderCount
                    LargestFile = $overallLargestFile
                    LargestFileSize = $overallLargestFileSize
                    IsAccessible = $true
                    HasCloudFiles = $overallHasCloudFiles
                    Error = $null
                }
                # Store the total cumulative size for summary statistics
                $rootStats | Add-Member -NotePropertyName "TotalCumulativeSize" -NotePropertyValue $totalCumulativeSize
                $rootStats | Add-Member -NotePropertyName "TotalCumulativeFiles" -NotePropertyValue $totalCumulativeFiles

                # CRITICAL FIX: Replace the root placeholder with the fully calculated root stats
                $nonRootResults = $results | Where-Object { $_.Path -ne $PSBoundParameters['StartPath'] }
                $results = @($rootStats) + $nonRootResults
                Write-DebugInfo -Message "Root placeholder replaced with calculated root stats" -Category "ROOT_CALC"
                $Script:ProcessedFolders++
            } catch {
                Write-Output "$($Script:Colors.Yellow)Warning: Could not analyze root directory files: $($_.Exception.Message)$($Script:Colors.Reset)"
            }
            # Secondary analysis: Collect direct subfolders of the largest directories for hierarchical display
            Write-ProgressUpdate -Activity "Analysis" -Status "Collecting subfolder details for hierarchical display..." -Color "Cyan"
            $additionalResults = @()
            # Add the additional results to the main results for hierarchical display
            $results += $additionalResults
            # Display results using hierarchical view
            # FINAL CLEANUP: Remove any invalid results that might have been added
            $cleanResults = $results | Where-Object {
                ($_ -is [PSCustomObject] -or $_.PSObject.TypeNames -contains 'Deserialized.System.Management.Automation.PSCustomObject') -and
                $_.PSObject.Properties.Name -contains 'Path' -and
                $_.PSObject.Properties.Name -contains 'IsAccessible' -and
                -not [string]::IsNullOrEmpty($_.Path)
            }
            Write-DebugInfo -Message "FINAL CLEANUP: Original results: $($results.Count), Clean results: $($cleanResults.Count)" -Category "CLEANUP"
            $results = $cleanResults
            Show-HierarchicalResult -Results $results -StartPath $PSBoundParameters['StartPath'] -Top $PSBoundParameters['Top'] -SafeResults $safeResults
            Show-ErrorSummary
            Write-Output "`n$($Script:Colors.Green)Analysis completed successfully!$($Script:Colors.Reset)"
        }
        else {
            Write-Output "$($Script:Colors.Yellow)WhatIf: Would analyze folder usage starting from '$($PSBoundParameters['StartPath'])'$($Script:Colors.Reset)"
            Write-Output "$($Script:Colors.Yellow)WhatIf: Would scan up to $($PSBoundParameters['MaxDepth']) levels deep$($Script:Colors.Reset)"
            Write-Output "$($Script:Colors.Yellow)WhatIf: Would display top $($PSBoundParameters['Top']) largest folders$($Script:Colors.Reset)"
            Write-Output "$($Script:Colors.Yellow)WhatIf: Would use $($PSBoundParameters['MaxThreads']) parallel threads$($Script:Colors.Reset)"
        }
    } catch {
        Write-Output "$($Script:Colors.Red)Fatal error: $($_.Exception.Message)$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.Red)Stack trace: $($_.ScriptStackTrace)$($Script:Colors.Reset)"
    }
    finally {
        # Cleanup and finalization - ALWAYS executed
        Write-ProgressUpdate -Activity "Analysis" -Status "Complete" -Completed
        # Clean up central logging
        Clear-CentralLogging
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
            } catch {
                # Ultimate fallback - force stop without our function
                try {
                    Stop-Transcript -ErrorAction SilentlyContinue
                    Write-Output "$($Script:Colors.Yellow)Transcript stopped using emergency fallback$($Script:Colors.Reset)"
                } catch {
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
    # Use parameters directly to satisfy PSScriptAnalyzer scoping requirements
    $scriptStartPath = $StartPath
    $scriptMaxDepth = $MaxDepth
    $scriptTop = $Top
    $scriptMaxThreads = $MaxThreads

    Main -StartPath $scriptStartPath -MaxDepth $scriptMaxDepth -Top $scriptTop -MaxThreads $scriptMaxThreads -WhatIf:$WhatIfPreference
}

