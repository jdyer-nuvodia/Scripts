# =============================================================================
# Script: Analyze-FolderUsage.ps1
# Created: 2025-06-13 20:57:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-06-18 00:10:00 UTC
# Updated By: jdyer-nuvodia
# Version: 2.7.0
# Additional Info: FIXED - Runspace debug messages now properly displayed; PSDataCollection handling corrected
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
- Multi-level progress reporting with ANSI color coding
- Advanced transcript logging with cleanup
- Drive information display for the target drive

The script uses PowerShell runspaces for maximum performance and includes memory management with garbage collection.
Supports Windows PowerShell 5.1 and later with modern .NET integration.
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
    [int]$Top = 10,    [Parameter(Mandatory = $false)]
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
$Script:InaccessibleFolderCount = 0
$Script:StartTime = Get-Date
$Script:MaxDepthReached = $false

# Centralized logging variables
$Script:CentralLogPath = $null
$Script:LogMutex = $null
$Script:EnableCentralLogging = $false

# Debug and Verbose Logging Functions
function Write-CentralLog {
    <#
    .SYNOPSIS
    Thread-safe centralized logging function that captures output from both main thread and runspaces.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [string]$Category = "INFO",

        [Parameter(Mandatory = $false)]
        [string]$Source = "MAIN",

        [Parameter(Mandatory = $false)]
        [switch]$NoConsole
    )

    if (-not $Script:EnableCentralLogging -or [string]::IsNullOrEmpty($Script:CentralLogPath)) {
        return
    }

    try {
        # Create timestamp
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
        $logEntry = "[$timestamp] [$Source] [$Category] $Message"

        # Write to console unless suppressed
        if (-not $NoConsole) {
            switch ($Category) {
                "ERROR" { Write-Output "$($Script:Colors.Red)$logEntry$($Script:Colors.Reset)" }
                "WARNING" { Write-Output "$($Script:Colors.Yellow)$logEntry$($Script:Colors.Reset)" }
                "DEBUG" { Write-Output "$($Script:Colors.DarkGray)$logEntry$($Script:Colors.Reset)" }
                "SUCCESS" { Write-Output "$($Script:Colors.Green)$logEntry$($Script:Colors.Reset)" }
                default { Write-Output "$($Script:Colors.White)$logEntry$($Script:Colors.Reset)" }
            }
        }

        # Thread-safe file writing
        $acquired = $false
        try {
            if ($Script:LogMutex) {
                $acquired = $Script:LogMutex.WaitOne(1000) # 1 second timeout
            }

            if ($acquired -or -not $Script:LogMutex) {
                # Write to log file
                Add-Content -Path $Script:CentralLogPath -Value $logEntry -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch {
            # Silently continue if logging fails to avoid disrupting main functionality
            Write-Verbose "Central logging failed: $($_.Exception.Message)"
        }
        finally {
            if ($acquired -and $Script:LogMutex) {
                $Script:LogMutex.ReleaseMutex()
            }
        }
    }
    catch {
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
        Write-CentralLog -Message $Message -Category $Category -Source "MAIN" -NoConsole
    }
}

function Write-VerboseInfo {
    [CmdletBinding()]
    param(
        [string]$Message,
        [string]$Category = "VERBOSE"
    )
    if ($VerbosePreference -ne 'SilentlyContinue' -or $DebugPreference -ne 'SilentlyContinue') {
        $timestamp = Get-Date -Format "HH:mm:ss.fff"
        Write-Output "$($Script:Colors.DarkGray)[$timestamp] [$Category] $Message$($Script:Colors.Reset)"
    }
}

function Write-DetailedProgress {
    [CmdletBinding()]
    param(
        [string]$Activity,
        [string]$Status,
        [string]$Detail = "",
        [string]$Color = "Cyan"
    )

    Write-ProgressUpdate -Activity $Activity -Status $Status -Color $Color

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
        $logFileName = "Analyze-FolderUsage_${computerName}_${timestamp}.log"
        $fullLogPath = Join-Path -Path $LogPath -ChildPath $logFileName

        # Ensure log directory exists
        if (-not (Test-Path -Path $LogPath)) {
            if ($PSCmdlet.ShouldProcess($LogPath, "Create log directory")) {
                New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
            }
        }

        # Initialize centralized logging
        $Script:CentralLogPath = $fullLogPath
        $Script:EnableCentralLogging = $true

        # Create mutex for thread-safe logging
        try {
            $Script:LogMutex = New-Object System.Threading.Mutex($false, "AnalyzeFolderUsageLogMutex")
        }
        catch {
            Write-Output "$($Script:Colors.Yellow)Warning: Could not create log mutex, logging may not be thread-safe$($Script:Colors.Reset)"
            $Script:LogMutex = $null
        }

        # Create initial log file with header
        $headerText = @"
===============================================
ULTRA-FAST FOLDER USAGE ANALYZER v2.7.0 LOG
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
            Write-CentralLog -Message "Centralized logging initialized" -Category "SYSTEM" -Source "MAIN"
        }
        return $fullLogPath
    }
    catch {
        Write-Output "$($Script:Colors.Yellow)Warning: Could not start transcript: $($_.Exception.Message). Continuing without logging.$($Script:Colors.Reset)"
        $Script:EnableCentralLogging = $false
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
            # Write final log entry
            Write-CentralLog -Message "Stopping transcript and cleaning up logging" -Category "SYSTEM" -Source "MAIN"

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
    finally {
        # Clean up centralized logging
        if ($Script:LogMutex) {
            try {
                $Script:LogMutex.Dispose()            }
            catch {
                # Ignore disposal errors
                Write-Verbose "Mutex disposal failed: $($_.Exception.Message)"
            }
            $Script:LogMutex = $null
        }

        $Script:EnableCentralLogging = $false
        $Script:CentralLogPath = $null
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
    }
    catch {
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
        # More specific OneDrive/cloud storage detection
        # Only check for actual cloud placeholder files, not all reparse points

        # Validate input to prevent null reference exceptions
        if ($null -eq $FileInfo) {
            Write-DebugInfo -Message "Null FileInfo passed to Test-CloudPlaceholder" -Category "CLOUD_CHECK"
            return $false
        }

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
        # Create scriptblock for parallel execution with comprehensive debugging
        $scriptBlock = {
            param($FolderPath, $MaxDepth, $EnableDebug, $CentralLogPath)

            # Ensure all output goes to the right place - suppress any cmdlet binding issues
            $ErrorActionPreference = 'Continue'
            $DebugPreference = 'SilentlyContinue'
            $VerbosePreference = 'SilentlyContinue'
            $InformationPreference = 'SilentlyContinue'
            $WarningPreference = 'SilentlyContinue'
            # Debug flag passed from main script
              $DebugEnabled = $EnableDebug

              # Make CentralLogPath available to nested functions
              $script:CentralLogPath = $CentralLogPath

              # Debug output function for runspace - using centralized logging
              function Write-RunspaceDebug {
                param([string]$Message, [string]$Category = "RUNSPACE")

                if ($DebugEnabled -and $script:CentralLogPath) {
                    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
                    $logEntry = "[$timestamp] [RUNSPACE] [$Category] $Message"

                    # Write directly to the centralized log file
                    try {
                        # Use Add-Content for thread-safe writing (basic level)
                        Add-Content -Path $script:CentralLogPath -Value $logEntry -Encoding UTF8 -ErrorAction SilentlyContinue
                    } catch {
                        # Fallback to error stream if file writing fails
                        try {
                            Microsoft.PowerShell.Utility\Write-Error -Message "RUNSPACE_DEBUG: [$timestamp] [$Category] $Message" -ErrorId "DebugMessage:$Category" -Category NotSpecified                        } catch {
                            # Last resort - do nothing rather than contaminate output
                            Write-Verbose "Runspace debug fallback failed: $($_.Exception.Message)"
                        }
                    }
                }
            }
            # Import required functions into runspace
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
                # Add timeout mechanism for individual directory scans
                $startTime = Get-Date
                $maxScanTime = 60 # Maximum 60 seconds per directory
                $maxItemsPerDirectory = 10000 # Limit items processed per directory

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
                    Write-RunspaceDebug -Message "Starting scan of: $FolderPath (Depth: $CurrentDepth/$MaxDepth)" -Category "FOLDER_SCAN"
                    # Check if it's a reparse point first
                    if (Test-ReparsePoint -Path $FolderPath) {
                        Write-RunspaceDebug -Message "Skipping reparse point: $FolderPath" -Category "FOLDER_SCAN"
                        return $stats
                    }
                    Write-RunspaceDebug -Message "Executing: Get-ChildItem -Path '$FolderPath' -Force -ErrorAction Stop" -Category "FOLDER_SCAN"
                    $items = Get-ChildItem -Path $FolderPath -Force -ErrorAction Stop
                    Write-RunspaceDebug -Message "SUCCESS: Scanned $FolderPath - found $($items.Count) items" -Category "FOLDER_SCAN"

                    # Limit the number of items processed to prevent hanging on huge directories
                    if ($items.Count -gt $maxItemsPerDirectory) {
                        Write-RunspaceDebug -Message "Large directory detected ($($items.Count) items), limiting to $maxItemsPerDirectory items" -Category "FOLDER_SCAN"
                        $items = $items | Select-Object -First $maxItemsPerDirectory
                        $stats.Error = "Large directory - only processed first $maxItemsPerDirectory items"
                    }

                    $itemCount = 0
                    foreach ($item in $items) {
                        $itemCount++
                        # Check timeout before processing each item
                        if ((Get-Date).Subtract($startTime).TotalSeconds -gt $maxScanTime) {
                            Write-RunspaceDebug -Message "Timeout reached for $FolderPath after $maxScanTime seconds, stopping scan" -Category "TIMEOUT"
                            $stats.Error = "Scan timeout after $maxScanTime seconds"
                            break
                        }

                        if ($item.PSIsContainer) {
                            $stats.SubfolderCount++
                            if ($CurrentDepth -lt $MaxDepth) {
                                Write-RunspaceDebug -Message "Recursing into subfolder: $($item.FullName)" -Category "FOLDER_SCAN"
                                $subStats = Get-FolderStatisticInternal -FolderPath $item.FullName -CurrentDepth ($CurrentDepth + 1) -MaxDepth $MaxDepth                                # Explicit null check to ensure we don't process empty results
                                if ($subStats) {
                                    $stats.SizeBytes += $subStats.SizeBytes
                                    $stats.FileCount += $subStats.FileCount
                                    $stats.SubfolderCount += $subStats.SubfolderCount

                                    if ($subStats.LargestFileSize -gt $stats.LargestFileSize) {
                                        $stats.LargestFile = $subStats.LargestFile
                                        $stats.LargestFileSize = $subStats.LargestFileSize
                                    }

                                    if ($subStats.HasCloudFiles) { $stats.HasCloudFiles = $true }

                                    # Propagate MaxDepthReached flag
                                    if ($subStats.MaxDepthReached) {
                                        $stats.MaxDepthReached = $true
                                    }

                                    # Propagate inaccessibility from subdirectories
                                    if (-not $subStats.IsAccessible) {
                                        Write-RunspaceDebug -Message "Subfolder marked as inaccessible: $($item.FullName)" -Category "ERROR_PROPAGATION"
                                    }
                                }
                            } else {
                                Write-RunspaceDebug -Message "Max depth reached, skipping subfolder: $($item.FullName)" -Category "FOLDER_SCAN"
                                # Max depth reached - set flag
                                $stats.MaxDepthReached = $true
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
                    $errorMessage = $_.Exception.Message
                    $stats.Error = $errorMessage
                    Write-RunspaceDebug -Message "ERROR: Failed to access $FolderPath - $errorMessage" -Category "FOLDER_ERROR"
                    Write-RunspaceDebug -Message "Exception Type: $($_.Exception.GetType().FullName)" -Category "FOLDER_ERROR"
                    Write-RunspaceDebug -Message "Full Exception: $($_.Exception | Out-String)" -Category "FOLDER_ERROR"
                }
                return $stats
            }
            # Execute the main function and ensure clean return value
            try {
                $finalResult = Get-FolderStatisticInternal -FolderPath $FolderPath -CurrentDepth 0 -MaxDepth $MaxDepth
                Write-RunspaceDebug -Message "Final result for $FolderPath`: Size=$($finalResult.SizeBytes), Accessible=$($finalResult.IsAccessible)" -Category "RETURN"                # Ensure we return only the PSCustomObject and suppress any other output
                # BREAKPOINT: Set breakpoint here to see what gets returned from runspace
                Write-Output $finalResult
                return
            } catch {
                # Return error result if main function fails
                $errorResult = [PSCustomObject]@{
                    Path = $FolderPath
                    SizeBytes = 0
                    FileCount = 0
                    SubfolderCount = 0
                    LargestFile = $null
                    LargestFileSize = 0
                    IsAccessible = $false
                    HasCloudFiles = $false
                    Error = $_.Exception.Message
                    MaxDepthReached = $false
                }
                Write-RunspaceDebug -Message "Returning error result for $FolderPath`: $($_.Exception.Message)" -Category "ERROR_RETURN"
                Write-Output -InputObject $errorResult
                return
            }
        }
        # Launch parallel jobs
        Write-CentralLog -Message "Starting parallel job execution for $($FolderPaths.Count) directories" -Category "PARALLEL" -Source "MAIN"
        foreach ($folder in $FolderPaths) {
            $powershell = [powershell]::Create()
            $powershell.RunspacePool = $runspacePool
            $powershell.AddScript($scriptBlock).AddParameter("FolderPath", $folder).AddParameter("MaxDepth", $MaxDepth).AddParameter("EnableDebug", $DebugPreference -eq 'Continue').AddParameter("CentralLogPath", $Script:CentralLogPath) | Out-Null
            $jobs += [PSCustomObject]@{
                PowerShell = $powershell
                Handle = $powershell.BeginInvoke()
                Path = $folder
                StartTime = Get-Date
            }
            Write-CentralLog -Message "Launched job for directory: $folder" -Category "PARALLEL" -Source "MAIN"
        }
        # Collect results with enhanced progress tracking and timeout
        $completed = 0
        $timeout = (Get-Date).AddMinutes(10)  # Reduced timeout to prevent hanging
        $lastProgress = Get-Date
        $progressCheckInterval = 500  # Faster check interval for better responsiveness
        $individualJobTimeout = 180  # Individual job timeout in seconds (3 minutes)
        Write-DebugInfo -Message "Starting result collection phase with enhanced monitoring" -Category "PARALLEL_MONITOR"
        Write-CentralLog -Message "Beginning result collection phase for $($jobs.Count) jobs" -Category "PARALLEL" -Source "MAIN"
        while ($jobs | Where-Object { -not $_.Handle.IsCompleted }) {
            $completedNow = ($jobs | Where-Object { $_.Handle.IsCompleted }).Count
            if ($completedNow -gt $completed) {
                $completed = $completedNow
                $percentComplete = [math]::Round(($completed / $jobs.Count) * 100, 0)
                Write-ProgressUpdate -Activity "Parallel Processing" -Status "Completed $completed of $($jobs.Count) folders ($percentComplete%)" -PercentComplete $percentComplete
                Write-DebugInfo -Message "Progress update: $completed/$($jobs.Count) jobs completed ($percentComplete%)" -Category "PARALLEL_MONITOR"
                Write-CentralLog -Message "Progress: $completed/$($jobs.Count) jobs completed ($percentComplete%)" -Category "PROGRESS" -Source "MAIN"
                $lastProgress = Get-Date
            }
            # Enhanced real-time stream monitoring - check for new debug messages from ALL running jobs
            foreach ($job in $jobs | Where-Object { -not $_.Handle.IsCompleted }) {
                if ($job.PowerShell -and $job.PowerShell.Streams.Error.Count -gt 0) {
                    # Process any new error stream messages (including debug messages)
                    $unprocessedErrors = @($job.PowerShell.Streams.Error)
                    foreach ($errorItem in $unprocessedErrors) {
                        # Check if this is a debug message disguised as an error - more flexible pattern matching
                        if ($errorItem.Exception.Message -match "RUNSPACE_DEBUG:") {
                            $debugMsg = $errorItem.Exception.Message -replace "^.*RUNSPACE_DEBUG: ", ""
                            # Write directly to the host to ensure it appears immediately
                            $host.UI.WriteLine("$($Script:Colors.Magenta)[$($job.Path)] $debugMsg$($Script:Colors.Reset)")
                            # Also write to transcript using Write-Output for better capture
                            Write-Output "[$($job.Path)] $debugMsg"
                        }
                        else {
                            # This is a real error, display it in red
                            $host.UI.WriteLine("$($Script:Colors.Red)[$($job.Path)] ERROR: $($errorItem.Exception.Message)$($Script:Colors.Reset)")
                            Write-Output "[$($job.Path)] ERROR: $($errorItem.Exception.Message)"
                        }
                    }
                    # Clear all error messages to free memory
                    $job.PowerShell.Streams.Error.Clear()
                }
            }

            # Check for timeout
            if ((Get-Date) -gt $timeout) {
                Write-Output "$($Script:Colors.Red)Error: Processing timeout reached after 15 minutes. Forcing completion.$($Script:Colors.Reset)"
                $incompleteJobs = $jobs | Where-Object { -not $_.Handle.IsCompleted }
                Write-DebugInfo -Message "Timeout reached. $($incompleteJobs.Count) jobs still incomplete" -Category "PARALLEL_MONITOR"
                break
            }            # Enhanced stuck job detection - check if no progress for 30 seconds
            if ((Get-Date).Subtract($lastProgress).TotalSeconds -gt 30) {
                Write-Output "$($Script:Colors.Yellow)Warning: No progress for 30 seconds. Analyzing job states...$($Script:Colors.Reset)"
                $stuckJobs = $jobs | Where-Object { -not $_.Handle.IsCompleted }
                Write-DebugInfo -Message "Stuck job analysis: $($stuckJobs.Count) jobs not completed" -Category "PARALLEL_MONITOR"
                # Force timeout for jobs running longer than the individual timeout
                $longRunningJobs = @()
                $currentTime = Get-Date
                foreach ($job in $stuckJobs) {
                    if ($job.PowerShell -and $job.PowerShell.InvocationStateInfo.State -eq 'Running') {
                        # Check if job has been running longer than the timeout
                        if ($job.StartTime -and ($currentTime.Subtract($job.StartTime).TotalSeconds -gt $individualJobTimeout)) {
                            $longRunningJobs += $job
                        }
                    }
                }

                if ($longRunningJobs.Count -gt 0) {
                    Write-Output "$($Script:Colors.Red)Forcing completion of $($longRunningJobs.Count) long-running jobs...$($Script:Colors.Reset)"
                    foreach ($longJob in $longRunningJobs) {
                        try {
                            Write-DebugInfo -Message "Stopping long-running job for: $($longJob.Path)" -Category "JOB_TIMEOUT"
                            $longJob.PowerShell.Stop()
                            Start-Sleep -Milliseconds 100
                        }
                        catch {
                            Write-DebugInfo -Message "Failed to stop job for $($longJob.Path): $($_.Exception.Message)" -Category "JOB_TIMEOUT"
                        }
                    }
                }

                if ($stuckJobs.Count -le 5) {
                    foreach ($stuckJob in $stuckJobs) {
                        $jobState = if ($stuckJob.PowerShell) { $stuckJob.PowerShell.InvocationStateInfo.State } else { "Unknown" }
                        Write-Output "$($Script:Colors.Yellow)  Stuck job: $($stuckJob.Path) (State: $jobState)$($Script:Colors.Reset)"
                        Write-DebugInfo -Message "Stuck job details: Path=$($stuckJob.Path), State=$jobState" -Category "PARALLEL_MONITOR"
                    }
                } else {
                    Write-Output "$($Script:Colors.Yellow)  $($stuckJobs.Count) jobs appear stuck$($Script:Colors.Reset)"
                }
                $lastProgress = Get-Date  # Reset timer to avoid spam
            }

            Start-Sleep -Milliseconds $progressCheckInterval
        }
        Write-DebugInfo -Message "Parallel processing wait loop completed. Final job states being collected." -Category "PARALLEL_MONITOR"

        # DIAGNOSTIC: Add debug output to see if we reach the result collection loop
        Write-DebugInfo -Message "DIAGNOSTIC: Starting foreach loop to collect results from $($jobs.Count) jobs" -Category "DEBUG_DIAGNOSTIC"

        # Gather all results with improved error handling and stream isolation
        foreach ($job in $jobs) {
            Write-DebugInfo -Message "DIAGNOSTIC: Processing job for path: $($job.Path), IsCompleted: $($job.Handle.IsCompleted)" -Category "DEBUG_DIAGNOSTIC"
            try {
                if ($job.Handle.IsCompleted) {
                    # Capture all streams to isolate return objects from debug output
                    $result = $null
                    $errorOutput = $null
                    $informationOutput = $null
                    try {
                        # Explicitly separate the output stream from other streams
                        # EndInvoke() should ONLY return the Write-Output results, not error stream
                        # Get results from completed runspace job
                        $rawResult = $job.PowerShell.EndInvoke($job.Handle)
                        # CRITICAL FIX: Convert PSDataCollection to proper array
                        $result = @($rawResult)

                        # Debug output for troubleshooting EndInvoke results (only when debug enabled)
                        if ($DebugPreference -ne 'SilentlyContinue') {
                            Write-DebugInfo -Message "Raw result type: $($rawResult.GetType().Name)" -Category "ENDINVOKE_DEBUG"
                            if ($null -eq $result) {
                                Write-DebugInfo -Message "EndInvoke returned NULL for job '$($job.Path)'" -Category "ENDINVOKE_DEBUG"
                            } elseif ($result -is [array]) {
                                Write-DebugInfo -Message "EndInvoke returned ARRAY for job '$($job.Path)' with $($result.Count) elements" -Category "ENDINVOKE_DEBUG"
                                if ($result.Count -gt 0) {
                                    $firstElement = $result[0]
                                    Write-DebugInfo -Message "  First element type: $($firstElement.GetType().Name)" -Category "ENDINVOKE_DEBUG"
                                    if ($firstElement -is [string]) {
                                        Write-DebugInfo -Message "  First element string content: $($firstElement.Substring(0, [Math]::Min($firstElement.Length, 100)))" -Category "ENDINVOKE_DEBUG"
                                    }
                                }
                            } else {
                                Write-DebugInfo -Message "EndInvoke returned SINGLE object for job '$($job.Path)' of type: $($result.GetType().Name)" -Category "ENDINVOKE_DEBUG"
                                if ($result -is [string]) {
                                    Write-DebugInfo -Message "  Single string content: $($result.Substring(0, [Math]::Min($result.Length, 100)))" -Category "ENDINVOKE_DEBUG"
                                }
                            }
                        }

                        # Capture streams separately
                        $errorOutput = @($job.PowerShell.Streams.Error)
                        $informationOutput = @($job.PowerShell.Streams.Information)
                        $warningOutput = @($job.PowerShell.Streams.Warning)
                        $verboseOutput = @($job.PowerShell.Streams.Verbose)
                        $debugOutput = @($job.PowerShell.Streams.Debug)
                        # Log stream contents for debugging
                        $resultCount = if ($null -eq $result) { 0 } elseif ($result -is [array]) { $result.Count } else { 1 }
                        Write-DebugInfo -Message "Job '$($job.Path)' stream counts: Result=$resultCount, Error=$($errorOutput.Count), Info=$($informationOutput.Count), Warning=$($warningOutput.Count), Verbose=$($verboseOutput.Count), Debug=$($debugOutput.Count)" -Category "STREAM_DEBUG"

                        # Additional debugging - log the actual result type details
                        if ($null -ne $result) {
                            if ($result -is [array]) {
                                $firstItemType = if ($result.Count -gt 0) { $result[0].GetType().Name } else { 'N/A' }
                                Write-DebugInfo -Message "Result array details: Count=$($result.Count), FirstItemType=$firstItemType" -Category "STREAM_DEBUG"
                            } else {
                                Write-DebugInfo -Message "Single result type: $($result.GetType().Name), ToString: $($result.ToString().Substring(0, [Math]::Min($result.ToString().Length, 50)))" -Category "STREAM_DEBUG"
                            }
                        }
                    }
                    catch {
                        Write-Output "$($Script:Colors.Red)Failed to get result from job for '$($job.Path)': $($_.Exception.Message)$($Script:Colors.Reset)"
                        continue
                    }
                    Write-DebugInfo -Message "Job for '$($job.Path)' completed. Processing results..." -Category "JOB_RESULTS"
                    # Process error stream for any real errors (not debug messages) - improved filtering
                    if ($errorOutput -and $errorOutput.Count -gt 0) {
                        Write-DebugInfo -Message "Job for '$($job.Path)' produced $($errorOutput.Count) error messages" -Category "JOB_RESULTS"

                        # Enhanced error output handling
                        if ($DebugPreference -ne 'SilentlyContinue') {
                            # Categorize errors by type first
                            $debugMessages = @()
                            $realErrors = @()

                            foreach ($errorMsg in $errorOutput) {
                                # More specific debug message detection
                                if ($errorMsg.Exception.Message -match "RUNSPACE_DEBUG:" -or
                                    ($errorMsg.CategoryInfo -and $errorMsg.CategoryInfo.Activity -eq "RUNSPACE_DEBUG_MESSAGE")) {
                                    $debugMessages += $errorMsg
                                }
                                else {
                                    $realErrors += $errorMsg
                                }
                            }                            # Process debug messages with improved formatting
                            foreach ($debugMsg in $debugMessages) {
                                $formattedMsg = $debugMsg.Exception.Message -replace "^.*RUNSPACE_DEBUG: ", ""
                                # Write directly to the host to ensure it appears immediately
                                $host.UI.WriteLine("$($Script:Colors.Magenta)[$($job.Path)] $formattedMsg$($Script:Colors.Reset)")
                                # Force transcript capture by writing to Information stream AND output
                                Write-Information "[$($job.Path)] $formattedMsg" -InformationAction Continue
                            }

                            # Process real errors separately
                            foreach ($errorMsg in $realErrors) {
                                $host.UI.WriteLine("$($Script:Colors.Red)[$($job.Path)] ERROR: $($errorMsg.Exception.Message)$($Script:Colors.Reset)")
                                # Force transcript capture by writing to Information stream AND error
                                Write-Error "[$($job.Path)] ERROR: $($errorMsg.Exception.Message)" -ErrorAction Continue
                            }

                            Write-DebugInfo -Message "Processed $($debugMessages.Count) debug messages and $($realErrors.Count) real errors" -Category "JOB_RESULTS"
                        }
                    }
                    # Process Information stream for any remaining messages (should be empty now)
                    if ($informationOutput -and $informationOutput.Count -gt 0 -and $DebugPreference -ne 'SilentlyContinue') {
                        Write-DebugInfo -Message "Job for '$($job.Path)' had unexpected Information stream output: $($informationOutput.Count) messages" -Category "JOB_RESULTS"
                    }
                    if ($result) {
                        # Enhanced result analysis with detailed logging
                        Write-DebugInfo -Message "Job for '$($job.Path)' returned result. Analyzing..." -Category "JOB_RESULTS"

                        # Log detailed result information
                        if ($result -is [array]) {
                            Write-DebugInfo -Message "Result is ARRAY with $($result.Count) elements" -Category "JOB_RESULTS"
                            for ($i = 0; $i -lt [Math]::Min($result.Count, 5); $i++) {
                                $item = $result[$i]
                                Write-DebugInfo -Message "  Element $i`: Type=$($item.GetType().Name), IsString=$($item -is [string]), IsPSCustomObject=$($item -is [PSCustomObject])" -Category "RESULT_ANALYSIS"
                                if ($item -is [string]) {
                                    Write-DebugInfo -Message "    String content: $($item.Substring(0, [Math]::Min($item.Length, 100)))" -Category "RESULT_ANALYSIS"
                                } elseif ($item -is [PSCustomObject]) {
                                    $properties = $item.PSObject.Properties.Name -join ", "
                                    Write-DebugInfo -Message "    PSCustomObject properties: $properties" -Category "RESULT_ANALYSIS"
                                }
                            }
                        } else {
                            Write-DebugInfo -Message "Result is SINGLE object: Type=$($result.GetType().Name), IsString=$($result -is [string]), IsPSCustomObject=$($result -is [PSCustomObject])" -Category "JOB_RESULTS"
                            if ($result -is [string]) {
                                Write-DebugInfo -Message "  String content: $($result.Substring(0, [Math]::Min($result.Length, 100)))" -Category "RESULT_ANALYSIS"
                            } elseif ($result -is [PSCustomObject]) {
                                $properties = $result.PSObject.Properties.Name -join ", "
                                Write-DebugInfo -Message "  PSCustomObject properties: $properties" -Category "RESULT_ANALYSIS"
                            }
                        }

                        # Separate debug messages from actual results - enhanced validation
                        $validResults = @()
                        # Debug the raw result type and content
                        if ($DebugPreference -ne 'SilentlyContinue') {
                            if ($null -eq $result) {
                                Write-DebugInfo -Message "Result is NULL" -Category "JOB_RESULTS_DEBUG"
                            } elseif ($result -is [array]) {
                                Write-DebugInfo -Message "Job result type: ARRAY with $($result.Count) elements" -Category "JOB_RESULTS_DEBUG"
                            } else {
                                Write-DebugInfo -Message "Job result type: $($result.GetType().FullName)" -Category "JOB_RESULTS_DEBUG"
                            }
                        }
                        # Strictly filter for ONLY PSCustomObject results, rejecting any string values that might be debug messages incorrectly captured in the result
                        # BREAKPOINT: Set breakpoint here to examine result filtering
                        if ($result -is [array]) {
                            Write-DebugInfo -Message "Job returned array with $($result.Count) elements" -Category "JOB_RESULTS"
                            # Only accept objects that match our exact expected schema
                            $validResults = @($result | Where-Object {
                                $_ -is [PSCustomObject] -and
                                $null -ne $_ -and
                                -not ($_ -is [string]) -and
                                $_.PSObject.Properties.Name -contains 'Path' -and
                                $_.PSObject.Properties.Name -contains 'SizeBytes' -and
                                $_.PSObject.Properties.Name -contains 'FileCount' -and
                                $_.PSObject.Properties.Name -contains 'SubfolderCount' -and
                                $_.PSObject.Properties.Name -contains 'IsAccessible' -and
                                -not [string]::IsNullOrEmpty($_.Path)
                            })

                            if ($validResults.Count -lt $result.Count) {
                                Write-DebugInfo -Message "Filtered out $($result.Count - $validResults.Count) invalid results" -Category "JOB_RESULTS"

                                # More detailed logging about what got filtered
                                if ($DebugPreference -ne 'SilentlyContinue') {
                                    foreach ($item in $result) {
                                        if ($item -is [string]) {
                                            Write-DebugInfo -Message "Filtered string item: $item" -Category "FILTER_DEBUG"
                                        } elseif (-not ($item -is [PSCustomObject])) {
                                            Write-DebugInfo -Message "Filtered non-PSCustomObject: $($item.GetType().Name)" -Category "FILTER_DEBUG"
                                        }
                                    }
                                }
                            }
                        } else {
                            Write-DebugInfo -Message "Job returned single object of type: $($result.GetType().Name)" -Category "JOB_RESULTS"
                            # Enhanced validation for single object results - explicitly reject strings
                            if ($result -is [string]) {
                                # String results are always debug messages - display but don't add to results
                                if ($DebugPreference -ne 'SilentlyContinue') {
                                    Write-DebugInfo -Message "String result rejected (not added): $result" -Category "FILTER_DEBUG"
                                }
                                # Don't add strings to valid results
                            } elseif ($result -is [PSCustomObject] -and
                                     $null -ne $result -and
                                     $result.PSObject.Properties.Name -contains 'Path' -and
                                     $result.PSObject.Properties.Name -contains 'SizeBytes' -and
                                     $result.PSObject.Properties.Name -contains 'FileCount' -and
                                     $result.PSObject.Properties.Name -contains 'SubfolderCount' -and
                                     $result.PSObject.Properties.Name -contains 'IsAccessible' -and
                                     -not [string]::IsNullOrEmpty($result.Path)) {
                                $validResults = @($result)
                                Write-DebugInfo -Message "Added valid PSCustomObject result for path: $($result.Path)" -Category "JOB_RESULTS"
                            } else {
                                Write-DebugInfo -Message "Received invalid result object, skipping" -Category "JOB_RESULTS"
                            }
                        }
                        if ($validResults.Count -gt 0) {
                            # Extra validation to ensure we're only adding PSCustomObject instances
                            $cleanValidResults = @($validResults | Where-Object {
                                $_ -is [PSCustomObject] -and -not ($_ -is [string])
                            })

                            if ($cleanValidResults.Count -ne $validResults.Count) {
                                Write-DebugInfo -Message "Extra filtering removed $($validResults.Count - $cleanValidResults.Count) invalid objects" -Category "FINAL_FILTER"
                            }

                            if ($cleanValidResults.Count -gt 0) {
                                $results += $cleanValidResults
                                Write-DebugInfo -Message "Added $($cleanValidResults.Count) valid results from job" -Category "JOB_RESULTS"
                                # Process each valid result for error tracking
                                foreach ($validResult in $cleanValidResults) {
                                    if (-not $validResult.IsAccessible) {
                                        $Script:InaccessibleFolderCount++
                                        if (-not [string]::IsNullOrWhiteSpace($validResult.Error)) {
                                            $Script:ErrorTracker[$validResult.Path] = $validResult.Error
                                            Write-DebugInfo -Message "Collected error: $($validResult.Path) - $($validResult.Error)" -Category "ERROR_COLLECTION"
                                        }
                                    }

                                    # Check if max depth was reached in any result
                                    if ($validResult.PSObject.Properties.Name -contains 'MaxDepthReached' -and $validResult.MaxDepthReached) {
                                        $Script:MaxDepthReached = $true
                                        Write-DebugInfo -Message "Max depth reached detected from: $($validResult.Path)" -Category "MAX_DEPTH_TRACKING"
                                    }
                                }
                            } else {
                                Write-DebugInfo -Message "Job for '$($job.Path)' produced no clean valid results after final filtering" -Category "JOB_RESULTS"
                            }
                        } else {
                            Write-DebugInfo -Message "Job for '$($job.Path)' produced no valid results - filtering out contaminated data" -Category "JOB_RESULTS"
                        }

                        $Script:ProcessedFolders++
                    } else {
                        Write-DebugInfo -Message "Job for '$($job.Path)' returned null result" -Category "JOB_RESULTS"
                    }
                } else {
                    Write-Output "$($Script:Colors.Yellow)Warning: Job for $($job.Path) did not complete. Skipping.$($Script:Colors.Reset)"
                    # Try to stop the incomplete job safely
                    try {
                        if ($job.PowerShell -and $job.PowerShell.InvocationStateInfo.State -eq 'Running') {
                            $job.PowerShell.Stop()
                        }
                    } catch {
                        Write-DebugInfo -Message "Unable to stop PowerShell job for '$($job.Path)': $($_.Exception.Message)" -Category "JOB_CLEANUP"
                    }
                }
            }
            catch {
                Write-Output "$($Script:Colors.Red)Error processing job for $($job.Path): $($_.Exception.Message)$($Script:Colors.Reset)"
                Write-DebugInfo -Message "Job processing error details: $($_.Exception | Out-String)" -Category "JOB_ERROR"
            }
            finally {
                # Safe disposal with null checks
                try {
                    if ($job.PowerShell) {
                        $job.PowerShell.Dispose()
                    }
                } catch {
                    Write-DebugInfo -Message "Unable to dispose PowerShell job for '$($job.Path)': $($_.Exception.Message)" -Category "JOB_CLEANUP"
                }
            }
        }
        # Enhanced runspace cleanup with detailed logging
        Write-DebugInfo -Message "Starting runspace pool cleanup" -Category "CLEANUP"
        try {
            if ($runspacePool) {
                if ($runspacePool.RunspacePoolStateInfo.State -eq 'Opened') {
                    Write-DebugInfo -Message "Closing runspace pool..." -Category "CLEANUP"
                    $runspacePool.Close()
                }
                Write-DebugInfo -Message "Disposing runspace pool..." -Category "CLEANUP"
                $runspacePool.Dispose()
                Write-DebugInfo -Message "Runspace pool cleanup completed successfully" -Category "CLEANUP"
            }
        }
        catch {
            Write-Output "$($Script:Colors.Yellow)Warning: Error during runspace cleanup: $($_.Exception.Message)$($Script:Colors.Reset)"
            Write-DebugInfo -Message "Runspace cleanup error details: $($_.Exception | Out-String)" -Category "CLEANUP_ERROR"
        }

        # Enhanced memory cleanup
        Write-DebugInfo -Message "Starting memory cleanup and garbage collection" -Category "CLEANUP"
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
        Write-DebugInfo -Message "Memory cleanup completed" -Category "CLEANUP"
        # Debug logging for error collection summary
        Write-DebugInfo -Message "Parallel processing completed. Errors collected: $($Script:ErrorTracker.Count)" -Category "PARALLEL_SUMMARY"

        # Count all inaccessible directories recursively in results
        # Define helper function inside try block to keep proper scope
        function Get-InaccessibleDirCount {
            param($ResultObject)
            $count = 0
            if ($null -ne $ResultObject -and -not $ResultObject.IsAccessible) {
                $count++
                Write-DebugInfo -Message "Found inaccessible directory in results: $($ResultObject.Path)" -Category "INACCESSIBLE_COUNT"
            }
            # Note: We only count top-level directories here since subdirectories are handled by the parallel jobs themselves
            return $count
        }

        # Process results and count inaccessible folders
        $totalInaccessibleFound = 0
        if ($null -ne $results -and $results.Count -gt 0) {
            foreach ($result in $results) {
                if ($null -ne $result -and -not $result.IsAccessible) {
                    $totalInaccessibleFound++
                    Write-DebugInfo -Message "Found inaccessible directory in results: $($result.Path)" -Category "INACCESSIBLE_COUNT"
                }
            }
        }

        # Log final counts for comparison and validation
        Write-DebugInfo -Message "Total inaccessible directories found in results: $totalInaccessibleFound" -Category "PARALLEL_SUMMARY"
        Write-DebugInfo -Message "Script counter shows: $Script:InaccessibleFolderCount inaccessible folders" -Category "PARALLEL_SUMMARY"
    }
    catch {
        Write-Output "$($Script:Colors.Red)Error during parallel processing: $($_.Exception.Message)$($Script:Colors.Reset)"
        return @()  # Return empty array on error
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
    }    # Level 1: Show top-level subfolders
    Write-Output "`n$($Script:Colors.Bold)$($Script:Colors.Cyan)LEVEL 1: TOP SUBFOLDERS OF $StartPath$($Script:Colors.Reset)"
    Write-DebugInfo -Message "Searching for Level 1 folders in results array" -Category "HIERARCHY"
    Write-DebugInfo -Message "Total results count: $($Results.Count)" -Category "HIERARCHY"
    Write-DebugInfo -Message "StartPath for comparison: '$StartPath'" -Category "HIERARCHY"    # Normalize the StartPath to ensure consistent comparison
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
    }    $allLevel1Candidates = $Results | Where-Object {
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

    # Debug logging for accessibility breakdown
    Write-DebugInfo -Message "Accessibility summary from results:" -Category "ACCESSIBILITY"
    Write-DebugInfo -Message "  Total results: $($Results.Count)" -Category "ACCESSIBILITY"
    Write-DebugInfo -Message "  Accessible results: $($accessibleResults.Count)" -Category "ACCESSIBILITY"
    Write-DebugInfo -Message "  Inaccessible results: $inaccessibleCount" -Category "ACCESSIBILITY"
    Write-DebugInfo -Message "  Error tracker entries: $($Script:ErrorTracker.Count)" -Category "ACCESSIBILITY"
      if ($DebugPreference -eq 'Continue') {
        $inaccessibleResults = $Results | Where-Object { -not $_.IsAccessible }
        if ($inaccessibleResults.Count -gt 0) {
            Write-DebugInfo -Message "Sample inaccessible directories:" -Category "ACCESSIBILITY"
            $sampleErrors = $inaccessibleResults | Select-Object -First 10
            foreach ($errorResult in $sampleErrors) {
                Write-DebugInfo -Message "  Path: '$($errorResult.Path)' | Error: '$($errorResult.Error)' | Type: $($errorResult.GetType().Name) | IsAccessible: $($errorResult.IsAccessible)" -Category "ACCESSIBILITY"
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
                Write-DebugInfo -Message "  $($type.Name): $($type.Count)" -Category "ACCESSIBILITY"
            }
        }
    }

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
                    $discrepancy = $driveUsedSpace - $totalSize
                    $discrepancyPercent = if ($driveUsedSpace -gt 0) { [math]::Round(($discrepancy / $driveUsedSpace) * 100, 1) } else { 0 }

                    Write-Output "`n$($Script:Colors.Bold)DRIVE USAGE COMPARISON$($Script:Colors.Reset)"
                    Write-Output "$($Script:Colors.White)Drive Used Space: $($Script:Colors.Cyan)$(Format-FileSize -SizeInBytes $driveUsedSpace)$($Script:Colors.Reset)"
                    Write-Output "$($Script:Colors.White)Script Calculated: $($Script:Colors.Cyan)$(Format-FileSize -SizeInBytes $totalSize)$($Script:Colors.Reset)"

                    if ($discrepancy -gt 0) {
                        Write-Output "$($Script:Colors.White)Unaccounted Space: $($Script:Colors.Yellow)$(Format-FileSize -SizeInBytes $discrepancy) ($discrepancyPercent%)$($Script:Colors.Reset)"

                        # Get NTFS overhead estimation
                        $ntfsOverhead = Get-NTFSOverhead -DriveLetter $driveLetter
                        if ($ntfsOverhead.TotalOverhead -gt 0) {
                            Write-Output "`n$($Script:Colors.Bold)NTFS SYSTEM OVERHEAD ANALYSIS$($Script:Colors.Reset)"
                            Write-Output "$($Script:Colors.White)MFT Size: $($Script:Colors.Cyan)$(Format-FileSize -SizeInBytes $ntfsOverhead.MFTSize)$($Script:Colors.Reset)"
                            if ($ntfsOverhead.JournalSize -gt 0) {
                                Write-Output "$($Script:Colors.White)Journal Size: $($Script:Colors.Cyan)$(Format-FileSize -SizeInBytes $ntfsOverhead.JournalSize)$($Script:Colors.Reset)"
                            }
                            Write-Output "$($Script:Colors.White)Total NTFS Overhead: $($Script:Colors.Cyan)$(Format-FileSize -SizeInBytes $ntfsOverhead.TotalOverhead)$($Script:Colors.Reset)"
                            Write-Output "$($Script:Colors.DarkGray)Method: $($ntfsOverhead.EstimationMethod)$($Script:Colors.Reset)"

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
                            Write-Output "$($Script:Colors.DarkGray)  - Virtual memory files and hibernation data$($Script:Colors.Reset)"

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
    }
    catch {
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
    }
    catch {
        $result.Error = "Access Denied: $($_.Exception.Message)"
    }

    return $result
}

function Get-NTFSOverhead {
    <#
    .SYNOPSIS
    Attempts to estimate NTFS file system overhead including MFT, journals, and reserved space.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$DriveLetter
    )

    try {
        $overhead = [PSCustomObject]@{
            MFTSize = 0
            JournalSize = 0
            SystemReserved = 0
            TotalOverhead = 0
            EstimationMethod = "Unknown"
        }

        # Try to get MFT size using fsutil
        try {
            $fsutilOutput = & fsutil fsinfo ntfsinfo "${DriveLetter}:" 2>$null
            if ($fsutilOutput) {
                foreach ($line in $fsutilOutput) {
                    if ($line -match "Mft Valid Data Length\s*:\s*0x([0-9A-Fa-f]+)") {
                        $overhead.MFTSize = [Convert]::ToInt64($matches[1], 16)
                        $overhead.EstimationMethod = "fsutil"
                    }
                    elseif ($line -match "Usn Journal Size\s*:\s*0x([0-9A-Fa-f]+)") {
                        $overhead.JournalSize = [Convert]::ToInt64($matches[1], 16)
                    }
                }
            }
        }
        catch {
            Write-Verbose "Could not get NTFS info via fsutil: $($_.Exception.Message)"
        }

        # Fallback: Estimate based on drive size (typical NTFS overhead is 1-3% of drive size)
        if ($overhead.MFTSize -eq 0) {
            try {
                $volume = Get-Volume -DriveLetter $DriveLetter -ErrorAction Stop
                $estimatedOverhead = [math]::Round($volume.Size * 0.02, 0)  # 2% estimate
                $overhead.TotalOverhead = $estimatedOverhead
                $overhead.EstimationMethod = "Percentage estimate (2%)"
            }
            catch {
                $overhead.EstimationMethod = "Unable to estimate"
            }
        } else {
            $overhead.TotalOverhead = $overhead.MFTSize + $overhead.JournalSize
        }

        return $overhead
    }
    catch {
        return [PSCustomObject]@{
            MFTSize = 0
            JournalSize = 0
            SystemReserved = 0
            TotalOverhead = 0
            EstimationMethod = "Error: $($_.Exception.Message)"
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
                    $parentSize = try { (Get-ChildItem -Path (Split-Path $path -Parent) -Directory -Force | Where-Object { $_.Name -ne (Split-Path $path -Leaf) } | Measure-Object Length -Sum).Sum } catch { 0 }
                    $pathEstimate = [math]::Max(1MB, $parentSize * 0.1)  # Conservative estimate
                    $method = "Parent directory pattern"
                }
            }
            catch {
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
                }
                catch {
                    Write-Verbose "Could not query WMI for directory $path"
                }
            }

            # Method 3: Check if it's a known system directory and use typical sizes
            if ($pathEstimate -eq 0) {
                $dirName = Split-Path $path -Leaf
                switch -Regex ($dirName) {
                    "^Users$" {
                        $pathEstimate = 50GB  # Users directories can be very large
                        $method = "Known directory type estimate (Users)"
                    }
                    "^Windows$" {
                        $pathEstimate = 25GB  # Windows directory typical size
                        $method = "Known directory type estimate (Windows)"
                    }
                    "^WinSxS$" {
                        $pathEstimate = 15GB  # WinSxS can be very large
                        $method = "Known directory type estimate (WinSxS)"
                    }
                    "^System32$" {
                        $pathEstimate = 5GB   # System32 typical size
                        $method = "Known directory type estimate (System32)"
                    }
                    "^Temp" {
                        $pathEstimate = 2GB   # Temp directories
                        $method = "Known directory type estimate (Temp)"
                    }
                    default {
                        $pathEstimate = 100MB  # Conservative default
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
        }
        catch {
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
        Write-Output "$($Script:Colors.Bold)$($Script:Colors.Cyan)    ULTRA-FAST FOLDER USAGE ANALYZER v2.7.0    $($Script:Colors.Reset)"
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

        # Extract and display drive information for the target drive
        try {
            $driveLetter = [System.IO.Path]::GetPathRoot($StartPath).TrimEnd('\').TrimEnd(':')
            if ($driveLetter -and $driveLetter.Length -eq 1) {
                Show-DriveInfo -DriveLetter $driveLetter
            }
        }
        catch {
            Write-Output "$($Script:Colors.Yellow)Warning: Could not extract drive information from path: $StartPath$($Script:Colors.Reset)"
        }

        Write-Output ""

        if ($PSCmdlet.ShouldProcess($StartPath, "Analyze folder usage")) {            # Get top-level directories for parallel processing
            Write-ProgressUpdate -Activity "Initialization" -Status "Scanning top-level directories..." -Color "Cyan"            # Known problematic directories that can cause hangs
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

                # Check for system files at root level
                try {
                    $rootItems = Get-ChildItem -Path $StartPath -File -Force -ErrorAction SilentlyContinue
                    foreach ($file in $rootItems) {
                        if ($systemFiles -contains $file.Name) {
                            $systemFilesFound += $file
                        }
                    }
                }
                catch {
                    Write-Output "$($Script:Colors.Yellow)Warning: Could not check for system files: $($_.Exception.Message)$($Script:Colors.Reset)"
                }
                $Script:TotalFolders = $topLevelDirs.Count

                Write-Output "$($Script:Colors.Green)Found $($topLevelDirs.Count) top-level directories to analyze$($Script:Colors.Reset)"

                # Debug logging for directory discovery
                Write-DebugInfo -Message "Total directories found in $StartPath`: $($items.Count)" -Category "DISCOVERY"
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
            }
            catch {
                Write-Output "$($Script:Colors.Red)Error accessing start path: $($_.Exception.Message)$($Script:Colors.Reset)"
                return
            }

            if ($topLevelDirs.Count -eq 0) {
                Write-Output "$($Script:Colors.Yellow)No directories found to analyze.$($Script:Colors.Reset)"
                return
            }            # Perform parallel analysis
            Write-DetailedProgress -Activity "Analysis" -Status "Starting parallel folder analysis..." -Color "Cyan" -Detail "Processing $($topLevelDirs.Count) directories with $MaxThreads threads, MaxDepth=$MaxDepth"

            Write-DebugInfo -Message "Starting parallel analysis with parameters:" -Category "PARALLEL_START"
            Write-DebugInfo -Message "  Directories to process: $($topLevelDirs.Count)" -Category "PARALLEL_START"
            Write-DebugInfo -Message "  Max Depth: $MaxDepth" -Category "PARALLEL_START"
            Write-DebugInfo -Message "  Max Threads: $MaxThreads" -Category "PARALLEL_START"

            $results = Get-ParallelFolderStatistic -FolderPaths $topLevelDirs -MaxDepth $MaxDepth -MaxThreads $MaxThreads

            Write-DebugInfo -Message "Parallel analysis completed. Results count: $($results.Count)" -Category "PARALLEL_COMPLETE"

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
                    $Script:ProcessedFolders++  # Count problematic directories as processed
                }
            }            # Analyze the StartPath directory itself for files at the root level
            Write-ProgressUpdate -Activity "Analysis" -Status "Analyzing root directory files..." -Color "Cyan"
            try {
                $rootFiles = Get-ChildItem -Path $StartPath -File -Force -ErrorAction Stop

                # Calculate cumulative size: subdirectories + root files + problematic directories
                $subdirectoryTotalSize = ($results | Where-Object { $_.IsAccessible } | Measure-Object -Property SizeBytes -Sum).Sum
                $rootFilesSize = if ($rootFiles.Count -gt 0) { ($rootFiles | Measure-Object -Property Length -Sum).Sum } else { 0 }

                # Add problematic directories size (safely scanned)
                $problematicDirsSize = if ($safeResults.Count -gt 0) { ($safeResults | Measure-Object -Property SizeBytes -Sum).Sum } else { 0 }                # Calculate total cumulative size including all components
                $totalCumulativeSize = $subdirectoryTotalSize + $rootFilesSize + $problematicDirsSize

                # Debug logging for root calculation
                Write-DebugInfo -Message "Root calculation breakdown:" -Category "ROOT_CALC"
                Write-DebugInfo -Message "  Subdirectory total size: $(Format-FileSize -SizeInBytes $subdirectoryTotalSize)" -Category "ROOT_CALC"
                Write-DebugInfo -Message "  Root files size: $(Format-FileSize -SizeInBytes $rootFilesSize) ($($rootFiles.Count) files)" -Category "ROOT_CALC"
                Write-DebugInfo -Message "  Problematic dirs size: $(Format-FileSize -SizeInBytes $problematicDirsSize)" -Category "ROOT_CALC"
                Write-DebugInfo -Message "  TOTAL cumulative size: $(Format-FileSize -SizeInBytes $totalCumulativeSize)" -Category "ROOT_CALC"

                if ($DebugPreference -ne 'SilentlyContinue' -and $rootFiles.Count -gt 0) {
                    Write-DebugInfo -Message "Root files found:" -Category "ROOT_FILES"
                    $sortedRootFiles = $rootFiles | Sort-Object Length -Descending | Select-Object -First 10
                    foreach ($file in $sortedRootFiles) {
                        $fileSize = Format-FileSize -SizeInBytes $file.Length
                        Write-DebugInfo -Message "  $($file.Name): $fileSize" -Category "ROOT_FILES"
                    }
                }# Calculate cumulative file count: subdirectories + root files + problematic directories
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
                    SizeBytes = $totalCumulativeSize
                    FileCount = $totalCumulativeFiles
                    SubfolderCount = $totalSubfolderCount
                    LargestFile = $overallLargestFile
                    LargestFileSize = $overallLargestFileSize
                    IsAccessible = $true
                    HasCloudFiles = $overallHasCloudFiles
                    Error = $null
                }

                # Replace or add the root stats at the beginning of results
                $results = @($rootStats) + $results
                $Script:ProcessedFolders++
            }
            catch {
                Write-Output "$($Script:Colors.Yellow)Warning: Could not analyze root directory files: $($_.Exception.Message)$($Script:Colors.Reset)"
            }
            # Secondary analysis: Collect direct subfolders of the largest directories for hierarchical display
            Write-ProgressUpdate -Activity "Analysis" -Status "Collecting subfolder details for hierarchical display..." -Color "Cyan"
            $additionalResults = @()

            # Add the additional results to the main results for hierarchical display
            $results += $additionalResults            # Display results using hierarchical view
            # FINAL CLEANUP: Remove any invalid results that might have been added
            $cleanResults = $results | Where-Object {
                $_ -is [PSCustomObject] -and
                $_.PSObject.Properties.Name -contains 'Path' -and
                $_.PSObject.Properties.Name -contains 'IsAccessible' -and
                -not [string]::IsNullOrEmpty($_.Path)
            }

            Write-DebugInfo -Message "FINAL CLEANUP: Original results: $($results.Count), Clean results: $($cleanResults.Count)" -Category "CLEANUP"
            $results = $cleanResults

            Show-HierarchicalResult -Results $results -StartPath $StartPath -Top $Top -SafeResults $safeResults
            Show-ErrorSummary

            Write-Output "`n$($Script:Colors.Green)Analysis completed successfully!$($Script:Colors.Reset)"
        }
        else {
            Write-Output "$($Script:Colors.Yellow)WhatIf: Would analyze folder usage starting from '$StartPath'$($Script:Colors.Reset)"
            Write-Output "$($Script:Colors.Yellow)WhatIf: Would scan up to $MaxDepth levels deep$($Script:Colors.Reset)"
            Write-Output "$($Script:Colors.Yellow)WhatIf: Would display top $Top largest folders$($Script:Colors.Reset)"
            Write-Output "$($Script:Colors.Yellow)WhatIf: Would use $MaxThreads parallel threads$($Script:Colors.Reset)"
        }
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
