# =============================================================================
# Script: Clear-SystemStorage.ps1
# Created: 2025-02-27 18:55:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-11 20:11:00 UTC
# Updated By: jdyer-nuvodia
# Version: 4.2.0
# Additional Info: Added silent cleanup configuration for cleanmgr.exe
# =============================================================================

<#
.SYNOPSIS
    Cleans up system storage by removing temporary files and managing shadow copies.
.DESCRIPTION
    This script performs system storage cleanup operations to free up disk space by:
     - Running Windows Disk Cleanup utility silently
     - Managing and reducing Volume Shadow Copies
     - Displaying drive space information before and after cleanup
     - Always elevates to SYSTEM context for thorough cleanup
.EXAMPLE
    .\Clear-SystemStorage.ps1
    Runs the script with default settings, elevating to SYSTEM context.
.NOTES
    Security Level: Medium
    Required Permissions: Administrative access for full functionality
    Validation Requirements: Verify disk space is freed after execution
#>

# ----- Global Variable Initialization -----
# Set debug preference to continue for detailed output
$DebugPreference = "Continue"
$VerbosePreference = "Continue"

$scriptPath = $MyInvocation.PSCommandPath
if (-not $scriptPath) {
    # Fallback for direct console execution
    $scriptPath = $MyInvocation.MyCommand.Definition
    if (-not $scriptPath) {
        $scriptPath = Join-Path -Path (Get-Location) -ChildPath "Clear-SystemStorage.ps1"
    }
}
$script:OriginalScriptDirectory = Split-Path -Path $scriptPath -Parent
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$computerName = $env:COMPUTERNAME
$script:LogFile = Join-Path -Path $script:OriginalScriptDirectory -ChildPath "ClearSystemStorage_${computerName}_$timestamp.log"

# ----- Function: Write-Log -----
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter()]
        [ValidateSet('Info','Warning','Error','Debug','Verbose')]
        [string]$Level = 'Info'
    )
    
    $timestampMsg = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestampMsg] [$Level] $Message"
    
    # Always write to console immediately with color coding for clear visibility.
    switch ($Level) {
        'Info' { Write-Host "$timestampMsg - INFO: $Message" -ForegroundColor White }
        'Warning' { Write-Host "$timestampMsg - WARNING: $Message" -ForegroundColor Yellow }
        'Error' { Write-Host "$timestampMsg - ERROR: $Message" -ForegroundColor Red }
        'Debug' { Write-Host "$timestampMsg - DEBUG: $Message" -ForegroundColor Magenta }
        'Verbose' { Write-Host "$timestampMsg - VERBOSE: $Message" -ForegroundColor Cyan }
    }
    
    # Always append the log entry to the log file.
    Add-Content -Path $script:LogFile -Value $logEntry
}

# Log script start with header
Write-Log "===== SCRIPT EXECUTION STARTED =====" -Level Info
Write-Log "Script version: 4.1.1" -Level Info
Write-Log "Computer Name: $computerName" -Level Info
Write-Log "Log file: $script:LogFile" -Level Info
Write-Log "Running as user: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)" -Level Info

# ----- Function: Test-RunningAsSystem -----
function Test-RunningAsSystem {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $isSystem = $currentUser.User.Value -eq "S-1-5-18"
    Write-Log -Message "Checking if running as SYSTEM: $isSystem" -Level Verbose
    return $isSystem
}

# ----- Function: Start-SystemContext -----
function Start-SystemContext {
    param(
        [string]$ScriptPath = $MyInvocation.PSCommandPath,
        [int]$TimeoutSeconds = 900  # Decreased to 15 minutes to ensure timely completion
    )
    
    if (Test-RunningAsSystem) {
        Write-Log "Already running as SYSTEM account." -Level Info
        return $true
    }
    
    Write-Log "Elevating to SYSTEM context..." -Level Info
    Write-Host "`nNOTE: Script will continue execution as SYSTEM. Watch for shadow copy details below." -ForegroundColor Cyan
    
    try {
        $jobName = "SystemContextJob_$([Guid]::NewGuid())"
        Write-Log "Created job with name: $jobName" -Level Debug
        
        # Use more reliable temp path creation
        $systemAccessibleTemp = [System.IO.Path]::Combine($env:SystemRoot, "Temp")
        $scriptFileName = Split-Path -Leaf $ScriptPath
        $systemAccessibleScriptPath = [System.IO.Path]::Combine($systemAccessibleTemp, $scriptFileName)
        
        # Ensure cleanup of any existing files
        $filesToClean = @(
            $systemAccessibleScriptPath,
            "$markerFile.status",
            $executorScript,
            $logFile
        )
        
        foreach ($file in $filesToClean) {
            if ([System.IO.File]::Exists($file)) {
                try {
                    [System.IO.File]::Delete($file)
                    Write-Log "Cleaned up existing file: $file" -Level Debug
                }
                catch {
                    Write-Log "Warning: Could not delete existing file $file : $($_.Exception.Message)" -Level Warning
                }
            }
        }
        
        Write-Log "Copying script to system-accessible location: $systemAccessibleScriptPath" -Level Verbose
        
        # Use Copy-Item instead of Get-Content/Set-Content to preserve encoding
        Copy-Item -Path $ScriptPath -Destination $systemAccessibleScriptPath -Force
        
        # Verify script integrity after copying
        if (Test-Path $systemAccessibleScriptPath) {
            $fileSize = (Get-Item $systemAccessibleScriptPath).Length
            Write-Log "Script copied successfully. File size: $fileSize bytes" -Level Verbose
        } else {
            Write-Log "Failed to copy script to system location!" -Level Warning
            
            # Fallback to Get-Content/Set-Content with explicit encoding
            try {
                Get-Content -Path $ScriptPath -Raw -Encoding UTF8 | 
                    Set-Content -Path $systemAccessibleScriptPath -Encoding UTF8 -Force
                if (Test-Path $systemAccessibleScriptPath) {
                    Write-Log "Script copied via content method." -Level Info
                }
            }
            catch {
                Write-Log "All copy attempts failed: $($_.Exception.Message)" -Level Error
                return $false
            }
        }
        
        # Create an executor script that will run our main script as SYSTEM.
        $executorScript = Join-Path -Path $systemAccessibleTemp -ChildPath "$jobName-executor.ps1"
        $logFile = Join-Path $systemAccessibleTemp "$jobName.log"
        $markerFile = Join-Path $systemAccessibleTemp "$jobName.marker"
        
        Write-Log "Creating executor with logfile: $logFile and marker: $markerFile" -Level Debug
        
        $executorContent = @"
# Executor script for Clear-SystemStorage running as SYSTEM
`$ErrorActionPreference = 'Stop'
`$DebugPreference = 'Continue'
`$VerbosePreference = 'Continue'
Start-Transcript -Path '$logFile' -Force

function Write-StatusFile {
    param([string]`$Status)
    Set-Content -Path '$markerFile.status' -Value `$Status -Force
}

try {
    Write-StatusFile "Starting system cleanup process..."
    
    # Run cleanup with detailed output
    Write-StatusFile "Configuring cleanup settings..."
    `$sagesetProcess = Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sageset:1" -Wait -PassThru
    Write-StatusFile "Configuration completed. Starting cleanup..."
    
    `$cleanmgrProcess = Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sagerun:1" -PassThru -NoNewWindow
    `$processId = `$cleanmgrProcess.Id
    
    while (!`$cleanmgrProcess.HasExited) {
        try {
            `$process = Get-Process -Id `$processId -ErrorAction Stop
            `$wmi = Get-WmiObject -Class Win32_Process -Filter "ProcessId = `$processId"
            
            if (`$wmi) {
                `$cpuCounter = Get-Counter -Counter "\Process(cleanmgr)\% Processor Time" -ErrorAction SilentlyContinue
                `$cpuUsage = if (`$cpuCounter) {
                    [math]::Round(`$cpuCounter.CounterSamples[0].CookedValue, 1)
                } else { 0 }
                
                `$memoryMB = [math]::Round(`$process.WorkingSet64 / 1MB, 2)
                `$runtime = (Get-Date) - `$startTime
                
                `$status = @{
                    Runtime = `$runtime.ToString('mm\:ss')
                    CPU = `$cpuUsage
                    Memory = `$memoryMB
                    Path = `$wmi.CommandLine
                } | ConvertTo-Json
                
                Write-StatusFile `$status
            }
            
            Start-Sleep -Seconds 1
        }
        catch {
            Write-StatusFile "Error: `$(`$_.Exception.Message)"
            break
        }
    }
    
    # Continue with main script
    & '$systemAccessibleScriptPath' -Debug -Verbose *>> '$logFile'
    Write-StatusFile "Complete"
}
catch {
    Write-StatusFile "Error: `$(`$_.Exception.Message)"
    exit 1
}
"@
        Write-Log "Creating executor script at $executorScript" -Level Verbose
        Set-Content -Path $executorScript -Value $executorContent -Encoding UTF8 -Force
        
        $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$executorScript`""
        $principal = New-ScheduledTaskPrincipal -UserId "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
        $task = New-ScheduledTask -Action $action -Principal $principal -Settings $settings
        
        Write-Log "Registering scheduled task $jobName" -Level Verbose
        Register-ScheduledTask -TaskName $jobName -InputObject $task | Out-Null
        Start-ScheduledTask -TaskName $jobName
        Write-Log "Started scheduled task as SYSTEM" -Level Debug
        
        Write-Log "Waiting for system cleanup to complete..." -Level Info
        Write-Host "`n===== System Cleanup Progress =====" -ForegroundColor Cyan
        Write-Host "Task running as SYSTEM. Monitoring cleanup progress..." -ForegroundColor Cyan
        
        $statusFile = Join-Path -Path $systemAccessibleTemp -ChildPath "$jobName.status"
        $completed = $false

        # Initialize monitoring variables
        $maxRetries = 3
        $retryDelay = 2
        $maxMonitoringTime = 300 # 5 minutes
        $statusCheckInterval = 500 # milliseconds
        $completed = $false
        $iterationCount = 0

        # Create and initialize status file before monitoring
        $statusFile = Join-Path -Path $systemAccessibleTemp -ChildPath "$jobName.status"
        [System.IO.File]::WriteAllText($statusFile, "Initializing cleanup process...")

        try {
            $completed = $false
            $iterationCount = 0

            while ($iterationCount -lt $maxRetries -and -not $completed) {
                try {
                    Write-Log "Starting cleanup iteration $($iterationCount + 1)..." -Level Info
                    
                    # Update status file for current iteration
                    [System.IO.File]::WriteAllText($statusFile, "Running cleanup iteration $($iterationCount + 1)...")
                    
                    # Run cleanup utilities
                    Write-Log "Executing disk cleanup utilities..." -Level Info
                    $sagesetProcess = Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sageset:1" -Wait -PassThru
                    
                    if ($sagesetProcess.ExitCode -eq 0) {
                        Write-Log "Configuration completed successfully. Starting cleanup..." -Level Info
                        [System.IO.File]::WriteAllText($statusFile, "Running disk cleanup...")
                        
                        $cleanmgrProcess = Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sagerun:1" -PassThru -NoNewWindow
                        $processStartTime = Get-Date
                        
                        # Monitor cleanup process
                        while (!$cleanmgrProcess.HasExited) {
                            if ((Get-Date) -gt $processStartTime.AddSeconds($maxMonitoringTime)) {
                                Write-Log "Cleanup process timeout reached" -Level Warning
                                [System.IO.File]::WriteAllText($statusFile, "Error: Cleanup process timed out")
                                Stop-Process -Id $cleanmgrProcess.Id -Force -ErrorAction SilentlyContinue
                                break
                            }
                            
                            $status = [System.IO.File]::ReadAllText($statusFile)
                            Write-Host "`r$(' ' * 80)" -NoNewline
                            Write-Host "`r$status" -NoNewline -ForegroundColor Cyan
                            Start-Sleep -Milliseconds $statusCheckInterval
                        }
                        
                        if ($cleanmgrProcess.ExitCode -eq 0) {
                            Write-Log "Cleanup completed successfully" -Level Info
                            [System.IO.File]::WriteAllText($statusFile, "Complete")
                            $completed = $true
                            break
                        }
                    } else {
                        Write-Log "Cleanup configuration failed with exit code: $($sagesetProcess.ExitCode)" -Level Error
                        [System.IO.File]::WriteAllText($statusFile, "Error: Configuration failed")
                    }
                }
                catch {
                    $errorMsg = "Error in iteration $($iterationCount + 1): $($_.Exception.Message)"
                    Write-Log $errorMsg -Level Error
                    [System.IO.File]::WriteAllText($statusFile, "Error: $errorMsg")
                    
                    $iterationCount++
                    if ($iterationCount -lt $maxRetries) {
                        Write-Log "Retrying cleanup..." -Level Warning
                        Start-Sleep -Seconds $retryDelay
                    }
                }
            }
        }
        catch {
            Write-Log "Critical error in cleanup process: $($_.Exception.Message)" -Level Error
            [System.IO.File]::WriteAllText($statusFile, "Error: Critical failure in cleanup process")
            throw
        }
        finally {
            # Ensure status file is properly closed and removed
            if ([System.IO.File]::Exists($statusFile)) {
                try {
                    [System.IO.File]::Delete($statusFile)
                }
                catch {
                    Write-Log "Warning: Could not remove status file: $($_.Exception.Message)" -Level Warning
                }
            }
        }

        # Enhanced final cleanup
        Write-Log "Performing final cleanup..." -Level Info
        
        $filesToClean = @($executorScript, $systemAccessibleScriptPath, $markerFile, "$markerFile.status", $logFile)
        foreach ($file in $filesToClean) {
            $retryCount = 0
            while ($retryCount -lt $maxRetries -and [System.IO.File]::Exists($file)) {
                try {
                    # Force garbage collection before attempting deletion
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                    
                    # Try to acquire exclusive access
                    $fs = [System.IO.FileStream]::new(
                        $file,
                        [System.IO.FileMode]::Open,
                        [System.IO.FileAccess]::ReadWrite,
                        [System.IO.FileShare]::None
                    )
                    $fs.Close()
                    $fs.Dispose()
                    
                    [System.IO.File]::Delete($file)
                    Write-Log "Successfully removed: $file" -Level Debug
                    break
                }
                catch {
                    $retryCount++
                    Write-Log "Cleanup retry $retryCount for ${file}: $($_.Exception.Message)" -Level Warning
                    Start-Sleep -Seconds 2
                }
                finally {
                    if ($fs) { $fs.Dispose() }
                }
            }
        }

        return $completed
    }
    catch {
        Write-Log "Critical error during system context operation: $($_.Exception.Message)" -Level Error
        return $false
    }
}

# ----- Function: Get-SystemDiskSpace -----
function Get-SystemDiskSpace {
    Write-Log "Retrieving drive space information..." -Level Debug
    try {
        $systemDrive = Get-Volume -DriveLetter $env:SystemDrive[0]
        Write-Log "Successfully retrieved drive information for $($env:SystemDrive)" -Level Debug
        return $systemDrive
    } catch {
        Write-Log "Error retrieving drive information: $($_.Exception.Message)" -Level Error
        return $null
    }
}

# ----- Function: Show-DriveInfo -----
function Show-DriveInfo {
    param (
        [Parameter(Mandatory=$true)]
        [object]$Volume,
        
        [Parameter()]
        [string]$Label = "Drive Volume Details"
    )
    
    Write-Log "Displaying drive information for $($Volume.DriveLetter): label=$Label" -Level Debug
    
    # Calculate used space and percentage
    $usedSpace = $Volume.Size - $Volume.SizeRemaining
    $usedPercentage = [math]::Round(($usedSpace / $Volume.Size) * 100, 1)
    $freePercentage = [math]::Round(($Volume.SizeRemaining / $Volume.Size) * 100, 1)
    
    Write-Host "`n${Label}:" -ForegroundColor Green
    Write-Host "------------------------" -ForegroundColor Green
    Write-Host "Drive Letter: $($Volume.DriveLetter)" -ForegroundColor Cyan
    Write-Host "Drive Label: $($Volume.FileSystemLabel)" -ForegroundColor Cyan
    Write-Host "File System: $($Volume.FileSystem)" -ForegroundColor Cyan
    Write-Host "Drive Type: $($Volume.DriveType)" -ForegroundColor Cyan
    Write-Host "Total Size: $([math]::Round($Volume.Size/1GB, 2)) GB" -ForegroundColor Cyan
    Write-Host "Used Space: $([math]::Round($usedSpace/1GB, 2)) GB ($usedPercentage%)" -ForegroundColor $(if ($usedPercentage -gt 90) { "Yellow" } else { "Cyan" })
    Write-Host "Free Space: $([math]::Round($Volume.SizeRemaining/1GB, 2)) GB ($freePercentage%)" -ForegroundColor $(if ($freePercentage -lt 10) { "Red" } elseif ($freePercentage -lt 20) { "Yellow" } else { "Green" })
    Write-Host "Health Status: $($Volume.HealthStatus)" -ForegroundColor Cyan
    
    Write-Log "Drive $($Volume.DriveLetter) has $([math]::Round($Volume.SizeRemaining/1GB, 2)) GB free of $([math]::Round($Volume.Size/1GB, 2)) GB total" -Level Info
}

# ----- Function: Clear-SystemStorageOptimized -----
function Clear-SystemStorageOptimized {
    [CmdletBinding()]
    param()
    
    Write-Log "Initializing optimized system storage cleanup..." -Level Info
    
    try {
        # Initialize only required VSS object
        $vssService = Get-WmiObject -Class Win32_ShadowCopy

        # Enhanced cleanup paths using .NET methods
        $cleanupPaths = @(
            [System.Environment]::GetFolderPath('Windows') + '\Temp',
            [System.Environment]::GetFolderPath('LocalApplicationData') + '\Temp',
            [System.Environment]::GetFolderPath('InternetCache')
        )

        $totalBytesFreed = 0
        $errors = @()

        # Process VSS copies first for maximum space recovery
        Write-Log "Managing Volume Shadow Copies..." -Level Info
        try {
            # Get all shadow copies
            $shadowCopies = $vssService | Where-Object { $null -ne $_.ID }
            
            if ($null -ne ($shadowCopies | Where-Object { $null -ne $_.ID })) {
                $shadowCopies | ForEach-Object {
                    try {
                        $_.Delete()
                        Write-Log "Removed shadow copy: $($_.ID)" -Level Debug
                    }
                    catch {
                        Write-Log "Warning: Could not remove shadow copy $($_.ID): $($_.Exception.Message)" -Level Warning
                        $errors += "VSS:$($_.ID)"
                    }
                }
            }
            else {
                Write-Log "No shadow copies found to clean" -Level Debug
            }
        }
        catch {
            Write-Log "Warning: VSS management error: $($_.Exception.Message)" -Level Warning
            $errors += "VSS:General"
        }

        # Enhanced temp file cleanup using .NET methods
        foreach ($path in $cleanupPaths) {
            if ([System.IO.Directory]::Exists($path)) {
                Write-Log "Processing directory: $path" -Level Info
                
                try {
                    $files = [System.IO.Directory]::GetFiles($path, "*", [System.IO.SearchOption]::AllDirectories)
                    $directories = [System.IO.Directory]::GetDirectories($path, "*", [System.IO.SearchOption]::AllDirectories)
                    
                    # Process files first
                    foreach ($file in $files) {
                        try {
                            $fileInfo = [System.IO.FileInfo]::new($file)
                            $size = $fileInfo.Length
                            
                            # Use robust file deletion with retries
                            $retryCount = 0
                            $maxRetries = 3
                            $deleted = $false
                            
                            while (-not $deleted -and $retryCount -lt $maxRetries) {
                                try {
                                    [System.IO.File]::Delete($file)
                                    $totalBytesFreed += $size
                                    $deleted = $true
                                }
                                catch [System.IO.IOException] {
                                    $retryCount++
                                    if ($retryCount -lt $maxRetries) {
                                        Start-Sleep -Milliseconds 500
                                        [System.GC]::Collect()
                                        [System.GC]::WaitForPendingFinalizers()
                                    }
                                }
                            }
                            
                            if (-not $deleted) {
                                $errors += "File:$file"
                            }
                        }
                        catch {
                            Write-Log "Warning: Could not process file $file : $($_.Exception.Message)" -Level Warning
                            $errors += "File:$file"
                            continue
                        }
                    }
                    
                    # Process directories in reverse order (deepest first)
                    [array]::Reverse($directories)
                    foreach ($dir in $directories) {
                        try {
                            if ([System.IO.Directory]::GetFileSystemEntries($dir).Count -eq 0) {
                                [System.IO.Directory]::Delete($dir, $false)
                            }
                        }
                        catch {
                            Write-Log "Warning: Could not remove empty directory $dir : $($_.Exception.Message)" -Level Warning
                            $errors += "Dir:$dir"
                            continue
                        }
                    }
                }
                catch {
                    Write-Log "Error processing path $path : $($_.Exception.Message)" -Level Error
                    $errors += "Path:$path"
                    continue
                }
            }
        }

        # Report results
        $freedSpaceGB = [math]::Round($totalBytesFreed / 1GB, 2)
        Write-Log "Cleanup completed. Freed $freedSpaceGB GB of space." -Level Info
        
        if ($errors.Count -gt 0) {
            Write-Log "Cleanup completed with $($errors.Count) non-critical errors" -Level Warning
        }
        else {
            Write-Log "Cleanup completed successfully with no errors" -Level Info
        }

        return @{
            BytesFreed = $totalBytesFreed
            Errors = $errors
        }
    }
    catch {
        Write-Log "Critical error during cleanup: $($_.Exception.Message)" -Level Error
        throw
    }
}

# --------------------------------------------------
# Main Script Execution (after elevation if needed)
# --------------------------------------------------

# If not running as SYSTEM, initiate SYSTEM context.
if (-not (Test-RunningAsSystem)) {
    Start-SystemContext -ScriptPath $scriptPath
    # Suppress boolean output
    exit
}

Write-Log "Starting main cleanup process as SYSTEM account" -Level Info

# ----- Display Starting Drive Information -----
Write-Log "Collecting initial system drive information" -Level Info
$initialDrive = Get-SystemDiskSpace
if ($initialDrive) {
    Show-DriveInfo -Volume $initialDrive -Label "INITIAL Drive Volume Details"
} else {
    Write-Log "Failed to get initial drive information" -Level Warning
}

# ----- Display Shadow Copy Count -----
Write-Log "Checking initial shadow copy status..." -Level Info
$shadowOutputInitial = vssadmin list shadows 2>&1
Write-Log "Shadow copy command executed. Processing output..." -Level Debug

$totalShadowCopiesInitial = ($shadowOutputInitial | Select-String "Shadow Copy ID").Count
Write-Log "Initial Shadow Copy Count: $totalShadowCopiesInitial" -Level Info
Write-Host "Shadow Copy Count: ${totalShadowCopiesInitial}" -ForegroundColor Yellow

# Extract and display more details about shadow copies
if ($totalShadowCopiesInitial -gt 0) {
    Write-Host "`n===== Initial Shadow Copy Details =====" -ForegroundColor Yellow
    $shadowDetails = $shadowOutputInitial | Select-String -Pattern "Shadow Copy Volume:" -Context 0,0
    foreach ($detail in $shadowDetails) {
        Write-Host $detail -ForegroundColor DarkYellow
    }
}

# ----- Run Disk Cleanup Silently -----
Write-Log "Starting Disk Cleanup..." -Level Info
Write-Host "`nExecuting Windows Disk Cleanup utility..." -ForegroundColor Cyan

# Define StateFlags for silent cleanup (all standard cleanup options)
$cleanupFlags = @(
    "Active Setup Temp Folders",
    "BranchCache",
    "Downloaded Program Files",
    "Internet Cache Files",
    "Memory Dump Files",
    "Offline Pages Files",
    "Old ChkDsk Files",
    "Previous Installations",
    "Recycle Bin",
    "Service Pack Cleanup",
    "Setup Log Files",
    "System error memory dump files",
    "System error minidump files",
    "Temporary Files",
    "Temporary Setup Files",
    "Temporary Sync Files",
    "Thumbnail Cache",
    "Update Cleanup",
    "Upgrade Discarded Files",
    "Windows Defender",
    "Windows Error Reporting Files",
    "Windows ESD installation files",
    "Windows Upgrade Log Files"
)

Write-Log "Configuring cleanup settings..." -Level Info
Write-Host "Configuring cleanup settings..." -ForegroundColor Cyan

# Set registry values for silent cleanup
$regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches"
foreach ($flag in $cleanupFlags) {
    $flagPath = Join-Path $regPath $flag
    if (Test-Path $flagPath) {
        Set-ItemProperty -Path $flagPath -Name "StateFlags0001" -Value 2 -Type DWord
        Write-Log "Enabled cleanup for: $flag" -Level Debug
    }
}

# Start the cleanup process
$cleanmgrProcess = Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sagerun:1" -PassThru -NoNewWindow
$processId = $cleanmgrProcess.Id
Write-Log "Disk Cleanup process started with PID: $processId" -Level Debug

$startTime = Get-Date
$cleanupTimedOut = $false

while (!$cleanmgrProcess.HasExited) {
    try {
        # Get current process information
        $processInfo = Get-Process -Id $processId -ErrorAction Stop
        $currentMemory = [math]::Round($processInfo.WorkingSet64 / 1MB, 2)
        
        # Display process status
        $status = "Runtime: $((Get-Date - $startTime).ToString('mm\:ss')) | Memory: ${currentMemory}MB"
        Write-Host "`r$(' ' * 80)" -NoNewline
        Write-Host "`r$status" -NoNewline -ForegroundColor Cyan
        
        if ((Get-Date) - $startTime -gt $cleanupTimeout) {
            Write-Log "Disk Cleanup timeout reached after 20 minutes" -Level Warning
            Write-Host "`rDisk Cleanup timed out after 20 minutes." -ForegroundColor Yellow
            $cleanupTimedOut = $true
            Stop-Process -Id $processId -Force -ErrorAction Stop
            break
        }
    }
    catch {
        Write-Log "Error monitoring cleanup process: $($_.Exception.Message)" -Level Warning
        break
    }
    
    Start-Sleep -Seconds 1
}

$duration = (Get-Date) - $startTime
if (!$cleanupTimedOut) {
    Write-Host "`rDisk Cleanup completed in $($duration.ToString('mm\:ss')) minutes:seconds" -ForegroundColor Green
    Write-Log "Disk Cleanup complete. Duration: $($duration.ToString('mm\:ss'))" -Level Info
}

# Always proceed with Shadow Copy Management regardless of cleanup status
Write-Log "===== Starting Shadow Copy Management =====" -Level Info
Write-Host "`n===== Shadow Copy Management =====" -ForegroundColor Cyan

$shadowOutput = vssadmin list shadows 2>&1
$totalShadowCopies = ($shadowOutput | Select-String "Shadow Copy ID").Count
Write-Log ("Shadow Copies Found: $totalShadowCopies") -Level Info
Write-Host "Shadow Copies Found: ${totalShadowCopies}" -ForegroundColor Yellow

if ($totalShadowCopies -gt 1) {
    $shadowIDs = $shadowOutput | Select-String "Shadow Copy ID:" | ForEach-Object {
        if ($_ -match "Shadow Copy ID:\s+({[^}]+})") { $matches[1] }
    }
    
    # Display all shadow copies before removal
    Write-Host "`nShadow Copies to be managed:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $shadowIDs.Count; $i++) {
        $id = $shadowIDs[$i]
        $keepStatus = if ($i -eq $shadowIDs.Count - 1) { "KEEP" } else { "REMOVE" }
        Write-Host "  [$keepStatus] Shadow Copy ID: $id" -ForegroundColor $(if ($keepStatus -eq "KEEP") { "Green" } else { "Yellow" })
    }
    
    $removedCount = $shadowIDs.Count - 1
    Write-Host "`nRemoving $removedCount shadow copies..." -ForegroundColor Yellow
    
    # Remove all but the most recent shadow copy
    for ($i = 0; $i -lt $shadowIDs.Count - 1; $i++) {
        $id = $shadowIDs[$i]
        Write-Log ("Removing shadow copy: $id") -Level Info
        Write-Host "  Removing shadow copy: $id" -ForegroundColor Yellow
        $result = vssadmin delete shadows /shadow=$id /quiet 2>&1
        Write-Log "Result: $result" -Level Debug
    }
    $remainingCount = 1
    Write-Host "Keeping 1 most recent shadow copy" -ForegroundColor Green
} else {
    $removedCount = 0
    $remainingCount = $totalShadowCopies
    if ($totalShadowCopies -eq 1) {
        Write-Host "Only one shadow copy exists. Keeping it." -ForegroundColor Green
    } else {
        Write-Host "No shadow copies found." -ForegroundColor Cyan
    }
}

Write-Log ("Shadow Copy cleanup summary: Found=$totalShadowCopies, Removed=$removedCount, Remaining=$remainingCount") -Level Info
Write-Host "`nShadow Copy cleanup summary:" -ForegroundColor Cyan
Write-Host "  - Initial count: $totalShadowCopiesInitial" -ForegroundColor Cyan
Write-Host "  - Removed count: $removedCount" -ForegroundColor Yellow
Write-Host "  - Remaining count: $remainingCount" -ForegroundColor Green

# ----- Display Ending Drive Information -----
Write-Log "Collecting final system drive information" -Level Info
$finalDrive = Get-SystemDiskSpace
if ($finalDrive -and $initialDrive) {
    # Calculate space reclaimed
    $spaceReclaimed = $finalDrive.SizeRemaining - $initialDrive.SizeRemaining
    $freeSpaceBefore = [math]::Round($initialDrive.SizeRemaining/1GB, 2)
    $freeSpaceAfter = [math]::Round($finalDrive.SizeRemaining/1GB, 2)
    $spaceReclaimedGB = [math]::Round($spaceReclaimed/1GB, 2)
    
    Show-DriveInfo -Volume $finalDrive -Label "FINAL Drive Volume Details"
    
    Write-Host "`nCleanup Results:" -ForegroundColor Green
    Write-Host "------------------------" -ForegroundColor Green
    Write-Host "Free Space Before: $freeSpaceBefore GB" -ForegroundColor Yellow
    Write-Host "Free Space After:  $freeSpaceAfter GB" -ForegroundColor Green
    Write-Host "Space Reclaimed:   $spaceReclaimedGB GB" -ForegroundColor $(if ($spaceReclaimedGB -gt 0) {"Green"} else {"Yellow"})
    
    Write-Log "Cleanup complete. Space reclaimed: $spaceReclaimedGB GB" -Level Info
} else {
    Write-Log "Failed to get final drive information" -Level Warning
    Write-Host "Failed to get final drive information" -ForegroundColor Red
}

Write-Log "===== System Storage Cleanup Completed =====" -Level Info
Write-Host "`n===== System Storage Cleanup Completed =====" -ForegroundColor Green
Write-Host "Log file: $script:LogFile" -ForegroundColor Cyan

# Update the main script block to use the optimized cleanup
try {
    $result = Clear-SystemStorageOptimized
    
    if ($result.BytesFreed -gt 0) {
        $freedGB = [math]::Round($result.BytesFreed / 1GB, 2)
        Write-Host "Successfully freed $freedGB GB of system storage" -ForegroundColor Green
        
        if ($result.Errors.Count -gt 0) {
            Write-Host "Completed with $($result.Errors.Count) non-critical errors. See log for details." -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "No significant space was freed. System may already be clean." -ForegroundColor Cyan
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
