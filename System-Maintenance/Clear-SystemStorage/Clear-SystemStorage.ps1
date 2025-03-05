# =============================================================================
# Script: Clear-SystemStorage.ps1
# Created: 2025-02-27 18:55:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-05 10:15:00 UTC
# Updated By: jdyer-nuvodia
# Version: 4.0
# Additional Info: Enhanced logging and detailed progress output during script execution.
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
$script:LogFile = Join-Path -Path $script:OriginalScriptDirectory -ChildPath "ClearSystemStorage_$timestamp.log"

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
Write-Log "Script version: 4.0" -Level Info
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
        [int]$TimeoutSeconds = 300
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
        
        $systemAccessibleTemp = "$env:SystemRoot\Temp"
        $scriptFileName = Split-Path -Leaf $ScriptPath
        $systemAccessibleScriptPath = Join-Path -Path $systemAccessibleTemp -ChildPath $scriptFileName
        
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
try {
    Write-Host '[SYSTEM EXECUTOR] Starting execution of main script as SYSTEM'
    # Run with debug and verbose output enabled
    & '$systemAccessibleScriptPath' -Debug -Verbose
    if (`$?) {
        Write-Host '[SYSTEM EXECUTOR] Script executed successfully.'
        Set-Content -Path '$markerFile' -Value 'Complete' -Force
    } else {
        Write-Error '[SYSTEM EXECUTOR] Script failed with non-zero exit code.'
    }
}
catch {
    Write-Host '[SYSTEM EXECUTOR] Error occurred: ' + `$_.Exception.Message -ForegroundColor Red
    Write-Error '[SYSTEM EXECUTOR] Error occurred: ' + `$_.Exception.Message
    exit 1
}
finally {
    Write-Host '[SYSTEM EXECUTOR] Execution complete.'
    Stop-Transcript
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
        Write-Host "Task running as SYSTEM. This may take a few minutes..." -ForegroundColor Cyan
        
        $timeout = (Get-Date).AddSeconds($TimeoutSeconds)
        $completed = $false
        $counter = 0
        $progressChars = @('|', '/', '-', '\')
        
        while ((Get-Date) -lt $timeout -and -not $completed) {
            if (Test-Path $markerFile) {
                $completed = $true
                Write-Log "Completion marker found." -Level Verbose
                break
            }
            
            # Show a simple spinner to indicate progress
            $counter++
            $progressChar = $progressChars[$counter % $progressChars.Length]
            $timeNow = Get-Date -Format "HH:mm:ss"
            Write-Host "`r[$timeNow] Processing $progressChar" -NoNewline -ForegroundColor Cyan
            
            Start-Sleep -Seconds 1
            
            # Check if the log file exists and show latest entries every second
            if (Test-Path $logFile) {
                try {
                    $latestLogs = Get-Content -Path $logFile -Tail 2 -ErrorAction SilentlyContinue
                    if ($latestLogs) {
                        Write-Host "`r                                                                     " -NoNewline
                        $latestLog = $latestLogs | Where-Object { $_ -match '\S' } | Select-Object -Last 1
                        if ($latestLog) {
                            Write-Host "`r$timeNow - $latestLog" -ForegroundColor Cyan
                        }
                    }
                } catch {
                    # Silently continue if we can't read the log
                }
            }
        }
        
        # Clear the progress line
        Write-Host "`r                                                   " -NoNewline
        
        if ($completed) {
            Write-Host "`rSystem cleanup completed!" -ForegroundColor Green
            Write-Log "System cleanup completed via SYSTEM context" -Level Info
            
            # Display the SYSTEM log to the console for more visibility
            if (Test-Path $logFile) {
                Write-Host "`n===== SYSTEM Context Execution Log =====" -ForegroundColor Yellow
                Get-Content -Path $logFile | ForEach-Object {
                    Write-Host $_ -ForegroundColor DarkGray
                }
                Write-Host "===== End of SYSTEM Context Log =====" -ForegroundColor Yellow
            }
            
            # Clean up files after successful completion
            try {
                Write-Log "Cleaning up temporary task and files" -Level Debug
                if (Get-ScheduledTask -TaskName $jobName -ErrorAction SilentlyContinue) {
                    Unregister-ScheduledTask -TaskName $jobName -Confirm:$false -ErrorAction SilentlyContinue
                    Write-Log "Unregistered scheduled task: $jobName" -Level Debug
                }
                Remove-Item -Path $executorScript -Force -ErrorAction SilentlyContinue
                Remove-Item -Path $systemAccessibleScriptPath -Force -ErrorAction SilentlyContinue
                Remove-Item -Path $markerFile -Force -ErrorAction SilentlyContinue
                Write-Log "Temporary files removed" -Level Debug
            } catch {
                Write-Log "Cleanup warning: $($_.Exception.Message)" -Level Warning
            }
        } else {
            Write-Host "`rTimeout reached waiting for system cleanup to complete." -ForegroundColor Yellow
            Write-Log "Timeout reached before completion." -Level Warning
        }
        
        return $completed
    }
    catch {
        Write-Log "Error during SYSTEM context setup: $($_.Exception.Message)" -Level Error
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
    
    Write-Host "`n$Label:" -ForegroundColor Green
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
Write-Host "Shadow Copies Found: $totalShadowCopiesInitial" -ForegroundColor Yellow

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

# Create a temporary file to monitor disk cleanup progress
$cleanmgrLogPath = "$env:TEMP\cleanmgr_progress.txt"
$cleanmgrProcess = Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sagerun:1" -PassThru -NoNewWindow
$processId = $cleanmgrProcess.Id
Write-Log "Disk Cleanup process started with PID: $processId" -Level Debug

$startTime = Get-Date
Write-Host "Disk Cleanup started at: $(Get-Date -Format 'HH:mm:ss')" -ForegroundColor Cyan

# Monitor the cleanmgr process
$i = 0
$spinChars = '|','/','-','\'
Write-Host "Progress: " -NoNewline -ForegroundColor Cyan

while (!$cleanmgrProcess.HasExited) {
    $char = $spinChars[$i % $spinChars.Length]
    Write-Host "`r$(Get-Date -Format 'HH:mm:ss') - Cleaning in progress $char" -NoNewline -ForegroundColor Cyan
    
    # Check if we can get the window title of cleanmgr for more info
    try {
        $title = (Get-Process -Id $processId -ErrorAction SilentlyContinue).MainWindowTitle
        if ($title -and $title -ne "") {
            Write-Host "`r$(Get-Date -Format 'HH:mm:ss') - $title" -ForegroundColor Cyan
        }
    } catch {
        # Just continue
    }
    
    $i++
    Start-Sleep -Milliseconds 500
}

$duration = (Get-Date) - $startTime
Write-Host "`rDisk Cleanup completed in $($duration.ToString('mm\:ss')) minutes:seconds" -ForegroundColor Green
Write-Log "Disk Cleanup complete. Duration: $($duration.ToString('mm\:ss'))" -Level Info

# ----- Shadow Copy Management -----
Write-Log "===== Starting Shadow Copy Management =====" -Level Info
Write-Host "`n===== Shadow Copy Management =====" -ForegroundColor Cyan

$shadowOutput = vssadmin list shadows 2>&1
$totalShadowCopies = ($shadowOutput | Select-String "Shadow Copy ID").Count
Write-Log ("Shadow Copies Found: $totalShadowCopies") -Level Info
Write-Host "Shadow Copies Found After Cleanup: $totalShadowCopies" -ForegroundColor Yellow

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
