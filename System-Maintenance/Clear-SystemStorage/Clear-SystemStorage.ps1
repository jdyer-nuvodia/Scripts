# =============================================================================
# Script: Clear-SystemStorage.ps1
# Created: 2025-02-27 18:55:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-04 00:12:00 UTC
# Updated By: jdyer-nuvodia
# Version: 3.9
# Additional Info: Simplified elevation logic to always run as SYSTEM and
#                streamlined log file handling to the script directory.
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
        'Info' { Write-Host $Message -ForegroundColor White }
        'Warning' { Write-Host $Message -ForegroundColor Yellow }
        'Error' { Write-Host $Message -ForegroundColor Red }
        'Debug' { if ($DebugPreference -eq 'Continue') { Write-Host "DEBUG: $Message" -ForegroundColor Magenta } }
        'Verbose' { if ($VerbosePreference -eq 'Continue') { Write-Host "VERBOSE: $Message" -ForegroundColor Cyan } }
    }
    
    # Always append the log entry to the log file.
    Add-Content -Path $script:LogFile -Value $logEntry
}

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
        
        $executorContent = @"
# Executor script for Clear-SystemStorage running as SYSTEM
`$ErrorActionPreference = 'Stop'
Start-Transcript -Path '$logFile' -Force
try {
    Write-Host 'Starting execution of main script as SYSTEM'
    & '$systemAccessibleScriptPath'
    if (`$?) {
        Write-Host 'Script executed successfully.'
        Set-Content -Path '$markerFile' -Value 'Complete' -Force
    } else {
        Write-Error 'Script failed with non-zero exit code.'
    }
}
catch {
    Write-Host 'Error occurred: ' + `$_.Exception.Message -ForegroundColor Red
    Write-Error 'Error occurred: ' + `$_.Exception.Message
    exit 1
}
finally {
    Write-Host 'Execution complete.'
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
            Write-Host "`rProcessing $progressChar" -NoNewline -ForegroundColor Cyan
            
            Start-Sleep -Seconds 2
            
            # Check if the log file exists and show latest entries
            if ($counter % 5 -eq 0 -and (Test-Path $logFile)) {
                try {
                    $latestLog = Get-Content -Path $logFile -Tail 1
                    Write-Host "`rLatest status: $latestLog" -ForegroundColor Cyan
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
            # Clean up files after successful completion
            try {
                if (Get-ScheduledTask -TaskName $jobName -ErrorAction SilentlyContinue) {
                    Unregister-ScheduledTask -TaskName $jobName -Confirm:$false -ErrorAction SilentlyContinue
                }
                Remove-Item -Path $executorScript -Force -ErrorAction SilentlyContinue
                Remove-Item -Path $systemAccessibleScriptPath -Force -ErrorAction SilentlyContinue
                Remove-Item -Path $markerFile -Force -ErrorAction SilentlyContinue
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

# --------------------------------------------------
# Main Script Execution (after elevation if needed)
# --------------------------------------------------

# If not running as SYSTEM, initiate SYSTEM context.
if (-not (Test-RunningAsSystem)) {
    $result = Start-SystemContext -ScriptPath $scriptPath
    # Suppress boolean output
    exit
}

# ----- Display Starting Drive Information -----
Write-Log "===== Starting Drive Information =====" -Level Info
Get-PSDrive -PSProvider FileSystem | ForEach-Object {
    $driveInfo = "Drive $($_.Name): Used = $($_.Used) Free = $($_.Free)"
    Write-Log -Message $driveInfo -Level Info
    Write-Host $driveInfo -ForegroundColor Cyan
}

# ----- Display Shadow Copy Count -----
$shadowOutputInitial = vssadmin list shadows 2>&1
$totalShadowCopiesInitial = ($shadowOutputInitial | Select-String "Shadow Copy ID").Count
Write-Log "Initial Shadow Copy Count: $totalShadowCopiesInitial" -Level Info
Write-Host "Shadow Copies Found: $totalShadowCopiesInitial" -ForegroundColor Yellow

# ----- Run Disk Cleanup Silently -----
Write-Log "Starting Disk Cleanup..." -Level Info
Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sagerun:1" -Wait -NoNewWindow
Write-Log "Disk Cleanup complete." -Level Info

# ----- Shadow Copy Management -----
Write-Log "===== Starting Shadow Copy Management =====" -Level Info
$shadowOutput = vssadmin list shadows 2>&1
$totalShadowCopies = ($shadowOutput | Select-String "Shadow Copy ID").Count
Write-Log ("Shadow Copies Found: $totalShadowCopies") -Level Info

if ($totalShadowCopies -gt 1) {
    $shadowIDs = $shadowOutput | Select-String "Shadow Copy ID:" | ForEach-Object {
        if ($_ -match "Shadow Copy ID:\s+({[^}]+})") { $matches[1] }
    }
    $removedCount = $shadowIDs.Count - 1
    for ($i = 0; $i -lt $shadowIDs.Count - 1; $i++) {
        $id = $shadowIDs[$i]
        Write-Log ("Removing shadow copy: $id") -Level Info
        vssadmin delete shadows /shadow=$id /quiet | Out-Null
    }
    $remainingCount = 1
} else {
    $removedCount = 0
    $remainingCount = $totalShadowCopies
}

Write-Log ("Shadow Copy cleanup summary: Found=$totalShadowCopies, Removed=$removedCount, Remaining=$remainingCount") -Level Info

# ----- Display Ending Drive Information -----
Write-Log "===== Ending Drive Information =====" -Level Info
Get-PSDrive -PSProvider FileSystem | ForEach-Object {
    $driveInfo = "Drive $($_.Name): Used = $($_.Used) Free = $($_.Free)"
    Write-Log -Message $driveInfo -Level Info
    Write-Host $driveInfo -ForegroundColor Green
}

Write-Log "System Storage Cleanup Completed." -Level Info
