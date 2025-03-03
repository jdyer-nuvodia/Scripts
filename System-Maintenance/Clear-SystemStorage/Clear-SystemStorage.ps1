# =============================================================================
# Script: Clear-SystemStorage.ps1
# Created: 2025-02-27 18:55:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-06 19:15:00 UTC
# Updated By: jdyer-nuvodia
# Version: 2.9
# Additional Info: Enhanced progress feedback during cleanup operations
# =============================================================================

<#
.SYNOPSIS
    Performs comprehensive system storage cleanup by managing disk space and shadow copies.
.DESCRIPTION
    This script performs two main cleanup operations:
    1. Runs Windows Disk Cleanup utility silently
    2. Manages Volume Shadow Copies by keeping only the newest restore point
    
    Key actions:
    - Executes disk cleanup with predefined settings
    - Lists and manages shadow copies
    - Maintains the most recent shadow copy while removing others
    - Displays drive information before and after cleanup for comparison
    - Logs detailed operations to a file for debugging
    
    Dependencies:
    - Windows Volume Shadow Copy Service
    - Administrative privileges
.PARAMETER NoElevate
    Prevents the script from attempting to elevate to SYSTEM context when already running as a scheduled task
.EXAMPLE
    .\Clear-SystemStorage.ps1
    Runs the script with default settings
.EXAMPLE
    .\Clear-SystemStorage.ps1 -Verbose
    Runs the script with verbose console output
.NOTES
    Security Level: High
    Required Permissions: Administrative privileges
    Validation Requirements: 
    - Verify shadow copy retention
    - Check disk space recovery
#>
param(
    [switch]$NoElevate
)

# Initialize global variables
$scriptPath = $MyInvocation.PSCommandPath
if (-not $scriptPath) {
    # Fallback for direct console execution
    $scriptPath = $MyInvocation.MyCommand.Definition
    if (-not $scriptPath) {
        # Ultimate fallback to current directory
        $scriptPath = Join-Path -Path (Get-Location) -ChildPath "Clear-SystemStorage.ps1"
    }
}
$scriptDirectory = Split-Path -Path $scriptPath -Parent
$script:LogFile = Join-Path -Path $scriptDirectory -ChildPath "ClearSystemStorage_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter()]
        [ValidateSet('Info', 'Warning', 'Error', 'Debug', 'Verbose')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Special handling for shadow copy information with enhanced visibility
    $isShadowCopyInfo = $Message -match "Shadow Copy|shadow cop"
    
    # Write to console with appropriate color
    switch ($Level) {
        'Info' { 
            # Use cyan for shadow copy information to make it stand out
            if ($isShadowCopyInfo -and $Message -match "^(🔍|✅|🗑️|📊)") {
                if ($Message -match "^✅|^📊") {
                    Write-Host $Message -ForegroundColor Green
                } 
                elseif ($Message -match "^🗑️") {
                    Write-Host $Message -ForegroundColor Yellow
                }
                else {
                    Write-Host $Message -ForegroundColor Cyan 
                }
            }
            else {
                Write-Host $Message
            }
            
            if ($VerbosePreference -eq 'Continue') { Write-Host $logEntry }
        }
        'Warning' { Write-Host $Message -ForegroundColor Yellow }
        'Error' { Write-Host $Message -ForegroundColor Red }
        'Debug' { 
            if ($DebugPreference -eq 'Continue') { 
                Write-Host "DEBUG: $Message" -ForegroundColor Magenta 
            }
        }
        'Verbose' { 
            if ($VerbosePreference -eq 'Continue') { 
                Write-Host "VERBOSE: $Message" -ForegroundColor Cyan 
            }
        }
    }
    
    # Always write to log file - logging is enabled by default now
    Add-Content -Path $script:LogFile -Value $logEntry
}

function Test-RunningAsSystem {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $isSystem = $currentUser.User.Value -eq "S-1-5-18"
    Write-Log -Message "Checking if running as SYSTEM: $isSystem" -Level Verbose
    return $isSystem
}

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
    
    try {
        $jobName = "SystemContextJob_$([Guid]::NewGuid())"
        $systemAccessibleTemp = "$env:SystemRoot\Temp"
        $scriptFileName = Split-Path -Leaf $ScriptPath
        $systemAccessibleScriptPath = Join-Path -Path $systemAccessibleTemp -ChildPath $scriptFileName
        
        # Copy the script to system-accessible location
        Write-Log "Copying script to system-accessible location: $systemAccessibleScriptPath" -Level Verbose
        Copy-Item -Path $ScriptPath -Destination $systemAccessibleScriptPath -Force
        
        # Create executor script that will run our main script
        $executorScript = Join-Path -Path $systemAccessibleTemp -ChildPath "$jobName-executor.ps1"
        $logFile = Join-Path $systemAccessibleTemp "$jobName.log"
        $markerFile = Join-Path $systemAccessibleTemp "$jobName.marker"
        
        $executorContent = @"
# Executor script for Clear-SystemStorage
`$ErrorActionPreference = 'Stop'
`$VerbosePreference = '$(if ($VerbosePreference -eq 'Continue') { 'Continue' } else { 'SilentlyContinue' })'

Start-Transcript -Path "$logFile" -Force

try {
    Write-Host "Starting execution of main script as SYSTEM"
    & "$systemAccessibleScriptPath" -NoElevate $(if ($VerbosePreference -eq 'Continue') { '-Verbose' })
    
    if (`$?) {
        Write-Host "Script executed successfully"
        Set-Content -Path "$markerFile" -Value "Complete" -Force
    } else {
        Write-Error "Script failed with non-zero exit code"
    }
}
catch {
    Write-Host "Error occurred: `$(`$_.Exception.Message)" -ForegroundColor Red
    Write-Error "Error occurred: `$(`$_.Exception.Message)"
    exit 1
}
finally {
    Write-Host "Execution complete"
    Stop-Transcript
}
"@
        
        Write-Log "Creating executor script at $executorScript" -Level Verbose
        Set-Content -Path $executorScript -Value $executorContent -Force
        
        # Create action for scheduled task - use the executor script
        $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$executorScript`""
        $principal = New-ScheduledTaskPrincipal -UserId "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
        $task = New-ScheduledTask -Action $action -Principal $principal -Settings $settings
        
        Write-Log "Registering scheduled task $jobName" -Level Verbose
        Register-ScheduledTask -TaskName $jobName -InputObject $task | Out-Null
        Start-ScheduledTask -TaskName $jobName

        # Wait for completion with progress indicator
        $timeout = (Get-Date).AddSeconds($TimeoutSeconds)
        $completed = $false
        $seenLogLines = @{}
        $dotCount = 0
        $progressDisplayTime = [DateTime]::MinValue
        $lastStatusUpdate = [DateTime]::Now
        
        Write-Log "Waiting for system cleanup to complete..." -Level Info
        Write-Host "System cleanup progress:" -ForegroundColor Cyan
        cus on shadow copy information
        while ((Get-Date) -lt $timeout -and -not $completed) {
            # Check for completion marker file firstow Copy Details:"; Color = "Cyan" },
            if (Test-Path $markerFile) {rn = "Found \d+ shadow cop"; Color = "Cyan" },
                $completed = $true   @{ Pattern = "ID: [A-Za-z0-9-]+"; Color = "White" },
                Write-Log "Completion marker found" -Level Verbose       @{ Pattern = "Created: "; Color = "White" },
                continue        @{ Pattern = "Preserving most recent"; Color = "Green" },
            }y"; Color = "Yellow" },
            eted shadow copy"; Color = "Yellow" },
            # Then check task status= "Green" }
            try {
                $status = Get-ScheduledTask -TaskName $jobName -ErrorAction SilentlyContinue
                if ($status -and $status.State -eq "Ready") {dy removed
                    $completed = $true
                    Write-Log "Task completed according to scheduler" -Level Verbose
                    continue   continue
                }}
            } catch {
                # Task might be completed and already removed
                if (Test-Path $markerFile) {
                    $completed = $true
                    continuees
                }
            }
            dotCount = ($dotCount + 1) % 4
            # Display log updates without duplication   Write-Host "► Cleanup in progress$('.' * $dotCount)    " -ForegroundColor Cyan -NoNewline
            if (Test-Path $logFile) {    Write-Host "`r" -NoNewline
                # Show periodic status updates even if no new log lines
                if (([DateTime]::Now - $lastStatusUpdate).TotalSeconds -gt 15) {
                    $lastStatusUpdate = [DateTime]::Now
                    $dotCount = ($dotCount + 1) % 4ion|failed" -SimpleMatch
                    Write-Host "► Cleanup in progress$('.' * $dotCount)    " -ForegroundColor Cyan -NoNewline
                    Write-Host "`r" -NoNewline
                }
                or Red
                # Add logic to detect errors in the log and report them
                $errorLines = Select-String -Path $logFile -Pattern "Error|Exception|failed" -SimpleMatch
                foreach ($errorLine in $errorLines) {
                    $errorText = $errorLine.Line.Trim()
                    if (-not $seenLogLines.ContainsKey($errorText)) {# Check for important progress events
                        Write-Host "ERROR detected: $errorText" -ForegroundColor Red
                        $seenLogLines[$errorText] = $true
                    } [Cc]leanup"; Color = "Cyan" },
                }try keys"; Color = "Cyan" },
                r = "Cyan" },
                # Check for important progress events
                $progressKeywords = @(
                    @{ Pattern = "Disk [Cc]leanup.*complete"; Color = "Green" }, "Space freed"; Color = "Green" },
                    @{ Pattern = "Starting [Dd]isk [Cc]leanup"; Color = "Cyan" },{ Pattern = "Found.*drive"; Color = "Cyan" },
                    @{ Pattern = "Setting registry keys"; Color = "Cyan" },@{ Pattern = "minutes"; Color = "Magenta" }  # For long-running operations
                    @{ Pattern = "Starting Shadow Copy"; Color = "Cyan" },
                    @{ Pattern = "Shadow Copy.*complete"; Color = "Green" },
                    @{ Pattern = "[Dd]elete.*shadow copy"; Color = "Yellow" },t all new log lines with better filtering
                    @{ Pattern = "Space freed"; Color = "Green" }, 20 -ErrorAction SilentlyContinue
                    @{ Pattern = "Found.*drive"; Color = "Cyan" },ntLines) {
                    @{ Pattern = "minutes"; Color = "Magenta" }  # For long-running operations
                )
                e) -or
                # Get all new log lines with better filteringranscript started|^Transcript ended|^Windows PowerShell transcript") {
                $currentLines = Get-Content $logFile -Tail 20 -ErrorAction SilentlyContinue
                foreach ($line in $currentLines) {
                    $trimmedLine = $line.Trim()
                    # Skip empty lines and already seen lines Mark line as seen
                    if (-not $trimmedLine -or $seenLogLines.ContainsKey($trimmedLine) -or$seenLogLines[$trimmedLine] = $true
                        $trimmedLine -match "^Transcript started|^Transcript ended|^Windows PowerShell transcript") {
                        continueant information
                    }
                    oreach ($keyword in $progressKeywords) {
                    # Mark line as seen       if ($trimmedLine -match $keyword.Pattern) {
                    $seenLogLines[$trimmedLine] = $true        Write-Host $trimmedLine -ForegroundColor $keyword.Color
                    
                    # Use colors for important information= [DateTime]::Now
                    $colorMatch = $false
                    foreach ($keyword in $progressKeywords) {
                        if ($trimmedLine -match $keyword.Pattern) {
                            Write-Host $trimmedLine -ForegroundColor $keyword.Color
                            $colorMatch = $truer lines
                            $lastStatusUpdate = [DateTime]::Now   if (-not $colorMatch) {
                            break           Write-Host $trimmedLine
                        }        }
                    }
                       } else {
                    # Default display for other lines        # Show a simple activity indicator if no log file yet
                    if (-not $colorMatch) {
                        Write-Host $trimmedLine                if (($currentTime - $progressDisplayTime).TotalMilliseconds -gt 500) {
                    }tTime
                }= ($dotCount + 1) % 4
            } else { -ForegroundColor Cyan -NoNewline
                # Show a simple activity indicator if no log file yet        Write-Host "`r" -NoNewline
                $currentTime = [DateTime]::Now
                if (($currentTime - $progressDisplayTime).TotalMilliseconds -gt 500) {
                    $progressDisplayTime = $currentTime
                    $dotCount = ($dotCount + 1) % 4
                    Write-Host "Waiting for task to start$('.' * $dotCount)    " -ForegroundColor Cyan -NoNewline
                    Write-Host "`r" -NoNewline
                }
            }
            
            Start-Sleep -Milliseconds 250pleted) {
        }-Log "Task did not complete within timeout period" -Level Error
        
        Write-Host ""  # Clear the line after activity indicator
-Log "Checking for task status..." -Level Info
        # Check final status with more diagnosticsry {
        if (-not $completed) {skStatus = Get-ScheduledTask -TaskName $jobName -ErrorAction SilentlyContinue
            Write-Log "Task did not complete within timeout period" -Level Error
                   Write-Log "Task state: $($taskStatus.State)" -Level Info
            # More detailed diagnostics        $taskInfo = Get-ScheduledTaskInfo -TaskName $jobName -ErrorAction SilentlyContinue
            Write-Log "Checking for task status..." -Level Info
            try {sk info: Last run time: $($taskInfo.LastRunTime), Result: $($taskInfo.LastTaskResult)" -Level Info
                $taskStatus = Get-ScheduledTask -TaskName $jobName -ErrorAction SilentlyContinue
                if ($taskStatus) {
                    Write-Log "Task state: $($taskStatus.State)" -Level Info   else {
                    $taskInfo = Get-ScheduledTaskInfo -TaskName $jobName -ErrorAction SilentlyContinue  Write-Log "Task no longer exists - it may have completed but failed to create marker file" -Level Warning
                    if ($taskInfo) {
                        Write-Log "Task info: Last run time: $($taskInfo.LastRunTime), Result: $($taskInfo.LastTaskResult)" -Level Info
                    }   catch {
                }                Write-Log "Error getting task status: $_" -Level Error
                else {
                    Write-Log "Task no longer exists - it may have completed but failed to create marker file" -Level Warning
                }
            }
            catch {       Write-Log "Contents of log file:" -Level Info
                Write-Log "Error getting task status: $_" -Level Error        Get-Content $logFile | ForEach-Object { Write-Log $_ -Level Info }
            }
            
            # Look for log content for clues was created" -Level Warning
            if (Test-Path $logFile) {
                Write-Log "Contents of log file:" -Level Info
                Get-Content $logFile | ForEach-Object { Write-Log $_ -Level Info }
            }
            else {
                Write-Log "No log file was created" -Level WarninguledTask -TaskName $jobName -ErrorAction SilentlyContinue) {
            }Action SilentlyContinue
        }

        # Cleanup
        Write-Log "Cleaning up temporary files and tasks" -Level Verbose$filePath in @($logFile, $markerFile, $systemAccessibleScriptPath, $executorScript)) {
        if (Get-ScheduledTask -TaskName $jobName -ErrorAction SilentlyContinue) {f (Test-Path $filePath) {
            Unregister-ScheduledTask -TaskName $jobName -Confirm:$false -ErrorAction SilentlyContinue       try {
        }            if ($filePath -eq $logFile) {
         debugging
        # Cleanup files with better error handling                   $logContent = Get-Content -Path $logFile -Raw -ErrorAction SilentlyContinue
        foreach ($filePath in @($logFile, $markerFile, $systemAccessibleScriptPath, $executorScript)) {             Write-Log "SYSTEM execution log: $logContent" -Level Debug
            if (Test-Path $filePath) {
                try {emove-Item $filePath -Force -ErrorAction SilentlyContinue
                    if ($filePath -eq $logFile) {           }
                        # Save the log content for debugging               catch {
                        $logContent = Get-Content -Path $logFile -Raw -ErrorAction SilentlyContinue                    Write-Log "Could not remove temporary file $filePath`: $_" -Level Warning
                        Write-Log "SYSTEM execution log: $logContent" -Level Debug
                    }
                    Remove-Item $filePath -Force -ErrorAction SilentlyContinue    }
                }
                catch {it 1 }
                    Write-Log "Could not remove temporary file $filePath`: $_" -Level Warning
                }
            }
        }
        
        if ($completed) { exit 0 } else { exit 1 }
    }
    catch {
        Write-Log "Failed to elevate to SYSTEM context: $_" -Level Errorisk Cleanup process..." -Level Info
        return $false
    }    try {
}

function Start-DiskCleanup {try {
    Write-Log "Starting Disk Cleanup process..." -Level Infoew-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\*' `
    
    try {Write-Log "Registry keys set successfully" -Level Verbose
        # Create the StateFlags registry key
        Write-Log "Setting registry keys for automatic cleanup" -Level Verbose
        try {ing registry keys: $_" -Level Error
            New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\*' `
                -Name StateFlags0001 -Value 2 -PropertyType DWord -Force -ErrorAction Stop | Out-Null
            Write-Log "Registry keys set successfully" -Level Verbose
        }arameter to prevent GUI
        catch {parameters" -Level Verbose
            Write-Log "Error setting registry keys: $_" -Level Error
            return $false
        } -ArgumentList '/sagerun:1 /LOWDISK' -WindowStyle Hidden -PassThru

        # Run Disk Cleanup silently with LOWDISK parameter to prevent GUI
        Write-Log "Executing cleanmgr.exe with /sagerun:1 /LOWDISK parameters" -Level Verbose
        -Date
        try {
            $cleanmgrProcess = Start-Process -FilePath cleanmgr -ArgumentList '/sagerun:1 /LOWDISK' -WindowStyle Hidden -PassThru
            not $cleanmgrProcess.HasExited) {
            # Monitor the process with updates= (Get-Date) - $startTime
            Write-Log "Disk cleanup process started with ID: $($cleanmgrProcess.Id)" -Level Debug
            $startTime = Get-Date
            $lastUpdateTime = $startTime
            
            while (-not $cleanmgrProcess.HasExited) {UpdateTime = Get-Date
                $runtime = (Get-Date) - $startTimef ($runtime.TotalSeconds -gt 60) {
                $timeSinceLastUpdate = (Get-Date) - $lastUpdateTimete-Log "Disk cleanup running for $([int]$runtime.TotalMinutes) minutes and $([int]($runtime.TotalSeconds % 60)) seconds..." -Level Info
                
                # More frequent updates on cleanup progressalSeconds) seconds..." -Level Info
                if ($timeSinceLastUpdate.TotalSeconds -gt 30) {
                    $lastUpdateTime = Get-Date   
                    if ($runtime.TotalSeconds -gt 60) {    try {
                        Write-Log "Disk cleanup running for $([int]$runtime.TotalMinutes) minutes and $([int]($runtime.TotalSeconds % 60)) seconds..." -Level Info-Process -Id $cleanmgrProcess.Id -ErrorAction SilentlyContinue
                    } else {           if ($process) {
                        Write-Log "Disk cleanup running for $([int]$runtime.TotalSeconds) seconds..." -Level Info                $memUsage = [math]::Round($process.WorkingSet64 / 1MB, 2)
                    }mory usage: $memUsage MB" -Level Info
                    
                    try {        }
                        $process = Get-Process -Id $cleanmgrProcess.Id -ErrorAction SilentlyContinue
                        if ($process) {
                            $memUsage = [math]::Round($process.WorkingSet64 / 1MB, 2)           Write-Log "Unable to get process info: $_" -Level Debug
                            Write-Log "Current memory usage: $memUsage MB" -Level Info  }
                        }
                    }   
                    catch {       Start-Sleep -Seconds 5
                        # Process might have exited between checks
                        Write-Log "Unable to get process info: $_" -Level Debug
                    }cleanmgrProcess.ExitCode
                }   Write-Log "Disk cleanup process completed with exit code: $exitCode" -Level Verbose
                            
                Start-Sleep -Seconds 5
            }Info
            
            $exitCode = $cleanmgrProcess.ExitCode
            Write-Log "Disk cleanup process completed with exit code: $exitCode" -Level Verbosero exit code: $exitCode" -Level Warning
            
            if ($exitCode -eq 0) {
                Write-Log "Disk cleanup completed successfully" -Level Info
            } Error
            else {
                Write-Log "Disk cleanup completed with non-zero exit code: $exitCode" -Level Warning
            }
        }
        catch {emoving temporary registry settings" -Level Verbose
            Write-Log "Error monitoring disk cleanup process: $_" -Level Error   try {
            return $false Remove-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\*' `
        }
 "Registry cleanup completed" -Level Debug
        # Clean up registry settings   }
        Write-Log "Removing temporary registry settings" -Level Verbose       catch {
        try {            Write-Log "Error during registry cleanup: $_" -Level Warning
            Remove-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\*' `ven if registry cleanup fails
                -Name StateFlags0001 -Force -ErrorAction Stop
            Write-Log "Registry cleanup completed" -Level Debug    
        }rite-Log "Disk Cleanup completed successfully." -Level Info
        catch {
            Write-Log "Error during registry cleanup: $_" -Level Warning
            # Continue execution even if registry cleanup fails
        }
        return $false
        Write-Log "Disk Cleanup completed successfully." -Level Info
        return $true
    }
    catch {
        Write-Log "Error during Disk Cleanup: $_" -Level Errornfo
        return $false
    }
}
eving list of shadow copies" -Level Verbose
function Start-ShadowCopyCleanup {vssList = vssadmin list shadows
    Write-Log "Starting Shadow Copy cleanup process..." -Level Info        Write-Log "Shadow copy details: $vssList" -Level Debug
    
    try {Copy ID:"}
        # List all shadow copies() }
        Write-Log "Retrieving list of shadow copies" -Level Verbose
        $vssList = vssadmin list shadowspy details
        Write-Log "Shadow copy details: $vssList" -Level Debugnfo -Level Info
        nfo
        $shadowCopies = $vssList | Where-Object {$_ -match "Shadow Copy ID:"}
        $shadowIds = $shadowCopies | ForEach-Object { $_.Split(":")[1].Trim() }rning

        $shadowCount = $shadowIds.Count
        Write-Log "Found $shadowCount shadow copies" -Level Infonfo
        se creation dates to display them to user
        if ($shadowCount -eq 0) {dateLines = $vssList | Where-Object {$_ -match "Created:"}
            Write-Log "No shadow copies found." -Level Warning        $dates = $dateLines | ForEach-Object { $_.Split(":", 2)[1].Trim() }
            return $true
        }s with dates
y Details:" -Level Info
        # Parse creation dates to display them to userfo
        $dateLines = $vssList | Where-Object {$_ -match "Created:"}i++) {
        $dates = $dateLines | ForEach-Object { $_.Split(":", 2)[1].Trim() }el Info
                    Write-Log "Created: $($dates[$i])" -Level Info
        # Show all shadow copies with datesount - 1) {
        Write-Log "Shadow Copy Details:" -Level Info "-------------------" -Level Info
        Write-Log "-------------------" -Level Info            }
        for ($i = 0; $i -lt $shadowCount; $i++) {
            Write-Log "ID: $($shadowIds[$i])" -Level Info
            Write-Log "Created: $($dates[$i])" -Level Infoestore point
            if ($i -lt $shadowCount - 1) {g shadow copy ID: $id (Created: $($dates[$i]))" -Level Info
                Write-Log "-------------------" -Level Info
            }ut = vssadmin delete shadows /shadow=$id /quieteserving most recent shadow copy:" -Level Info
        }l Info

        # Keep only the newest restore point
        $keepId = $shadowIds[0] {eleted shadows
        $keepDate = $dates[0]og "Error deleting shadow copy ${id}: ${_}" -Level Warning 0
        Write-Log "Preserving most recent shadow copy:" -Level Info
        Write-Log "ID: $keepId" -Level Infostore points
        Write-Log "Created: $keepDate" -Level Infoi in 0..($shadowIds.Count-1)) {
 Enhanced summary with emoji for visibility   $id = $shadowIds[$i]
        # Count deleted shadowsWrite-Log "📊 Shadow Copy cleanup summary: $deletedCount copies removed, 1 preserved" -Level Info    if ($id -ne $keepId) {
        $deletedCount = 0

        # Delete older restore pointsatch {               $output = vssadmin delete shadows /shadow=$id /quiet
        foreach ($i in 0..($shadowIds.Count-1)) {te-Log "Error during Shadow Copy cleanup: $_" -Level Error         Write-Log "Deleted shadow copy from $($dates[$i])" -Level Info
            $id = $shadowIds[$i]
            if ($id -ne $keepId) {
                Write-Log "Removing shadow copy ID: $id (Created: $($dates[$i]))" -Level Verbose       catch {
                try {                  Write-Log "Error deleting shadow copy ${id}: ${_}" -Level Warning
                    $output = vssadmin delete shadows /shadow=$id /quietfunction Show-DriveInfo {                }
                    Write-Log "Deleted shadow copy from $($dates[$i])" -Level Info
                    $deletedCount++rameter(Mandatory=$true)]
                }
                catch {ummary: $deletedCount copies removed, 1 preserved" -Level Info
                    Write-Log "Error deleting shadow copy ${id}: ${_}" -Level Warning[Parameter(Mandatory=$false)]return $true
                }
            }
        }  Write-Log "Error during Shadow Copy cleanup: $_" -Level Error
        Write-Log "`n$State Drive Volume Details:" -Level Info    return $false
        Write-Log "Shadow Copy cleanup summary: $deletedCount copies removed, 1 preserved" -Level Info
        return $true -Level Info
    }nfo
    catch {
        Write-Log "Error during Shadow Copy cleanup: $_" -Level Error
        return $falseB" -Level Info
    }GB" -Level Info
}

function Show-DriveInfo {# Additional debug information    [string]$State = "Current"
    param (olume | Out-String)" -Level Debug
        [Parameter(Mandatory=$true)]
        [object]$Volume,  Write-Log "`n$State Drive Volume Details:" -Level Info
        # Main execution    Write-Log "------------------------" -Level Info
        [Parameter(Mandatory=$false)]ng enabled. Log file: $script:LogFile" -Level Inforive Letter: $($Volume.DriveLetter)" -Level Info
        [string]$State = "Current"
    )Write-Log "Script started with PowerShell version $($PSVersionTable.PSVersion)" -Level Verbose    Write-Log "File System: $($Volume.FileSystem)" -Level Info
    
    Write-Log "`n$State Drive Volume Details:" -Level InfoOperatingSystem).Caption)" -Level Verboseevel Info
    Write-Log "------------------------" -Level Info
    Write-Log "Drive Letter: $($Volume.DriveLetter)" -Level Infoif (-not (Test-RunningAsSystem) -and -not $NoElevate) {    Write-Log "Health Status: $($Volume.HealthStatus)" -Level Info
    Write-Log "Drive Label: $($Volume.FileSystemLabel)" -Level InfoEM" -Level Info
    Write-Log "File System: $($Volume.FileSystem)" -Level Info
    Write-Log "Drive Type: $($Volume.DriveType)" -Level Infoe | Out-String)" -Level Debug
    Write-Log "Size: $([math]::Round($Volume.Size/1GB, 2)) GB" -Level Info
    Write-Log "Free Space: $([math]::Round($Volume.SizeRemaining/1GB, 2)) GB" -Level Info
    Write-Log "Health Status: $($Volume.HealthStatus)" -Level InfoWrite-Log "Executing as SYSTEM account" -Level Info# Main execution
    el Infole" -Level Info
    # Additional debug information
    Write-Log "Drive details: $($Volume | Out-String)" -Level Debug# Get drive information before cleanupWrite-Log "Script started with PowerShell version $($PSVersionTable.PSVersion)" -Level Verbose
}
rite-Log "Collecting drive information before cleanup" -Level Verbose-Log "Operating system: $((Get-CimInstance -ClassName Win32_OperatingSystem).Caption)" -Level Verbose
# Main execution
Write-Log "Logging enabled. Log file: $script:LogFile" -Level Info# Get all available volumes with drive letters and sort them-not (Test-RunningAsSystem) -and -not $NoElevate) {

Write-Log "Script started with PowerShell version $($PSVersionTable.PSVersion)" -Level VerboseiveLetter } | 
Write-Log "Running on computer: $env:COMPUTERNAME" -Level Verbose
Write-Log "Operating system: $((Get-CimInstance -ClassName Win32_OperatingSystem).Caption)" -Level Verbose
    if ($volumes.Count -eq 0) {
if (-not (Test-RunningAsSystem) -and -not $NoElevate) {th letters found on the system." -Level Erroraccount" -Level Info
    Write-Log "Initial execution - will elevate to SYSTEM" -Level Info
    Start-SystemContext
    exite information before cleanup
}    Write-Log "Found $($volumes.Count) volumes with drive letters" -Level Debugtry {

Write-Log "Executing as SYSTEM account" -Level Info# Select the volume with lowest drive letter
Write-Log "Starting system storage cleanup..." -Level Infot them

# Get drive information before cleanupWrite-Log "Found lowest drive letter: $($volumeBeforeCleanup.DriveLetter)" -Level Info    Where-Object { $_.DriveLetter } | 
try {
    Write-Log "Collecting drive information before cleanup" -Level Verbose
    atch {   if ($volumes.Count -eq 0) {
    # Get all available volumes with drive letters and sort themte-Log "Error accessing drive information. Error: $_" -Level Error Write-Log "No drives with letters found on the system." -Level Error
    $volumes = Get-Volume | 
        Where-Object { $_.DriveLetter } |   }
        Sort-Object DriveLetter# Perform cleanup operations
 operations" -Level Verbosemes.Count) volumes with drive letters" -Level Debug
    if ($volumes.Count -eq 0) {
        Write-Log "No drives with letters found on the system." -Level Erroreanupve letter
        exit
    }if ($diskCleanupSuccess -and $shadowCopySuccess) {    
ccessfully!" -Level InfoeBeforeCleanup.DriveLetter)" -Level Info
    Write-Log "Found $($volumes.Count) volumes with drive letters" -Level Debug
    lse {
    # Select the volume with lowest drive letterite-Log "System storage cleanup encountered issues. Please check the logs." -Level Error{
    $volumeBeforeCleanup = $volumes[0]
    
    Write-Log "Found lowest drive letter: $($volumeBeforeCleanup.DriveLetter)" -Level Info# Get drive information after cleanup
    Show-DriveInfo -Volume $volumeBeforeCleanup -State "Before Cleanup"
}rite-Log "Collecting drive information after cleanup" -Level Verbose-Log "Beginning cleanup operations" -Level Verbose
catch {
    Write-Log "Error accessing drive information. Error: $_" -Level Error# Get all available volumes with drive letters and sort themdowCopySuccess = Start-ShadowCopyCleanup
}
iveLetter } |  $shadowCopySuccess) {
# Perform cleanup operationsssfully!" -Level Info
Write-Log "Beginning cleanup operations" -Level Verbose
$diskCleanupSuccess = Start-DiskCleanup    if ($volumes.Count -eq 0) {else {
$shadowCopySuccess = Start-ShadowCopyCleanupth letters found on the system." -Level Errorleanup encountered issues. Please check the logs." -Level Error

if ($diskCleanupSuccess -and $shadowCopySuccess) {
    Write-Log "System storage cleanup completed successfully!" -Level Infoe information after cleanup
}    # Select the volume with lowest drive lettertry {
else { Verbose
    Write-Log "System storage cleanup encountered issues. Please check the logs." -Level Error
}Write-Log "Found lowest drive letter: $($lowestVolume.DriveLetter)" -Level Info# Get all available volumes with drive letters and sort them

# Get drive information after cleanup
try {# Calculate and display space freed    Sort-Object DriveLetter
    Write-Log "Collecting drive information after cleanup" -Level Verbose
    beforeFreeSpace = $volumeBeforeCleanup.SizeRemainingvolumes.Count -eq 0) {
    # Get all available volumes with drive letters and sort theml Error
    $volumes = Get-Volume | pace) / 1GB
        Where-Object { $_.DriveLetter } | 
        Sort-Object DriveLetterif ($spaceSaved -gt 0) {
ed by cleanup: $([math]::Round($spaceSaved, 2)) GB" -Level Infoest drive letter
    if ($volumes.Count -eq 0) {
        Write-Log "No drives with letters found on the system." -Level Errorlse {
        exitite-Log "No measurable space was freed during cleanup" -Level Warning"Found lowest drive letter: $($lowestVolume.DriveLetter)" -Level Info
    }

    # Select the volume with lowest drive letteratch { Calculate and display space freed
    $lowestVolume = $volumes[0]te-Log "Unable to calculate space saved: $_" -Level Debug
    
    Write-Log "Found lowest drive letter: $($lowestVolume.DriveLetter)" -Level InfoafterFreeSpace = $lowestVolume.SizeRemaining
    Show-DriveInfo -Volume $lowestVolume -State "After Cleanup"atch {       $spaceSaved = ($afterFreeSpace - $beforeFreeSpace) / 1GB
    te-Log "Error accessing drive information. Error: $_" -Level Error 
    # Calculate and display space freed
    try {          Write-Log "Space freed by cleanup: $([math]::Round($spaceSaved, 2)) GB" -Level Info
        $beforeFreeSpace = $volumeBeforeCleanup.SizeRemaining# Display cleanup summary        }
        $afterFreeSpace = $lowestVolume.SizeRemainingary:" -Level Info
        $spaceSaved = ($afterFreeSpace - $beforeFreeSpace) / 1GBs freed during cleanup" -Level Warning
        
        if ($spaceSaved -gt 0) {# Replace ternary operators with standard if-else for PowerShell 5.1 compatibility    }
            Write-Log "Space freed by cleanup: $([math]::Round($spaceSaved, 2)) GB" -Level Info "Failed" }
        }
        else {
            Write-Log "No measurable space was freed during cleanup" -Level Warning
        }$shadowCopyMessage = if ($shadowCopySuccess) { "Completed Successfully" } else { "Failed" }catch {
    }
    catch {Level
        Write-Log "Unable to calculate space saved: $_" -Level Debug
    }Write-Log "Script execution completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Level Info# Display cleanup summary
}
catch {Write-Log "Log file created at: $script:LogFile" -Level Info# Replace ternary operators with standard if-else for PowerShell 5.1 compatibility
    Write-Log "Error accessing drive information. Error: $_" -Level ErrorSuccessfully" } else { "Failed" }


















Write-Log "Log file created at: $script:LogFile" -Level InfoWrite-Log "Script execution completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Level InfoWrite-Log "Shadow Copy Cleanup: $shadowCopyMessage" -Level $shadowCopyLevel$shadowCopyLevel = if ($shadowCopySuccess) { "Info" } else { "Error" }$shadowCopyMessage = if ($shadowCopySuccess) { "Completed Successfully" } else { "Failed" }Write-Log "Disk Cleanup: $diskCleanupMessage" -Level $diskCleanupLevel$diskCleanupLevel = if ($diskCleanupSuccess) { "Info" } else { "Error" }$diskCleanupMessage = if ($diskCleanupSuccess) { "Completed Successfully" } else { "Failed" }# Replace ternary operators with standard if-else for PowerShell 5.1 compatibilityWrite-Log "---------------" -Level InfoWrite-Log "`nCleanup Summary:" -Level Info# Display cleanup summary}$diskCleanupLevel = if ($diskCleanupSuccess) { "Info" } else { "Error" }
Write-Log "Disk Cleanup: $diskCleanupMessage" -Level $diskCleanupLevel

$shadowCopyMessage = if ($shadowCopySuccess) { "Completed Successfully" } else { "Failed" }
$shadowCopyLevel = if ($shadowCopySuccess) { "Info" } else { "Error" }
Write-Log "Shadow Copy Cleanup: $shadowCopyMessage" -Level $shadowCopyLevel

Write-Log "Script execution completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Level Info

Write-Log "Log file created at: $script:LogFile" -Level Info