# =============================================================================
# Script: Clear-SystemStorage.ps1
# Created: 2025-02-27 18:55:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-06 15:25:00 UTC
# Updated By: jdyer-nuvodia
# Version: 2.5
# Additional Info: Fixed SYSTEM context elevation for OneDrive paths
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
    
    # Write to console with appropriate color
    switch ($Level) {
        'Info' { 
            Write-Host $Message
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
        
        $scriptDirectory = $systemAccessibleTemp
        $logFile = Join-Path $scriptDirectory "$jobName.log"
        $markerFile = Join-Path $scriptDirectory "$jobName.marker"
        
        Write-Log "Creating temporary job files in $scriptDirectory" -Level Verbose
        
        # Create parameters string to pass logging preference
        $paramString = "-NoElevate"
        if ($VerbosePreference -eq 'Continue') { $paramString += " -Verbose" }
        if ($DebugPreference -eq 'Continue') { $paramString += " -Debug" }
        
        # Create a wrapper script that calls the copied script with parameters
        $tempScriptPath = Join-Path $scriptDirectory "$jobName.ps1"
        $wrapperContent = @"
# Temporary wrapper script generated for SYSTEM context execution
try {
    # Execute the copied script with parameters
    & '$systemAccessibleScriptPath' $paramString
} 
catch {
    Write-Error "Error executing script as SYSTEM: `$_"
    exit 1
}
"@
        
        Write-Log "Creating wrapper script at $tempScriptPath" -Level Debug
        Set-Content -Path $tempScriptPath -Value $wrapperContent -Force
        
        # Create action with logging to execute the temp script
        $argument = "-NoProfile -ExecutionPolicy Bypass -Command `"& {Start-Transcript '$logFile'; & '$tempScriptPath'; Set-Content -Path '$markerFile' -Value 'Complete' -Force; Stop-Transcript}`""
        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $argument
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
        Write-Log "Waiting for system cleanup to complete..." -Level Info
        
        while ((Get-Date) -lt $timeout -and -not $completed) {
            # Check for completion marker file first
            if (Test-Path $markerFile) {
                $completed = $true
                Write-Log "Completion marker found" -Level Verbose
                continue
            }
            
            # Then check task status
            try {
                $status = Get-ScheduledTask -TaskName $jobName -ErrorAction SilentlyContinue
                if ($status -and $status.State -eq "Ready") {
                    $completed = $true
                    Write-Log "Task completed according to scheduler" -Level Verbose
                    continue
                }
            } catch {
                # Task might be completed and already removed
                if (Test-Path $markerFile) {
                    $completed = $true
                    continue
                }
            }
            
            # Display log updates without duplication
            if (Test-Path $logFile) {
                $currentLines = Get-Content $logFile -Tail 10
                foreach ($line in $currentLines) {
                    $trimmedLine = $line.Trim()
                    if ($trimmedLine -and 
                        -not $seenLogLines.ContainsKey($trimmedLine) -and 
                        $trimmedLine -notmatch "^Transcript started|^Transcript ended|^Windows PowerShell transcript") {
                        Write-Host $trimmedLine
                        $seenLogLines[$trimmedLine] = $true
                    }
                }
            }
            
            Start-Sleep -Seconds 3
        }

        # Check final status
        if (-not $completed) {
            Write-Log "Task did not complete within timeout period" -Level Error
        }

        # Cleanup
        Write-Log "Cleaning up temporary files and tasks" -Level Verbose
        if (Get-ScheduledTask -TaskName $jobName -ErrorAction SilentlyContinue) {
            Unregister-ScheduledTask -TaskName $jobName -Confirm:$false
        }
        
        # Cleanup files
        if (Test-Path $logFile) { Remove-Item $logFile -Force -ErrorAction SilentlyContinue }
        if (Test-Path $markerFile) { Remove-Item $markerFile -Force -ErrorAction SilentlyContinue }
        if (Test-Path $tempScriptPath) { Remove-Item $tempScriptPath -Force -ErrorAction SilentlyContinue }
        if (Test-Path $systemAccessibleScriptPath) { Remove-Item $systemAccessibleScriptPath -Force -ErrorAction SilentlyContinue }
        
        if ($completed) { exit 0 } else { exit 1 }
    }
    catch {
        Write-Log "Failed to elevate to SYSTEM context: $_" -Level Error
        return $false
    }
}

function Start-DiskCleanup {
    Write-Log "Starting Disk Cleanup process..." -Level Info
    
    try {
        # Create the StateFlags registry key
        Write-Log "Setting registry keys for automatic cleanup" -Level Verbose
        try {
            New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\*' `
                -Name StateFlags0001 -Value 2 -PropertyType DWord -Force -ErrorAction Stop | Out-Null
            Write-Log "Registry keys set successfully" -Level Verbose
        }
        catch {
            Write-Log "Error setting registry keys: $_" -Level Error
            return $false
        }

        # Run Disk Cleanup silently with LOWDISK parameter to prevent GUI
        Write-Log "Executing cleanmgr.exe with /sagerun:1 /LOWDISK parameters" -Level Verbose
        
        try {
            $cleanmgrProcess = Start-Process -FilePath cleanmgr -ArgumentList '/sagerun:1 /LOWDISK' -WindowStyle Hidden -PassThru
            
            # Monitor the process with updates
            Write-Log "Disk cleanup process started with ID: $($cleanmgrProcess.Id)" -Level Debug
            $startTime = Get-Date
            
            while (-not $cleanmgrProcess.HasExited) {
                $runtime = (Get-Date) - $startTime
                if ($runtime.TotalSeconds -gt 180) {
                    Write-Log "Disk cleanup has been running for $([int]$runtime.TotalMinutes) minutes..." -Level Verbose
                }
                
                # Get process memory usage to show it's still working
                try {
                    $process = Get-Process -Id $cleanmgrProcess.Id -ErrorAction SilentlyContinue
                    if ($process) {
                        $memUsage = [math]::Round($process.WorkingSet64 / 1MB, 2)
                        Write-Log "Disk cleanup still running. Memory usage: $memUsage MB" -Level Debug
                    }
                }
                catch {
                    # Process might have exited between checks
                    Write-Log "Unable to get process info: $_" -Level Debug
                }
                
                Start-Sleep -Seconds 10
            }
            
            $exitCode = $cleanmgrProcess.ExitCode
            Write-Log "Disk cleanup process completed with exit code: $exitCode" -Level Verbose
            
            if ($exitCode -eq 0) {
                Write-Log "Disk cleanup completed successfully" -Level Info
            }
            else {
                Write-Log "Disk cleanup completed with non-zero exit code: $exitCode" -Level Warning
            }
        }
        catch {
            Write-Log "Error monitoring disk cleanup process: $_" -Level Error
            return $false
        }

        # Clean up registry settings
        Write-Log "Removing temporary registry settings" -Level Verbose
        try {
            Remove-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\*' `
                -Name StateFlags0001 -Force -ErrorAction Stop
            Write-Log "Registry cleanup completed" -Level Debug
        }
        catch {
            Write-Log "Error during registry cleanup: $_" -Level Warning
            # Continue execution even if registry cleanup fails
        }
        
        Write-Log "Disk Cleanup completed successfully." -Level Info
        return $true
    }
    catch {
        Write-Log "Error during Disk Cleanup: $_" -Level Error
        return $false
    }
}

function Start-ShadowCopyCleanup {
    Write-Log "Starting Shadow Copy cleanup process..." -Level Info
    
    try {
        # List all shadow copies
        Write-Log "Retrieving list of shadow copies" -Level Verbose
        $vssList = vssadmin list shadows
        Write-Log "Shadow copy details: $vssList" -Level Debug
        
        $shadowCopies = $vssList | Where-Object {$_ -match "Shadow Copy ID:"}
        $shadowIds = $shadowCopies | ForEach-Object { $_.Split(":")[1].Trim() }

        $shadowCount = $shadowIds.Count
        Write-Log "Found $shadowCount shadow copies" -Level Info
        
        if ($shadowCount -eq 0) {
            Write-Log "No shadow copies found." -Level Warning
            return $true
        }

        # Parse creation dates to display them to user
        $dateLines = $vssList | Where-Object {$_ -match "Created:"}
        $dates = $dateLines | ForEach-Object { $_.Split(":", 2)[1].Trim() }
        
        # Show all shadow copies with dates
        Write-Log "Shadow Copy Details:" -Level Info
        Write-Log "-------------------" -Level Info
        for ($i = 0; $i -lt $shadowCount; $i++) {
            Write-Log "ID: $($shadowIds[$i])" -Level Info
            Write-Log "Created: $($dates[$i])" -Level Info
            if ($i -lt $shadowCount - 1) {
                Write-Log "-------------------" -Level Info
            }
        }

        # Keep only the newest restore point
        $keepId = $shadowIds[0]
        $keepDate = $dates[0]
        Write-Log "Preserving most recent shadow copy:" -Level Info
        Write-Log "ID: $keepId" -Level Info
        Write-Log "Created: $keepDate" -Level Info

        # Count deleted shadows
        $deletedCount = 0

        # Delete older restore points
        foreach ($i in 0..($shadowIds.Count-1)) {
            $id = $shadowIds[$i]
            if ($id -ne $keepId) {
                Write-Log "Removing shadow copy ID: $id (Created: $($dates[$i]))" -Level Verbose
                try {
                    $output = vssadmin delete shadows /shadow=$id /quiet
                    Write-Log "Deleted shadow copy from $($dates[$i])" -Level Info
                    $deletedCount++
                }
                catch {
                    Write-Log "Error deleting shadow copy ${id}: ${_}" -Level Warning
                }
            }
        }
        
        Write-Log "Shadow Copy cleanup summary: $deletedCount copies removed, 1 preserved" -Level Info
        return $true
    }
    catch {
        Write-Log "Error during Shadow Copy cleanup: $_" -Level Error
        return $false
    }
}

function Show-DriveInfo {
    param (
        [Parameter(Mandatory=$true)]
        [object]$Volume,
        
        [Parameter(Mandatory=$false)]
        [string]$State = "Current"
    )
    
    Write-Log "`n$State Drive Volume Details:" -Level Info
    Write-Log "------------------------" -Level Info
    Write-Log "Drive Letter: $($Volume.DriveLetter)" -Level Info
    Write-Log "Drive Label: $($Volume.FileSystemLabel)" -Level Info
    Write-Log "File System: $($Volume.FileSystem)" -Level Info
    Write-Log "Drive Type: $($Volume.DriveType)" -Level Info
    Write-Log "Size: $([math]::Round($Volume.Size/1GB, 2)) GB" -Level Info
    Write-Log "Free Space: $([math]::Round($Volume.SizeRemaining/1GB, 2)) GB" -Level Info
    Write-Log "Health Status: $($Volume.HealthStatus)" -Level Info
    
    # Additional debug information
    Write-Log "Drive details: $($Volume | Out-String)" -Level Debug
}

# Main execution
Write-Log "Logging enabled. Log file: $script:LogFile" -Level Info

Write-Log "Script started with PowerShell version $($PSVersionTable.PSVersion)" -Level Verbose
Write-Log "Running on computer: $env:COMPUTERNAME" -Level Verbose
Write-Log "Operating system: $((Get-CimInstance -ClassName Win32_OperatingSystem).Caption)" -Level Verbose

if (-not (Test-RunningAsSystem) -and -not $NoElevate) {
    Write-Log "Initial execution - will elevate to SYSTEM" -Level Info
    Start-SystemContext
    exit
}

Write-Log "Executing as SYSTEM account" -Level Info
Write-Log "Starting system storage cleanup..." -Level Info

# Get drive information before cleanup
try {
    Write-Log "Collecting drive information before cleanup" -Level Verbose
    
    # Get all available volumes with drive letters and sort them
    $volumes = Get-Volume | 
        Where-Object { $_.DriveLetter } | 
        Sort-Object DriveLetter

    if ($volumes.Count -eq 0) {
        Write-Log "No drives with letters found on the system." -Level Error
        exit
    }

    Write-Log "Found $($volumes.Count) volumes with drive letters" -Level Debug
    
    # Select the volume with lowest drive letter
    $volumeBeforeCleanup = $volumes[0]
    
    Write-Log "Found lowest drive letter: $($volumeBeforeCleanup.DriveLetter)" -Level Info
    Show-DriveInfo -Volume $volumeBeforeCleanup -State "Before Cleanup"
}
catch {
    Write-Log "Error accessing drive information. Error: $_" -Level Error
}

# Perform cleanup operations
Write-Log "Beginning cleanup operations" -Level Verbose
$diskCleanupSuccess = Start-DiskCleanup
$shadowCopySuccess = Start-ShadowCopyCleanup

if ($diskCleanupSuccess -and $shadowCopySuccess) {
    Write-Log "System storage cleanup completed successfully!" -Level Info
}
else {
    Write-Log "System storage cleanup encountered issues. Please check the logs." -Level Error
}

# Get drive information after cleanup
try {
    Write-Log "Collecting drive information after cleanup" -Level Verbose
    
    # Get all available volumes with drive letters and sort them
    $volumes = Get-Volume | 
        Where-Object { $_.DriveLetter } | 
        Sort-Object DriveLetter

    if ($volumes.Count -eq 0) {
        Write-Log "No drives with letters found on the system." -Level Error
        exit
    }

    # Select the volume with lowest drive letter
    $lowestVolume = $volumes[0]
    
    Write-Log "Found lowest drive letter: $($lowestVolume.DriveLetter)" -Level Info
    Show-DriveInfo -Volume $lowestVolume -State "After Cleanup"
    
    # Calculate and display space freed
    try {
        $beforeFreeSpace = $volumeBeforeCleanup.SizeRemaining
        $afterFreeSpace = $lowestVolume.SizeRemaining
        $spaceSaved = ($afterFreeSpace - $beforeFreeSpace) / 1GB
        
        if ($spaceSaved -gt 0) {
            Write-Log "Space freed by cleanup: $([math]::Round($spaceSaved, 2)) GB" -Level Info
        }
        else {
            Write-Log "No measurable space was freed during cleanup" -Level Warning
        }
    }
    catch {
        Write-Log "Unable to calculate space saved: $_" -Level Debug
    }
}
catch {
    Write-Log "Error accessing drive information. Error: $_" -Level Error
}

# Display cleanup summary
Write-Log "`nCleanup Summary:" -Level Info
Write-Log "---------------" -Level Info
Write-Log "Disk Cleanup: $($diskCleanupSuccess ? "Completed Successfully" : "Failed")" -Level $($diskCleanupSuccess ? "Info" : "Error")
Write-Log "Shadow Copy Cleanup: $($shadowCopySuccess ? "Completed Successfully" : "Failed")" -Level $($shadowCopySuccess ? "Info" : "Error")

Write-Log "Script execution completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Level Info

Write-Log "Log file created at: $script:LogFile" -Level Info