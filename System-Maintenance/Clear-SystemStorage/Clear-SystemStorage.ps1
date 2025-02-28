# =============================================================================
# Script: Clear-SystemStorage.ps1
# Created: 2025-02-27 18:55:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-28 22:05:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.8
# Additional Info: Fixed parameter handling in system context execution
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
    
    Dependencies:
    - Windows Volume Shadow Copy Service
    - Administrative privileges
.PARAMETER NoElevate
    Prevents the script from attempting to elevate to SYSTEM context when already running as a scheduled task
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

function Test-RunningAsSystem {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    return $currentUser.User.Value -eq "S-1-5-18"
}

function Start-SystemContext {
    param(
        [string]$ScriptPath = $MyInvocation.PSCommandPath,
        [int]$TimeoutSeconds = 300
    )

    if (Test-RunningAsSystem) {
        Write-Host "Already running as SYSTEM account." -ForegroundColor Green
        return $true
    }

    Write-Host "Elevating to SYSTEM context..." -ForegroundColor Cyan
    
    try {
        $jobName = "SystemContextJob_$([Guid]::NewGuid())"
        $scriptDirectory = Split-Path -Parent $ScriptPath
        $logFile = Join-Path $scriptDirectory "$jobName.log"
        $markerFile = Join-Path $scriptDirectory "$jobName.marker"
        
        # Create a direct execution script that doesn't invoke the original script but contains the code
        $tempScriptPath = Join-Path $scriptDirectory "$jobName.ps1"
        
        # Read the original script and properly handle the param block
        $scriptContent = Get-Content -Path $ScriptPath -Raw
        
        # Extract the param block if it exists and add NoElevate parameter after it
        if ($scriptContent -match '(?s)^(param\s*\([^)]*\))') {
            $paramBlock = $Matches[1]
            $modifiedContent = $scriptContent -replace [regex]::Escape($paramBlock), "$paramBlock`n`n# Added by system context elevation`n`$NoElevate = `$true"
        } else {
            # If no param block, add it at the beginning
            $modifiedContent = "param(`n    [switch]`$NoElevate`n)`n`n`$NoElevate = `$true`n`n$scriptContent"
        }
        
        # Write modified content to temp file
        Set-Content -Path $tempScriptPath -Value $modifiedContent -Force
        
        # Create action with logging to execute the temp script
        $argument = "-NoProfile -ExecutionPolicy Bypass -Command `"& {Start-Transcript '$logFile'; & '$tempScriptPath'; Set-Content -Path '$markerFile' -Value 'Complete' -Force; Stop-Transcript}`""
        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $argument
        $principal = New-ScheduledTaskPrincipal -UserId "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
        $task = New-ScheduledTask -Action $action -Principal $principal -Settings $settings
        
        # Register task using the task object
        Register-ScheduledTask -TaskName $jobName -InputObject $task | Out-Null
        Start-ScheduledTask -TaskName $jobName

        # Wait for completion with progress indicator
        $timeout = (Get-Date).AddSeconds($TimeoutSeconds)
        $completed = $false
        $seenLogLines = @{}
        Write-Host "Waiting for system cleanup to complete..." -ForegroundColor Cyan
        
        while ((Get-Date) -lt $timeout -and -not $completed) {
            # Check for completion marker file first
            if (Test-Path $markerFile) {
                $completed = $true
                continue
            }
            
            # Then check task status
            try {
                $status = Get-ScheduledTask -TaskName $jobName -ErrorAction SilentlyContinue
                if ($status -and $status.State -eq "Ready") {
                    $completed = $true
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
            Write-Error "Task did not complete within timeout period"
        }

        # Cleanup
        if (Get-ScheduledTask -TaskName $jobName -ErrorAction SilentlyContinue) {
            Unregister-ScheduledTask -TaskName $jobName -Confirm:$false
        }
        
        # Cleanup files
        if (Test-Path $logFile) { Remove-Item $logFile -Force -ErrorAction SilentlyContinue }
        if (Test-Path $markerFile) { Remove-Item $markerFile -Force -ErrorAction SilentlyContinue }
        if (Test-Path $tempScriptPath) { Remove-Item $tempScriptPath -Force -ErrorAction SilentlyContinue }
        
        if ($completed) { exit 0 } else { exit 1 }
    }
    catch {
        Write-Error "Failed to elevate to SYSTEM context: $_"
        return $false
    }
}

function Start-DiskCleanup {
    Write-Host "Starting Disk Cleanup process..." -ForegroundColor Cyan
    
    try {
        # Create the StateFlags registry key
        New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\*' `
            -Name StateFlags0001 -Value 2 -PropertyType DWord -Force | Out-Null

        # Run Disk Cleanup silently with LOWDISK parameter to prevent GUI
        Start-Process -FilePath cleanmgr -ArgumentList '/sagerun:1 /LOWDISK' -WindowStyle Hidden -Wait

        # Clean up registry settings
        Remove-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\*' `
            -Name StateFlags0001 -Force
        
        Write-Host "Disk Cleanup completed successfully." -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Error during Disk Cleanup: $_"
        return $false
    }
}

function Start-ShadowCopyCleanup {
    Write-Host "Starting Shadow Copy cleanup process..." -ForegroundColor Cyan
    
    try {
        # List all shadow copies
        $vssList = vssadmin list shadows
        $shadowCopies = $vssList | Where-Object {$_ -match "Shadow Copy ID:"}
        $shadowIds = $shadowCopies | ForEach-Object { $_.Split(":")[1].Trim() }

        if ($shadowIds.Count -eq 0) {
            Write-Host "No shadow copies found." -ForegroundColor Yellow
            return $true
        }

        # Keep only the newest restore point
        $keepId = $shadowIds[0]
        Write-Host "Preserving most recent shadow copy ID: $keepId" -ForegroundColor Cyan

        # Delete older restore points
        foreach ($id in $shadowIds | Where-Object {$_ -ne $keepId}) {
            Write-Host "Removing shadow copy ID: $id" -ForegroundColor Yellow
            vssadmin delete shadows /shadow=$id /quiet
        }
        
        Write-Host "Shadow Copy cleanup completed successfully." -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Error during Shadow Copy cleanup: $_"
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
    
    Write-Host "`n$State Drive Volume Details:" -ForegroundColor Green
    Write-Host "------------------------" -ForegroundColor Green
    Write-Host "Drive Letter: $($Volume.DriveLetter)" -ForegroundColor Cyan
    Write-Host "Drive Label: $($Volume.FileSystemLabel)" -ForegroundColor Cyan
    Write-Host "File System: $($Volume.FileSystem)" -ForegroundColor Cyan
    Write-Host "Drive Type: $($Volume.DriveType)" -ForegroundColor Cyan
    Write-Host "Size: $([math]::Round($Volume.Size/1GB, 2)) GB" -ForegroundColor Cyan
    Write-Host "Free Space: $([math]::Round($Volume.SizeRemaining/1GB, 2)) GB" -ForegroundColor Cyan
    Write-Host "Health Status: $($Volume.HealthStatus)" -ForegroundColor Cyan
}

# Main execution
if (-not (Test-RunningAsSystem) -and -not $NoElevate) {
    Write-Host "Initial execution - will elevate to SYSTEM" -ForegroundColor Cyan
    Start-SystemContext
    exit
}

Write-Host "Executing as SYSTEM account" -ForegroundColor Cyan
Write-Host "Starting system storage cleanup..." -ForegroundColor Cyan

# Get drive information before cleanup
try {
    # Get all available volumes with drive letters and sort them
    $volumes = Get-Volume | 
        Where-Object { $_.DriveLetter } | 
        Sort-Object DriveLetter

    if ($volumes.Count -eq 0) {
        Write-Error "No drives with letters found on the system."
        exit
    }

    # Select the volume with lowest drive letter
    $lowestVolume = $volumes[0]
    
    Write-Host "Found lowest drive letter: $($lowestVolume.DriveLetter)" -ForegroundColor Yellow
    Show-DriveInfo -Volume $lowestVolume -State "Before Cleanup"
}
catch {
    Write-Error "Error accessing drive information. Error: $_"
}

# Perform cleanup operations
$diskCleanupSuccess = Start-DiskCleanup
$shadowCopySuccess = Start-ShadowCopyCleanup

if ($diskCleanupSuccess -and $shadowCopySuccess) {
    Write-Host "System storage cleanup completed successfully!" -ForegroundColor Green
}
else {
    Write-Error "System storage cleanup encountered issues. Please check the logs."
}

# Get drive information after cleanup
try {
    # Get all available volumes with drive letters and sort them
    $volumes = Get-Volume | 
        Where-Object { $_.DriveLetter } | 
        Sort-Object DriveLetter

    if ($volumes.Count -eq 0) {
        Write-Error "No drives with letters found on the system."
        exit
    }

    # Select the volume with lowest drive letter
    $lowestVolume = $volumes[0]
    
    Write-Host "Found lowest drive letter: $($lowestVolume.DriveLetter)" -ForegroundColor Yellow
    Show-DriveInfo -Volume $lowestVolume -State "After Cleanup"
}
catch {
    Write-Error "Error accessing drive information. Error: $_"
}