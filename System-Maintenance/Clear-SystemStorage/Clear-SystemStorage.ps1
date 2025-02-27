# =============================================================================
# Script: Clear-SystemStorage.ps1
# Created: 2025-02-27 18:55:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-27 21:00:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.3
# Additional Info: Enhanced SYSTEM context execution with better completion tracking
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
    
    Dependencies:
    - Windows Volume Shadow Copy Service
    - Administrative privileges
.NOTES
    Security Level: High
    Required Permissions: Administrative privileges
    Validation Requirements: 
    - Verify shadow copy retention
    - Check disk space recovery
#>

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
        return $true
    }

    Write-Host "Elevating to SYSTEM context..." -ForegroundColor Cyan
    
    try {
        $jobName = "SystemContextJob_$([Guid]::NewGuid())"
        $logFile = Join-Path $env:TEMP "$jobName.log"
        
        # Create action with logging
        $argument = "-NoProfile -ExecutionPolicy Bypass -Command `"& {Start-Transcript '$logFile'; . '$ScriptPath'; Stop-Transcript}`""
        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $argument
        $principal = New-ScheduledTaskPrincipal -UserId "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
        $task = New-ScheduledTask -Action $action -Principal $principal -Settings $settings
        
        # Register task using the task object
        Register-ScheduledTask -TaskName $jobName -InputObject $task | Out-Null
        Start-ScheduledTask -TaskName $jobName

        # Wait for completion with progress bar
        $timeout = (Get-Date).AddSeconds($TimeoutSeconds)
        $completed = $false
        
        while ((Get-Date) -lt $timeout -and -not $completed) {
            $status = Get-ScheduledTask -TaskName $jobName
            $completed = $status.State -eq "Ready"
            
            Write-Host "Waiting for system cleanup to complete..." -ForegroundColor Cyan
            Start-Sleep -Seconds 5
            
            if (Test-Path $logFile) {
                Get-Content $logFile -Tail 1
            }
        }

        # Check final status
        if (-not $completed) {
            Write-Error "Task did not complete within timeout period"
        }
        elseif (Test-Path $logFile) {
            Get-Content $logFile | Write-Host
        }

        # Cleanup
        Unregister-ScheduledTask -TaskName $jobName -Confirm:$false
        if (Test-Path $logFile) { Remove-Item $logFile -Force }
        
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

# Main execution
if (-not (Test-RunningAsSystem)) {
    Start-SystemContext
    exit
}

Write-Host "Executing as SYSTEM account" -ForegroundColor Cyan
Write-Host "Starting system storage cleanup..." -ForegroundColor Cyan

$diskCleanupSuccess = Start-DiskCleanup
$shadowCopySuccess = Start-ShadowCopyCleanup

if ($diskCleanupSuccess -and $shadowCopySuccess) {
    Write-Host "System storage cleanup completed successfully!" -ForegroundColor Green
}
else {
    Write-Error "System storage cleanup encountered issues. Please check the logs."
}