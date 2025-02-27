# =============================================================================
# Script: Clear-SystemStorage.ps1
# Created: 2025-02-27 18:55:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-27 19:00:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1
# Additional Info: Fixed script formatting and structure
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

function Start-DiskCleanup {
    Write-Host "Starting Disk Cleanup process..." -ForegroundColor Cyan
    
    try {
        # Create the StateFlags registry key
        New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\*' `
            -Name StateFlags0001 -Value 2 -PropertyType DWord -Force | Out-Null

        # Run Disk Cleanup silently
        Start-Process -FilePath cleanmgr -ArgumentList "/sagerun:1" -WindowStyle Hidden -Wait

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
Write-Host "Starting system storage cleanup..." -ForegroundColor Cyan

$diskCleanupSuccess = Start-DiskCleanup
$shadowCopySuccess = Start-ShadowCopyCleanup

if ($diskCleanupSuccess -and $shadowCopySuccess) {
    Write-Host "System storage cleanup completed successfully!" -ForegroundColor Green
}
else {
    Write-Error "System storage cleanup encountered issues. Please check the logs."
}