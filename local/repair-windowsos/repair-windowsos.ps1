# Windows OS Repair Script
# Version: 2.0
# Author: Original by jdyer-nuvodia, optimized with GitHub Copilot
# Last Updated: 2025-02-05
# Description: Performs comprehensive Windows system repairs and health checks with detailed logging

# Ensure we stop on errors
$ErrorActionPreference = 'Stop'

# Script Variables
$logFile = Join-Path $env:TEMP "WindowsRepair_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$repairsMade = $false
$restartNeeded = $false

# Function to write to both console and log file
function Write-RepairLog {
    param (
        [string]$Message,
        [string]$Color = 'White',
        [switch]$NoNewline
    )
    
    # Get current timestamp
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    # Write to console with color
    if ($NoNewline) {
        Write-Host $Message -ForegroundColor $Color -NoNewline
    } else {
        Write-Host $Message -ForegroundColor $Color
    }
    
    # Write to log file with timestamp
    "$timestamp - $Message" | Out-File -FilePath $logFile -Append
}

# Function to run commands and handle errors
function Invoke-RepairCommand {
    param (
        [string]$CommandName,
        [scriptblock]$ScriptBlock,
        [string]$SuccessMessage,
        [string]$ErrorMessage
    )
    
    Write-RepairLog "Starting: $CommandName" -Color Cyan
    try {
        $result = & $ScriptBlock
        Write-RepairLog $SuccessMessage -Color Green
        return $result
    }
    catch {
        Write-RepairLog "Error in $CommandName : $_" -Color Red
        Write-RepairLog $ErrorMessage -Color Red
        return $false
    }
}

# Check for Administrator privileges
Write-RepairLog "Checking administrator privileges..." -Color Cyan
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-RepairLog "This script requires administrator privileges. Please run as administrator." -Color Red
    exit 1
}

# Display initial system information
Write-RepairLog "=== Windows Repair Script Started ===" -Color Cyan
Write-RepairLog "System Information:" -Color Cyan
Write-RepairLog "Windows Version: $(([System.Environment]::OSVersion.Version).ToString())" -Color White
Write-RepairLog "Computer Name: $env:COMPUTERNAME" -Color White
Write-RepairLog "Log File Location: $logFile" -Color White
Write-RepairLog "----------------------------------------" -Color Cyan

# Step 1: DISM Health Check with progress indication
Write-RepairLog "`nStep 1/3: DISM Health Check" -Color Cyan
Write-RepairLog "Checking component store health..." -Color Yellow -NoNewline

$dismCheck = Invoke-RepairCommand -CommandName "DISM CheckHealth" -ScriptBlock {
    $output = DISM.exe /Online /Cleanup-Image /CheckHealth
    if ($output -match "No component store corruption detected") {
        return @{ Success = $true; NeedsRepair = $false }
    }
    return @{ Success = $true; NeedsRepair = $true }
} -SuccessMessage "DISM check completed." -ErrorMessage "DISM check failed."

if ($dismCheck.NeedsRepair) {
    Write-RepairLog "Component store corruption detected. Initiating repair..." -Color Yellow
    $dismRepair = Invoke-RepairCommand -CommandName "DISM Repair" -ScriptBlock {
        $output = DISM.exe /Online /Cleanup-Image /RestoreHealth
        if ($output -match "The restore operation completed successfully") {
            $script:repairsMade = $true
            return $true
        }
        return $false
    } -SuccessMessage "DISM repair completed successfully." -ErrorMessage "DISM repair encountered issues."
}

# Step 2: System File Checker with progress bar
Write-RepairLog "`nStep 2/3: System File Checker" -Color Cyan
$sfcResult = Invoke-RepairCommand -CommandName "System File Checker" -ScriptBlock {
    $output = sfc.exe /scannow
    $outputString = $output -join "`n"
    
    if ($outputString -match "Windows Resource Protection found corrupt files and successfully repaired them") {
        $script:repairsMade = $true
        $script:restartNeeded = $true
        return @{ Success = $true; RepairsNeeded = $true }
    }
    return @{ Success = $true; RepairsNeeded = $false }
} -SuccessMessage "SFC scan completed." -ErrorMessage "SFC scan encountered issues."

# Step 3: Check Disk with detailed reporting
Write-RepairLog "`nStep 3/3: Disk Health Check" -Color Cyan
$systemDrive = $env:SystemDrive
$chkdskResult = Invoke-RepairCommand -CommandName "CheckDisk" -ScriptBlock {
    $output = chkdsk.exe $systemDrive /scan
    if ($output -match "found no problems") {
        return @{ Success = $true; Problems = $false }
    }
    return @{ Success = $true; Problems = $true }
} -SuccessMessage "Disk health check completed." -ErrorMessage "Disk health check failed."

if ($chkdskResult.Problems) {
    Write-RepairLog "Disk issues detected. Scheduling full CHKDSK for next restart..." -Color Yellow
    $scheduleChkdsk = Invoke-RepairCommand -CommandName "Schedule CHKDSK" -ScriptBlock {
        chkdsk.exe $systemDrive /f /x
        $script:restartNeeded = $true
        $script:repairsMade = $true
        return $true
    } -SuccessMessage "CHKDSK scheduled for next restart." -ErrorMessage "Failed to schedule CHKDSK."
}

# Final Summary
Write-RepairLog "`n=== Repair Summary ===" -Color Cyan
Write-RepairLog "Repairs performed: $($repairsMade ? 'Yes' : 'No')" -Color ($repairsMade ? 'Yellow' : 'Green')
Write-RepairLog "Restart required: $($restartNeeded ? 'Yes' : 'No')" -Color ($restartNeeded ? 'Yellow' : 'Green')
Write-RepairLog "Log file location: $logFile" -Color White

if ($restartNeeded) {
    Write-RepairLog "`nSystem restart is required to complete repairs." -Color Yellow
    $restart = Read-Host "Would you like to restart your computer now? (Y/N)"
    if ($restart -eq "Y" -or $restart -eq "y") {
        Write-RepairLog "Initiating system restart..." -Color Yellow
        Start-Sleep -Seconds 3
        Restart-Computer -Force
    } else {
        Write-RepairLog "Please restart your computer at your earliest convenience to complete repairs." -Color Yellow
    }
} else {
    Write-RepairLog "`nAll operations completed successfully. No restart required." -Color Green
}

Write-RepairLog "`n=== Windows Repair Script Completed ===" -Color Cyan