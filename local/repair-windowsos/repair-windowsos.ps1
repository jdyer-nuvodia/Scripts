# Windows OS Repair Script
# Version: 3.1
# Author: Original by jdyer-nuvodia, optimized with GitHub Copilot
# Last Updated: 2025-02-05 21:46:40
# Description: Performs comprehensive Windows system repairs and health checks using PowerShell cmdlets
# Requires: PowerShell 5.1 or later, Windows 10/Server 2016 or later

#Requires -RunAsAdministrator
#Requires -Version 5.1

# Ensure we stop on errors immediately
Set-StrictMode -Version Latest
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

# Function to run repair commands and handle errors
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
        Write-RepairLog "Stack Trace: $($_.ScriptStackTrace)" -Color Red
        Write-RepairLog $ErrorMessage -Color Red
        throw  # Re-throw the error to stop script execution
    }
}

# Display initial system information
Write-RepairLog "=== Windows Repair Script Started ===" -Color Cyan
Write-RepairLog "System Information:" -Color Cyan
Write-RepairLog "Windows Version: $([System.Environment]::OSVersion.Version)" -Color White
Write-RepairLog "PowerShell Version: $($PSVersionTable.PSVersion)" -Color White
Write-RepairLog "Computer Name: $env:COMPUTERNAME" -Color White
Write-RepairLog "Log File Location: $logFile" -Color White
Write-RepairLog "----------------------------------------" -Color Cyan

# Step 1: Windows Image Health Check
Write-RepairLog "`nStep 1/3: Windows Image Health Check" -Color Cyan
Write-RepairLog "Checking Windows image health..." -Color Yellow

$imageCheck = Invoke-RepairCommand -CommandName "Windows Image Health Check" -ScriptBlock {
    # Using DISM PowerShell module commands
    $componentState = Get-WindowsOptionalFeature -Online | 
        Where-Object { $_.State -eq 'Disabled' -or $_.State -eq 'EnablePending' -or $_.State -eq 'DisablePending' }
    
    if ($null -eq $componentState) {
        return @{ Success = $true; NeedsRepair = $false }
    }
    return @{ Success = $true; NeedsRepair = $true }
} -SuccessMessage "Windows image health check completed." -ErrorMessage "Windows image health check failed."

if ($imageCheck.NeedsRepair) {
    Write-RepairLog "Image corruption detected. Initiating repair..." -Color Yellow
    $imageRepair = Invoke-RepairCommand -CommandName "Windows Image Repair" -ScriptBlock {
        # Using DISM PowerShell module commands
        $repair = Repair-WindowsImage -Online -RestoreHealth -NoRestart
        if ($repair.ImageHealthState -eq "Healthy") {
            $script:repairsMade = $true
            return $true
        }
        throw "Image repair failed to restore health"
    } -SuccessMessage "Windows image repair completed successfully." -ErrorMessage "Windows image repair encountered issues."
}

# Step 2: System File Check
Write-RepairLog "`nStep 2/3: System File Check" -Color Cyan
$sfcResult = Invoke-RepairCommand -CommandName "System File Check" -ScriptBlock {
    $process = Start-Process -FilePath "sfc.exe" -ArgumentList "/scannow" -Wait -PassThru -NoNewWindow
    if ($process.ExitCode -ne 0) {
        throw "SFC returned error code: $($process.ExitCode)"
    }
    $script:repairsMade = $true
    return @{ Success = $true; RepairsNeeded = $true }
} -SuccessMessage "System File Check completed." -ErrorMessage "System File Check encountered issues."

# Step 3: Volume Health Check
Write-RepairLog "`nStep 3/3: Volume Health Check" -Color Cyan
$systemDrive = $env:SystemDrive.TrimEnd(':')

$volumeCheck = Invoke-RepairCommand -CommandName "Volume Health Check" -ScriptBlock {
    # Get volume health details
    $volume = Get-Volume -DriveLetter $systemDrive -ErrorAction Stop
    
    # Check for basic volume health indicators
    if ($volume.HealthStatus -eq "Healthy" -and $volume.OperationalStatus -eq "OK") {
        return @{ Success = $true; Problems = $false }
    }
    return @{ Success = $true; Problems = $true }
} -SuccessMessage "Volume health check completed." -ErrorMessage "Volume health check failed."

if ($volumeCheck.Problems) {
    Write-RepairLog "Volume issues detected. Initiating repair..." -Color Yellow
    $volumeRepair = Invoke-RepairCommand -CommandName "Volume Repair" -ScriptBlock {
        $result = Repair-Volume -DriveLetter $systemDrive -Scan -ErrorAction Stop
        if (-not $result) {
            throw "Volume repair scan failed"
        }
        $script:restartNeeded = $true
        $script:repairsMade = $true
        return $true
    } -SuccessMessage "Volume repair scheduled for next restart." -ErrorMessage "Failed to schedule volume repair."
}

# Final Summary
Write-RepairLog "`n=== Repair Summary ===" -Color Cyan
Write-RepairLog "Repairs performed: $($repairsMade ? 'Yes' : 'No')" -Color ($repairsMade ? 'Yellow' : 'Green')
Write-RepairLog "Restart required: $($restartNeeded ? 'Yes' : 'No')" -Color ($restartNeeded ? 'Yellow' : 'Green')
Write-RepairLog "Log file location: $logFile" -Color White

# Handle restart if needed
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