<#
.SYNOPSIS
    Windows OS Repair and Maintenance Script
.DESCRIPTION
    Performs various Windows OS repairs and maintenance tasks in the correct order:
    1. DISM - Repairs Windows component store
    2. SFC - Repairs system files using the repaired component store
    3. Windows Update cache cleanup
.NOTES
    Created: 2025-02-06
    Author: Updated by jdyer-nuvodia
#>

# Script Variables
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$logFile = Join-Path $scriptPath "WindowsRepair_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$global:repairsMade = $false
$global:restartNeeded = $false

# Function to write formatted log entries
function Write-RepairLog {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [string]$Color = "White",
        
        [Parameter(Mandatory=$false)]
        [switch]$NoNewLine
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] $Message"
    
    # Write to console with color
    if ($NoNewLine) {
        Write-Host $logEntry -ForegroundColor $Color -NoNewline
    } else {
        Write-Host $logEntry -ForegroundColor $Color
    }
    
    # Write to log file
    Add-Content -Path $logFile -Value $logEntry
}

# Function to check disk health
function Test-DiskHealth {
    Write-RepairLog "Checking disk health..." -Color Cyan
    
    # Get system drive letter
    $systemDrive = $env:SystemDrive
    
    # Run chkdsk in read-only mode first
    Write-RepairLog "Running initial disk check on $systemDrive..." -Color Cyan
    $chkdsk = Start-Process "chkdsk.exe" -ArgumentList "$systemDrive /scan" -Wait -PassThru -WindowStyle Hidden
    
    if ($chkdsk.ExitCode -ne 0) {
        Write-RepairLog "Disk errors detected. Scheduling full chkdsk on next restart..." -Color Yellow
        # Schedule full chkdsk with fix on next reboot
        $fullChk = Start-Process "chkdsk.exe" -ArgumentList "$systemDrive /f /r" -Wait -PassThru -WindowStyle Hidden
        $global:restartNeeded = $true
        $global:repairsMade = $true
    } else {
        Write-RepairLog "No disk errors detected." -Color Green
    }
}

# Function to repair Windows image
function Repair-WindowsImage {
    Write-RepairLog "Scanning Windows image for corruption..." -Color Cyan
    $dism = Start-Process "DISM.exe" -ArgumentList "/Online /Cleanup-Image /ScanHealth" -Wait -PassThru
    if ($dism.ExitCode -eq 0) {
        Write-RepairLog "DISM scan completed successfully." -Color Green
        
        Write-RepairLog "Attempting to repair Windows image..." -Color Cyan
        $dismRepair = Start-Process "DISM.exe" -ArgumentList "/Online /Cleanup-Image /RestoreHealth" -Wait -PassThru
        if ($dismRepair.ExitCode -eq 0) {
            Write-RepairLog "Windows image repair completed successfully." -Color Green
        } else {
            Write-RepairLog "Windows image repair encountered issues." -Color Yellow
            $global:repairsMade = $true
            $global:restartNeeded = $true
        }
    } else {
        Write-RepairLog "DISM scan encountered issues." -Color Yellow
    }
}

# Function to check system file integrity
function Test-SystemFileIntegrity {
    Write-RepairLog "Checking system file integrity..." -Color Cyan
    $sfc = Start-Process "sfc.exe" -ArgumentList "/scannow" -Wait -PassThru
    if ($sfc.ExitCode -eq 0) {
        Write-RepairLog "System File Checker completed successfully." -Color Green
    } else {
        Write-RepairLog "System File Checker encountered issues." -Color Yellow
        $global:repairsMade = $true
        $global:restartNeeded = $true
    }
}

# Function to clear Windows Update cache
function Clear-WindowsUpdateCache {
    Write-RepairLog "Clearing Windows Update cache..." -Color Cyan
    
    $services = @("wuauserv", "cryptSvc", "bits", "msiserver")
    
    # Stop services
    foreach ($service in $services) {
        Stop-Service -Name $service -Force -ErrorAction SilentlyContinue
        Write-RepairLog "Stopped service: $service" -Color Gray
    }
    
    # Clear Windows Update cache
    Remove-Item "$env:SystemRoot\SoftwareDistribution\*" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item "$env:SystemRoot\System32\catroot2\*" -Recurse -Force -ErrorAction SilentlyContinue
    
    # Start services
    foreach ($service in $services) {
        Start-Service -Name $service -ErrorAction SilentlyContinue
        Write-RepairLog "Started service: $service" -Color Gray
    }
    
    Write-RepairLog "Windows Update cache cleared." -Color Green
    $global:repairsMade = $true
}

# Main script execution
try {
    Write-RepairLog "=== Windows OS Repair Script ===" -Color Cyan
    Write-RepairLog "Started by: $env:USERNAME" -Color Cyan
    Write-RepairLog "Start Time (UTC): $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Color Cyan
    Write-RepairLog "----------------------------------------" -Color Cyan
    
    # Verify running as administrator
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "This script requires administrator privileges."
    }
    
    # Execute repair functions in correct order
    Test-DiskHealth         # Run disk check first
    Repair-WindowsImage     # Run DISM second to repair component store
    Test-SystemFileIntegrity # Run SFC last to repair system files
    Clear-WindowsUpdateCache
    
    # Report results
    Write-RepairLog "----------------------------------------" -Color Cyan
    
    $repairsStatus = if ($global:repairsMade) { 'Yes' } else { 'No' }
    $repairsColor = if ($global:repairsMade) { 'Yellow' } else { 'Green' }
    Write-RepairLog "Repairs performed: $repairsStatus" -Color $repairsColor
    
    $restartStatus = if ($global:restartNeeded) { 'Yes' } else { 'No' }
    $restartColor = if ($global:restartNeeded) { 'Yellow' } else { 'Green' }
    Write-RepairLog "Restart required: $restartStatus" -Color $restartColor
    
    Write-RepairLog "End Time (UTC): $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Color Cyan
    Write-RepairLog "Log file: $logFile" -Color Cyan
    
} catch {
    Write-RepairLog "Error: $($_.Exception.Message)" -Color Red
    exit 1
}

# Prompt for restart if needed
if ($global:restartNeeded) {
    Write-RepairLog "`nSystem restart is recommended to complete repairs." -Color Yellow
    $restart = Read-Host "Would you like to restart now? (Y/N)"
    if ($restart -eq 'Y' -or $restart -eq 'y') {
        Write-RepairLog "Initiating system restart..." -Color Yellow
        Restart-Computer -Force
    } else {
        Write-RepairLog "Please restart your computer at your earliest convenience." -Color Yellow
    }
}