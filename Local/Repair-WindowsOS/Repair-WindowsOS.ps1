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
    $systemDrive = $env:SystemDrive.TrimEnd(':')
    
    try {
        # First check if chkdsk is already scheduled
        $fsutil = Start-Process "fsutil" -ArgumentList "dirty query $systemDrive`:" -WindowStyle Hidden -Wait -PassThru -RedirectStandardOutput "$env:TEMP\fsutil.txt"
        $isDirty = Get-Content "$env:TEMP\fsutil.txt" | Select-String "dirty bit"
        
        if ($isDirty) {
            Write-RepairLog "Disk check is already scheduled for next restart." -Color Yellow
            $global:restartNeeded = $true
            return
        }

        # Run initial disk analysis
        Write-RepairLog "Running disk analysis on $systemDrive..." -Color Cyan
        $chkdskOutput = & chkdsk.exe $systemDrive`: /scan
        
        # Display chkdsk findings
        Write-RepairLog "Disk Analysis Results:" -Color White
        foreach ($line in $chkdskOutput) {
            if ($line -match "Windows has scanned the file system" -or
                $line -match "found \d+ problems" -or
                $line -match "Windows needs to repair") {
                Write-RepairLog "`t$line" -Color Yellow
            }
        }

        # If problems found, schedule full chkdsk
        if ($chkdskOutput -match "found \d+ problems" -or $chkdskOutput -match "Windows needs to repair") {
            Write-RepairLog "Disk errors detected. Scheduling full disk check..." -Color Yellow
            
            # Create temporary VBScript to handle UAC and show progress
            $vbsScript = @"
Set Shell = CreateObject("Shell.Application")
Set FSO = CreateObject("Scripting.FileSystemObject")
cmd = "cmd /c chkdsk $systemDrive`: /f /r & pause"
Shell.ShellExecute "cmd.exe", "/c " & cmd, "", "runas", 1
"@
            $vbsPath = Join-Path $env:TEMP "RunChkDsk.vbs"
            $vbsScript | Out-File -FilePath $vbsPath -Encoding ASCII
            
            # Run the VBScript to schedule chkdsk
            Start-Process "wscript.exe" -ArgumentList $vbsPath -Wait
            Remove-Item $vbsPath -Force
            
            Write-RepairLog "Full disk check has been scheduled for next system restart" -Color Yellow
            $global:restartNeeded = $true
            $global:repairsMade = $true
        } else {
            Write-RepairLog "No significant disk issues detected." -Color Green
        }
    }
    catch {
        Write-RepairLog "Error during disk check: $_" -Color Red
        throw
    }
    finally {
        Remove-Item "$env:TEMP\fsutil.txt" -Force -ErrorAction SilentlyContinue
    }
}

# Function to repair Windows image
function Repair-WindowsImage {
    Write-RepairLog "Scanning Windows image for corruption..." -Color Cyan
    
    # Function to run DISM with timeout
    function Start-DISMWithTimeout {
        param (
            [string]$Arguments,
            [int]$TimeoutMinutes = 30
        )
        
        try {
            $process = Start-Process "DISM.exe" -ArgumentList $Arguments -PassThru -NoNewWindow -Wait:$false
            $timeoutSeconds = $TimeoutMinutes * 60
            $processExited = $process.WaitForExit($timeoutSeconds * 1000)  # WaitForExit takes milliseconds
            
            if (-not $processExited) {
                Write-RepairLog "DISM operation timed out after $TimeoutMinutes minutes. Terminating process..." -Color Red
                Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
                return $false
            }
            
            return ($process.ExitCode -eq 0)
        }
        catch {
            Write-RepairLog "Error running DISM: $_" -Color Red
            return $false
        }
    }
    
    # Check DISM process is not already running
    $existingDISM = Get-Process "DISM" -ErrorAction SilentlyContinue
    if ($existingDISM) {
        Write-RepairLog "WARNING: Existing DISM process detected. Attempting to clean up..." -Color Yellow
        $existingDISM | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 5
    }
    
    # Clear any potential pending operations
    Write-RepairLog "Cleaning up any pending DISM operations..." -Color Cyan
    Start-DISMWithTimeout "/Online /Cleanup-Image /RevertPendingActions" | Out-Null
    
    # Perform the scan
    Write-RepairLog "Starting DISM scan..." -Color Cyan
    $scanSuccess = Start-DISMWithTimeout "/Online /Cleanup-Image /ScanHealth"
    
    if (-not $scanSuccess) {
        Write-RepairLog "DISM scan failed. Attempting alternative repair approach..." -Color Yellow
        # Try checking component store
        $checkStoreSuccess = Start-DISMWithTimeout "/Online /Cleanup-Image /CheckHealth"
        if (-not $checkStoreSuccess) {
            Write-RepairLog "Component store check failed. Attempting forced cleanup..." -Color Yellow
            Start-DISMWithTimeout "/Online /Cleanup-Image /StartComponentCleanup /ResetBase" | Out-Null
        }
    }
    
    # Attempt repair
    Write-RepairLog "Attempting to repair Windows image..." -Color Cyan
    $repairSuccess = Start-DISMWithTimeout "/Online /Cleanup-Image /RestoreHealth /LimitAccess"
    
    if (-not $repairSuccess) {
        Write-RepairLog "Initial repair attempt failed. Trying without /LimitAccess..." -Color Yellow
        $repairSuccess = Start-DISMWithTimeout "/Online /Cleanup-Image /RestoreHealth"
    }
    
    if ($repairSuccess) {
        Write-RepairLog "Windows image repair completed successfully." -Color Green
    } else {
        Write-RepairLog "Windows image repair encountered issues. System may require offline repair." -Color Red
        $global:repairsMade = $true
        $global:restartNeeded = $true
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