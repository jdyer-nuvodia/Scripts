# =============================================================================
# Script: Repair-WindowsOS.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1.5
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
# =============================================================================

<#
.SYNOPSIS
    Windows OS Repair and Maintenance Script
.DESCRIPTION
    Performs various Windows OS repairs and maintenance tasks in the correct order:
    - Key actions:
        1. DISM - Repairs Windows component store
        2. SFC - Repairs system files using the repaired component store
        3. Windows Update cache cleanup
        4. Disk health verification

    Dependencies:
    - Windows PowerShell 5.1 or later
    - Administrative privileges
    - Windows built-in repair tools (DISM, SFC, CHKDSK)

    Security considerations:
    - Requires elevation to run as Administrator
    - Modifies system files and Windows components
    - Interacts with critical Windows services

    Performance impact:
    - CPU intensive during repairs
    - May take 30+ minutes to complete
    - Requires system restart after completion
.NOTES
    Security Level: High
    Required Permissions: Local Administrator
    Validation Requirements:
    - Verify script is running as Administrator
    - Check log file for completion status
    - Confirm system stability after repairs
#>

# Script Variables
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$logFile = Join-Path $scriptPath "$env:COMPUTERNAME`_WindowsRepair_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Global state tracking
$script:repairsMade = $false
$script:restartNeeded = $false

# Color configuration for cross-platform compatibility
$script:UseAnsiColors = $PSVersionTable.PSVersion.Major -ge 7
$script:Colors = if ($script:UseAnsiColors) {
    @{
        White = "`e[97m"
        Cyan = "`e[96m"
        Green = "`e[92m"
        Yellow = "`e[93m"
        Red = "`e[91m"
        Gray = "`e[90m"
        Reset = "`e[0m"
    }
} else {
    @{
        White = "White"
        Cyan = "Cyan"
        Green = "Green"
        Yellow = "Yellow"
        Red = "Red"
        Gray = "DarkGray"
        Reset = ""
    }
}

# Function to write colored output compatible with both PowerShell 5.1 and 7+
function Write-ColorOutput {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$Color = "White",
        [Parameter(Mandatory = $false)]
        [switch]$NoNewLine
    )

    if ($script:UseAnsiColors) {
        # PowerShell 7+ with ANSI escape codes
        $colorCode = $script:Colors[$Color]
        $resetCode = $script:Colors.Reset
        if ($NoNewLine) {
            Write-Output "${colorCode}${Message}${resetCode}" -NoNewline
        } else {
            Write-Output "${colorCode}${Message}${resetCode}"
        }
    } else {
        # PowerShell 5.1 - Change console color, write output, then reset
        $originalColor = $Host.UI.RawUI.ForegroundColor
        try {
            if ($script:Colors[$Color] -and $script:Colors[$Color] -ne "") {
                $Host.UI.RawUI.ForegroundColor = $script:Colors[$Color]
            }
            if ($NoNewLine) {
                Write-Output $Message -NoNewline
            } else {
                Write-Output $Message
            }
        } finally {
            $Host.UI.RawUI.ForegroundColor = $originalColor
        }
    }
}

# Function to write formatted log entries
function Write-RepairLog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [string]$Color = "White",

        [Parameter(Mandatory = $false)]
        [switch]$NoNewLine
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] $Message"

    # Write to console with color
    if ($NoNewLine) {
        Write-ColorOutput -Message $logEntry -Color $Color -NoNewLine
    } else {
        Write-ColorOutput -Message $logEntry -Color $Color
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
        Start-Process "fsutil" -ArgumentList "dirty query $systemDrive`:" -WindowStyle Hidden -Wait -PassThru -RedirectStandardOutput "$env:TEMP\fsutil.txt" | Out-Null
        $isDirty = Get-Content "$env:TEMP\fsutil.txt" | Select-String "dirty bit"

        if ($isDirty) {
            Write-RepairLog "Disk check is already scheduled for next restart." -Color Yellow
            $script:restartNeeded = $true
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
cmd = "cmd /c chkdsk $systemDrive`: /f /r & pause"
Shell.ShellExecute "cmd.exe", "/c " & cmd, "", "runas", 1
"@
            $vbsPath = Join-Path $env:TEMP "RunChkDsk.vbs"
            $vbsScript | Out-File -FilePath $vbsPath -Encoding ASCII

            # Run the VBScript to schedule chkdsk
            Start-Process "wscript.exe" -ArgumentList $vbsPath -Wait
            Remove-Item $vbsPath -Force

            Write-RepairLog "Full disk check has been scheduled for next system restart" -Color Yellow
            $script:restartNeeded = $true
            $script:repairsMade = $true
        } else {
            Write-RepairLog "No significant disk issues detected." -Color Green
        }
    } catch {
        Write-RepairLog "Error during disk check: $_" -Color Red
        throw
    } finally {
        Remove-Item "$env:TEMP\fsutil.txt" -Force -ErrorAction SilentlyContinue
    }
}

# Function to repair Windows image
function Repair-WindowsImage {
    Write-RepairLog "Scanning Windows image for corruption..." -Color Cyan

    # Function to run DISM with timeout
    function Invoke-DISMWithTimeout {
        param (
            [string]$Arguments,
            # Increased from 30 to 60 minutes
            [int]$TimeoutMinutes = 60
        )

        try {
            # Kill any existing DISM processes first
            Get-Process "DISM" -ErrorAction SilentlyContinue | Stop-Process -Force
            Start-Sleep -Seconds 5

            $process = Start-Process "DISM.exe" -ArgumentList $Arguments -PassThru -NoNewWindow -Wait:$false
            $timeoutSeconds = $TimeoutMinutes * 60

            # Monitor process with progress indication
            $elapsed = 0
            while (-not $process.HasExited -and $elapsed -lt $timeoutSeconds) {
                Write-RepairLog "DISM operation in progress... ($elapsed seconds elapsed)" -Color Cyan -NoNewLine
                Write-ColorOutput -Message "`r" -NoNewLine
                Start-Sleep -Seconds 10
                $elapsed += 10
            }

            if (-not $process.HasExited) {
                Write-RepairLog "`nDISM operation timed out after $TimeoutMinutes minutes. Cleaning up..." -Color Yellow
                Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 10

                # Cleanup any pending operations
                Start-Process "DISM.exe" -ArgumentList "/Online /Cleanup-Image /RevertPendingActions" -Wait -NoNewWindow
                return $false
            }

            # Process has completed, wait for it to fully exit and get the exit code
            $process.WaitForExit()
            $exitCode = $process.ExitCode
            Write-RepairLog "`nDISM operation completed with exit code: $exitCode" -Color Cyan
            return ($exitCode -eq 0)
        } catch {
            Write-RepairLog "Error running DISM: $_" -Color Red
            return $false
        }
    }

    # Clear any potential pending operations
    Write-RepairLog "Cleaning up any pending DISM operations..." -Color Cyan
    Invoke-DISMWithTimeout "/Online /Cleanup-Image /RevertPendingActions" | Out-Null

    # Try cleanup first
    Write-RepairLog "Starting component store cleanup..." -Color Cyan
    Invoke-DISMWithTimeout "/Online /Cleanup-Image /StartComponentCleanup" | Out-Null

    # Perform the scan
    Write-RepairLog "Starting DISM scan..." -Color Cyan
    $scanSuccess = Invoke-DISMWithTimeout "/Online /Cleanup-Image /ScanHealth"

    if ($scanSuccess) {
        # If scan was successful, attempt a restore health operation
        Write-RepairLog "DISM scan completed successfully." -Color Green
        Write-RepairLog "Attempting Windows image repair..." -Color Cyan
        $repairSuccess = Invoke-DISMWithTimeout "/Online /Cleanup-Image /RestoreHealth"

        if ($repairSuccess) {
            Write-RepairLog "Windows image repair completed successfully." -Color Green
        } else {
            Write-RepairLog "Initial repair failed. Attempting with limited access..." -Color Yellow
            $repairSuccess = Invoke-DISMWithTimeout "/Online /Cleanup-Image /RestoreHealth /LimitAccess"

            if ($repairSuccess) {
                Write-RepairLog "Windows image repair completed successfully with limited access." -Color Green
            } else {
                Write-RepairLog "Windows image repair encountered issues. System may require offline repair." -Color Red
                $script:repairsMade = $true
                $script:restartNeeded = $true
            }
        }
    } else {
        Write-RepairLog "DISM scan failed. Attempting repair without scan..." -Color Yellow

        # Attempt repair with increasing aggressiveness
        $repairMethods = @(
            @{ Args = "/Online /Cleanup-Image /RestoreHealth /LimitAccess"; Desc = "limited access" },
            @{ Args = "/Online /Cleanup-Image /RestoreHealth"; Desc = "online sources" },
            @{ Args = "/Online /Cleanup-Image /RestoreHealth /Source:WIM:D:\sources\install.wim:1 /LimitAccess"; Desc = "WIM source" }
        )

        $repairSuccess = $false
        foreach ($method in $repairMethods) {
            if (-not $repairSuccess) {
                Write-RepairLog "Attempting repair with $($method.Desc)..." -Color Cyan
                $repairSuccess = Start-DISMWithTimeout $method.Args
                if ($repairSuccess) {
                    Write-RepairLog "Windows image repair completed successfully with $($method.Desc)." -Color Green
                    break
                }
            }
        }

        if (-not $repairSuccess) {
            Write-RepairLog "Windows image repair encountered issues. System may require offline repair." -Color Red
            $script:repairsMade = $true
            $script:restartNeeded = $true
        }
    }
}

# Function to check system file integrity
function Test-SystemFileIntegrity {
    Write-RepairLog "Checking system file integrity..." -Color Cyan

    # Run SFC with multiple attempts
    $maxAttempts = 3
    $attempt = 1
    $success = $false

    while (-not $success -and $attempt -le $maxAttempts) {
        Write-RepairLog "SFC scan attempt $attempt of $maxAttempts..." -Color Cyan

        $sfc = Start-Process "sfc.exe" -ArgumentList "/scannow" -Wait -PassThru
        if ($sfc.ExitCode -eq 0) {
            $success = $true
            Write-RepairLog "System File Checker completed successfully." -Color Green
        } else {
            Write-RepairLog "SFC attempt $attempt failed. Exit code: $($sfc.ExitCode)" -Color Yellow
            if ($attempt -lt $maxAttempts) {
                Write-RepairLog "Waiting 30 seconds before next attempt..." -Color Yellow
                Start-Sleep -Seconds 30
            }
            $attempt++
        }
    }

    if (-not $success) {
        Write-RepairLog "System File Checker failed after $maxAttempts attempts." -Color Red
        $script:repairsMade = $true
        $script:restartNeeded = $true
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
    $script:repairsMade = $true
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
    # Run disk check first
    Test-DiskHealth
    # Run DISM second to repair component store
    Repair-WindowsImage
    # Run SFC last to repair system files
    Test-SystemFileIntegrity
    Clear-WindowsUpdateCache

    # Report results
    Write-RepairLog "----------------------------------------" -Color Cyan

    $repairsStatus = if ($script:repairsMade) { 'Yes' } else { 'No' }
    $repairsColor = if ($script:repairsMade) { 'Yellow' } else { 'Green' }
    Write-RepairLog "Repairs performed: $repairsStatus" -Color $repairsColor

    $restartStatus = if ($script:restartNeeded) { 'Yes' } else { 'No' }
    $restartColor = if ($script:restartNeeded) { 'Yellow' } else { 'Green' }
    Write-RepairLog "Restart required: $restartStatus" -Color $restartColor

    Write-RepairLog "End Time (UTC): $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Color Cyan
    Write-RepairLog "Log file: $logFile" -Color Cyan

} catch {
    Write-RepairLog "Error: $($_.Exception.Message)" -Color Red
    exit 1
}

# Prompt for restart if needed
if ($script:restartNeeded) {
    Write-RepairLog "`nSystem restart is recommended to complete repairs." -Color Yellow
    $restart = Read-Host "Would you like to restart now? (Y/N)"
    if ($restart -eq 'Y' -or $restart -eq 'y') {
        Write-RepairLog "Initiating system restart..." -Color Yellow
        Restart-Computer -Force
    } else {
        Write-RepairLog "Please restart your computer at your earliest convenience." -Color Yellow
    }
}
