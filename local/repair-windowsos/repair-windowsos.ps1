# Windows OS Repair Script

# Ensure script is running with admin privileges
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Please run this script as an administrator." -ForegroundColor Red
    Exit
}

# Function to run commands and capture output
function Run-Command {
    param (
        [string]$command,
        [string[]]$arguments
    )
    $output = & $command $arguments
    return $output
}

# Variables to track repairs
$repairsMade = $false
$restartNeeded = $false

# Step 1: DISM Health Check
Write-Host "Step 1: Verifying and repairing component store with DISM..." -ForegroundColor Cyan

# Run DISM CheckHealth command and capture the output
$dismCheckOutput = & DISM.exe /Online /Cleanup-Image /CheckHealth

# Check if the output indicates no corruption
if ($dismCheckOutput -match "No component store corruption detected") {
    Write-Host "Component store is healthy." -ForegroundColor Green
} else {
    Write-Host "Potential issues found. Running DISM RestoreHealth..." -ForegroundColor Yellow
    
    # Run DISM RestoreHealth command and capture the output
    $dismRepairOutput = & DISM.exe /Online /Cleanup-Image /RestoreHealth
    
    # Check if the restore operation completed successfully
    if ($dismRepairOutput -match "The restore operation completed successfully") {
        Write-Host "DISM repair completed successfully." -ForegroundColor Green
        $repairsMade = $true
    } else {
        Write-Host "DISM repair encountered issues. Please check logs for details." -ForegroundColor Red
    }
}


# Step 2: System File Checker
Write-Host "`nStep 2: Running System File Checker..." -ForegroundColor Cyan

# Execute SFC and capture output
$sfcOutput = & sfc.exe /scannow

# Normalize the output by joining lines
$sfcOutputString = $sfcOutput -join "`n" | Out-String | ForEach-Object { $_.Trim() }

# Check for specific messages in the output
if ($sfcOutputString -match "[Windows Resource Protection did not find any integrity violations.]") {
    Write-Host "No system file integrity violations found." -ForegroundColor Green
} elseif ($sfcOutputString -match "[Windows Resource Protection found corrupt files and successfully repaired them.]") {
    Write-Host "Corrupt system files were detected and repaired." -ForegroundColor Yellow
    $repairsMade = $true
} else {
    Write-Host "SFC encountered issues. Please review the logs for details." -ForegroundColor Red
}


# Step 3: Check Disk
Write-Host "`nStep 3: Checking disk health..." -ForegroundColor Cyan

# Get the system drive
$systemDrive = $env:SystemDrive

# Execute CHKDSK and capture output
$chkdskOutput = & chkdsk.exe "$systemDrive" /scan

# Check for specific messages in the output
if ($chkdskOutput -match "Windows has scanned the file system and found no problems") {
    Write-Host "No file system errors detected." -ForegroundColor Green
} else {
    Write-Host "File system errors found. Scheduling a full chkdsk for next restart." -ForegroundColor Yellow
    & chkdsk.exe "$systemDrive" /f /x
    $repairsMade = $true
}

# Summary and restart prompt
Write-Host "`nWindows repair process completed." -ForegroundColor Cyan

# Initialize repairsMade if not already done
if (-not $repairsMade) {
    $repairsMade = $false
}

if ($repairsMade) {
    Write-Host "Some repairs were performed on your system." -ForegroundColor Yellow
}


if ($restartNeeded) {
    Write-Host "A restart is required to complete the repair process." -ForegroundColor Yellow
    $restart = Read-Host "Would you like to restart your computer now? (Y/N)"
    if ($restart -eq "Y" -or $restart -eq "y") {
        Restart-Computer -Force
    } else {
        Write-Host "Please restart your computer as soon as possible to complete the repair process." -ForegroundColor Yellow
    }
} else {
    Write-Host "No restart is required at this time." -ForegroundColor Green
}
