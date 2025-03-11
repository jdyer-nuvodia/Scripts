# =============================================================================
# Script: Clear-SystemStorage.ps1
# Created: 2025-02-27 18:55:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-11 20:32:00 UTC
# Updated By: jdyer-nuvodia
# Version: 5.1.3
# Additional Info: Removed all remaining cleanmgr references and associated monitoring code
# =============================================================================

<#
.SYNOPSIS
    Cleans up system storage by removing temporary files and managing shadow copies.
.DESCRIPTION
    This script performs system storage cleanup operations to free up disk space by:
     - Running Windows Disk Cleanup utility silently
     - Managing and reducing Volume Shadow Copies
     - Displaying drive space information before and after cleanup
     - Always elevates to SYSTEM context for thorough cleanup
.EXAMPLE
    .\Clear-SystemStorage.ps1
    Runs the script with default settings, elevating to SYSTEM context.
.NOTES
    Security Level: Medium
    Required Permissions: Administrative access for full functionality
    Validation Requirements: Verify disk space is freed after execution
#>

# ----- Global Variable Initialization -----
# Set debug preference to continue for detailed output
$DebugPreference = "Continue"
$VerbosePreference = "Continue"

$scriptPath = $MyInvocation.PSCommandPath
if (-not $scriptPath) {
    # Fallback for direct console execution
    $scriptPath = $MyInvocation.MyCommand.Definition
    if (-not $scriptPath) {
        $scriptPath = Join-Path -Path (Get-Location) -ChildPath "Clear-SystemStorage.ps1"
    }
}
$script:OriginalScriptDirectory = Split-Path -Path $scriptPath -Parent
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$computerName = $env:COMPUTERNAME
$script:LogFile = Join-Path -Path $script:OriginalScriptDirectory -ChildPath "ClearSystemStorage_${computerName}_$timestamp.log"

# ----- Function: Write-Log -----
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter()]
        [ValidateSet('Info','Warning','Error','Debug','Verbose')]
        [string]$Level = 'Info'
    )
    
    $timestampMsg = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestampMsg] [$Level] $Message"
    
    # Always write to console immediately with color coding for clear visibility.
    switch ($Level) {
        'Info' { Write-Host "$timestampMsg - INFO: $Message" -ForegroundColor White }
        'Warning' { Write-Host "$timestampMsg - WARNING: $Message" -ForegroundColor Yellow }
        'Error' { Write-Host "$timestampMsg - ERROR: $Message" -ForegroundColor Red }
        'Debug' { Write-Host "$timestampMsg - DEBUG: $Message" -ForegroundColor Magenta }
        'Verbose' { Write-Host "$timestampMsg - VERBOSE: $Message" -ForegroundColor Cyan }
    }
    
    # Always append the log entry to the log file.
    Add-Content -Path $script:LogFile -Value $logEntry
}

# Log script start with header
Write-Log "===== SCRIPT EXECUTION STARTED =====" -Level Info
Write-Log "Script version: 5.1.3" -Level Info
Write-Log "Computer Name: $computerName" -Level Info
Write-Log "Log file: $script:LogFile" -Level Info
Write-Log "Running as user: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)" -Level Info

# ----- Function: Test-RunningAsSystem -----
function Test-RunningAsSystem {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $isSystem = $currentUser.User.Value -eq "S-1-5-18"
    Write-Log -Message "Checking if running as SYSTEM: $isSystem" -Level Verbose
    return $isSystem
}

# ----- Function: Start-SystemContext -----
function Start-SystemContext {
    param(
        [string]$ScriptPath = $MyInvocation.PSCommandPath,
        [int]$TimeoutSeconds = 900
    )
    
    if (Test-RunningAsSystem) {
        Write-Log "Already running as SYSTEM account." -Level Info
        return $true
    }
    
    Write-Log "Elevating to SYSTEM context..." -Level Info
    Write-Host "`nNOTE: Script will continue execution as SYSTEM." -ForegroundColor Cyan
    
    try {
        $jobName = "SystemContextJob_$([Guid]::NewGuid())"
        Write-Log "Created job with name: $jobName" -Level Debug
        
        # Use more reliable temp path creation
        $systemAccessibleTemp = [System.IO.Path]::Combine($env:SystemRoot, "Temp")
        $scriptFileName = Split-Path -Leaf $ScriptPath
        $systemAccessibleScriptPath = [System.IO.Path]::Combine($systemAccessibleTemp, $scriptFileName)
        
        # Initialize monitoring variables
        $maxRetries = 3
        $retryDelay = 2
        $maxMonitoringTime = 300 # 5 minutes
        $statusCheckInterval = 500 # milliseconds
        $completed = $false
        $iterationCount = 0

        # Ensure cleanup of existing files
        $filesToClean = @($systemAccessibleScriptPath, $executorScript, $logFile)
        foreach ($file in $filesToClean) {
            if ([System.IO.File]::Exists($file)) {
                try {
                    [System.IO.File]::Delete($file)
                    Write-Log "Cleaned up existing file: $file" -Level Debug
                }
                catch {
                    Write-Log "Warning: Could not delete existing file $file : $($_.Exception.Message)" -Level Warning
                }
            }
        }

        # Copy script to system-accessible location
        Copy-Item -Path $ScriptPath -Destination $systemAccessibleScriptPath -Force
        if (-not [System.IO.File]::Exists($systemAccessibleScriptPath)) {
            Write-Log "Failed to copy script to system location!" -Level Error
            return $false
        }

        # Create simple executor script
        $executorScript = Join-Path -Path $systemAccessibleTemp -ChildPath "$jobName-executor.ps1"
        $logFile = Join-Path $systemAccessibleTemp "$jobName.log"

        $executorContent = @"
# System context executor for Clear-SystemStorage
`$ErrorActionPreference = 'Stop'
`$VerbosePreference = 'Continue'
Start-Transcript -Path '$logFile' -Force

try {
    & '$systemAccessibleScriptPath' -Verbose *>> '$logFile'
}
catch {
    Write-Error "`$(`$_.Exception.Message)"
    exit 1
}
"@
        Set-Content -Path $executorScript -Value $executorContent -Encoding UTF8 -Force

        # Schedule and start the task
        $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$executorScript`""
        $principal = New-ScheduledTaskPrincipal -UserId "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
        $task = New-ScheduledTask -Action $action -Principal $principal -Settings $settings
        
        Register-ScheduledTask -TaskName $jobName -InputObject $task | Out-Null
        Start-ScheduledTask -TaskName $jobName

        Write-Log "Started scheduled task as SYSTEM" -Level Info
        Write-Host "Task running as SYSTEM. Monitoring progress..." -ForegroundColor Cyan

        return $true
    }
    catch {
        Write-Log "Error during system context elevation: $($_.Exception.Message)" -Level Error
        return $false
    }
}

# ----- Function: Get-SystemDiskSpace -----
function Get-SystemDiskSpace {
    Write-Log "Retrieving drive space information..." -Level Debug
    try {
        $systemDrive = Get-Volume -DriveLetter $env:SystemDrive[0]
        Write-Log "Successfully retrieved drive information for $($env:SystemDrive)" -Level Debug
        return $systemDrive
    } catch {
        Write-Log "Error retrieving drive information: $($_.Exception.Message)" -Level Error
        return $null
    }
}

# ----- Function: Show-DriveInfo -----
function Show-DriveInfo {
    param (
        [Parameter(Mandatory=$true)]
        [object]$Volume,
        
        [Parameter()]
        [string]$Label = "Drive Volume Details"
    )
    
    Write-Log "Displaying drive information for $($Volume.DriveLetter): label=$Label" -Level Debug
    
    # Calculate used space and percentage
    $usedSpace = $Volume.Size - $Volume.SizeRemaining
    $usedPercentage = [math]::Round(($usedSpace / $Volume.Size) * 100, 1)
    $freePercentage = [math]::Round(($Volume.SizeRemaining / $Volume.Size) * 100, 1)
    
    Write-Host "`n${Label}:" -ForegroundColor Green
    Write-Host "------------------------" -ForegroundColor Green
    Write-Host "Drive Letter: $($Volume.DriveLetter)" -ForegroundColor Cyan
    Write-Host "Drive Label: $($Volume.FileSystemLabel)" -ForegroundColor Cyan
    Write-Host "File System: $($Volume.FileSystem)" -ForegroundColor Cyan
    Write-Host "Drive Type: $($Volume.DriveType)" -ForegroundColor Cyan
    Write-Host "Total Size: $([math]::Round($Volume.Size/1GB, 2)) GB" -ForegroundColor Cyan
    Write-Host "Used Space: $([math]::Round($usedSpace/1GB, 2)) GB ($usedPercentage%)" -ForegroundColor $(if ($usedPercentage -gt 90) { "Yellow" } else { "Cyan" })
    Write-Host "Free Space: $([math]::Round($Volume.SizeRemaining/1GB, 2)) GB ($freePercentage%)" -ForegroundColor $(if ($freePercentage -lt 10) { "Red" } elseif ($freePercentage -lt 20) { "Yellow" } else { "Green" })
    Write-Host "Health Status: $($Volume.HealthStatus)" -ForegroundColor Cyan
    
    Write-Log "Drive $($Volume.DriveLetter) has $([math]::Round($Volume.SizeRemaining/1GB, 2)) GB free of $([math]::Round($Volume.Size/1GB, 2)) GB total" -Level Info
}

# ----- Function: Clear-SystemStorageOptimized -----
function Clear-SystemStorageOptimized {
    [CmdletBinding()]
    param()
    
    Write-Log "Initializing optimized system storage cleanup..." -Level Info
    
    try {
        # Initialize only required VSS object
        $vssService = Get-WmiObject -Class Win32_ShadowCopy

        # Enhanced cleanup paths using .NET methods
        $cleanupPaths = @(
            [System.Environment]::GetFolderPath('Windows') + '\Temp',
            [System.Environment]::GetFolderPath('LocalApplicationData') + '\Temp',
            [System.Environment]::GetFolderPath('InternetCache')
        )

        $totalBytesFreed = 0
        $errors = @()

        # Process VSS copies first for maximum space recovery
        Write-Log "Managing Volume Shadow Copies..." -Level Info
        try {
            # Get all shadow copies
            $shadowCopies = $vssService | Where-Object { $null -ne $_.ID }
            
            if ($null -ne ($shadowCopies | Where-Object { $null -ne $_.ID })) {
                $shadowCopies | ForEach-Object {
                    try {
                        $_.Delete()
                        Write-Log "Removed shadow copy: $($_.ID)" -Level Debug
                    }
                    catch {
                        Write-Log "Warning: Could not remove shadow copy $($_.ID): $($_.Exception.Message)" -Level Warning
                        $errors += "VSS:$($_.ID)"
                    }
                }
            }
            else {
                Write-Log "No shadow copies found to clean" -Level Debug
            }
        }
        catch {
            Write-Log "Warning: VSS management error: $($_.Exception.Message)" -Level Warning
            $errors += "VSS:General"
        }

        # Enhanced temp file cleanup using .NET methods
        foreach ($path in $cleanupPaths) {
            if ([System.IO.Directory]::Exists($path)) {
                Write-Log "Processing directory: $path" -Level Info
                
                try {
                    $files = [System.IO.Directory]::GetFiles($path, "*", [System.IO.SearchOption]::AllDirectories)
                    $directories = [System.IO.Directory]::GetDirectories($path, "*", [System.IO.SearchOption]::AllDirectories)
                    
                    # Process files first
                    foreach ($file in $files) {
                        try {
                            $fileInfo = [System.IO.FileInfo]::new($file)
                            $size = $fileInfo.Length
                            
                            # Use robust file deletion with retries
                            $retryCount = 0
                            $maxRetries = 3
                            $deleted = $false
                            
                            while (-not $deleted -and $retryCount -lt $maxRetries) {
                                try {
                                    [System.IO.File]::Delete($file)
                                    $totalBytesFreed += $size
                                    $deleted = $true
                                }
                                catch [System.IO.IOException] {
                                    $retryCount++
                                    if ($retryCount -lt $maxRetries) {
                                        Start-Sleep -Milliseconds 500
                                        [System.GC]::Collect()
                                        [System.GC]::WaitForPendingFinalizers()
                                    }
                                }
                            }
                            
                            if (-not $deleted) {
                                $errors += "File:$file"
                            }
                        }
                        catch {
                            Write-Log "Warning: Could not process file $file : $($_.Exception.Message)" -Level Warning
                            $errors += "File:$file"
                            continue
                        }
                    }
                    
                    # Process directories in reverse order (deepest first)
                    [array]::Reverse($directories)
                    foreach ($dir in $directories) {
                        try {
                            if ([System.IO.Directory]::GetFileSystemEntries($dir).Count -eq 0) {
                                [System.IO.Directory]::Delete($dir, $false)
                            }
                        }
                        catch {
                            Write-Log "Warning: Could not remove empty directory $dir : $($_.Exception.Message)" -Level Warning
                            $errors += "Dir:$dir"
                            continue
                        }
                    }
                }
                catch {
                    Write-Log "Error processing path $path : $($_.Exception.Message)" -Level Error
                    $errors += "Path:$path"
                    continue
                }
            }
        }

        # Report results
        $freedSpaceGB = [math]::Round($totalBytesFreed / 1GB, 2)
        Write-Log "Cleanup completed. Freed $freedSpaceGB GB of space." -Level Info
        
        if ($errors.Count -gt 0) {
            Write-Log "Cleanup completed with $($errors.Count) non-critical errors" -Level Warning
        }
        else {
            Write-Log "Cleanup completed successfully with no errors" -Level Info
        }

        return @{
            BytesFreed = $totalBytesFreed
            Errors = $errors
        }
    }
    catch {
        Write-Log "Critical error during cleanup: $($_.Exception.Message)" -Level Error
        throw
    }
}

# --------------------------------------------------
# Main Script Execution (after elevation if needed)
# --------------------------------------------------

# If not running as SYSTEM, initiate SYSTEM context.
if (-not (Test-RunningAsSystem)) {
    Start-SystemContext -ScriptPath $scriptPath
    # Suppress boolean output
    exit
}

Write-Log "Starting main cleanup process as SYSTEM account" -Level Info

# ----- Display Starting Drive Information -----
Write-Log "Collecting initial system drive information" -Level Info
$initialDrive = Get-SystemDiskSpace
if ($initialDrive) {
    Show-DriveInfo -Volume $initialDrive -Label "INITIAL Drive Volume Details"
} else {
    Write-Log "Failed to get initial drive information" -Level Warning
}

# ----- Display Shadow Copy Count -----
Write-Log "Checking initial shadow copy status..." -Level Info
$shadowOutputInitial = vssadmin list shadows 2>&1
Write-Log "Shadow copy command executed. Processing output..." -Level Debug

$totalShadowCopiesInitial = ($shadowOutputInitial | Select-String "Shadow Copy ID").Count
Write-Log "Initial Shadow Copy Count: $totalShadowCopiesInitial" -Level Info
Write-Host "Shadow Copy Count: ${totalShadowCopiesInitial}" -ForegroundColor Yellow

# Extract and display more details about shadow copies
if ($totalShadowCopiesInitial -gt 0) {
    Write-Host "`n===== Initial Shadow Copy Details =====" -ForegroundColor Yellow
    $shadowDetails = $shadowOutputInitial | Select-String -Pattern "Shadow Copy Volume:" -Context 0,0
    foreach ($detail in $shadowDetails) {
        Write-Host $detail -ForegroundColor DarkYellow
    }
}

# ----- Function: Invoke-SystemCleanup -----
function Invoke-SystemCleanup {
    [CmdletBinding()]
    param()
    
    Write-Log "Starting native PowerShell system cleanup..." -Level Info
    $cleanupStats = @{
        TotalBytesRemoved = 0
        ItemsRemoved = 0
        Errors = @()
    }

    # Define cleanup paths with descriptions
    $cleanupPaths = @{
        "Windows Temp" = @{
            Path = "$env:SystemRoot\Temp"
            Pattern = "*"
            Recursive = $true
        }
        "User Temp" = @{
            Path = [System.IO.Path]::GetTempPath()
            Pattern = "*"
            Recursive = $true
        }
        "Windows SoftwareDistribution" = @{
            Path = "$env:SystemRoot\SoftwareDistribution\Download"
            Pattern = "*"
            Recursive = $true
        }
        "Delivery Optimization" = @{
            Path = "$env:SystemRoot\DeliveryOptimization"
            Pattern = "*"
            Recursive = $true
        }
        "Windows Error Reports" = @{
            Path = "$env:ProgramData\Microsoft\Windows\WER"
            Pattern = "*"
            Recursive = $true
        }
        "Chrome Cache" = @{
            Path = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache"
            Pattern = "*"
            Recursive = $true
        }
        "Edge Cache" = @{
            Path = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache"
            Pattern = "*"
            Recursive = $true
        }
        "Firefox Cache" = @{
            Path = "$env:LOCALAPPDATA\Mozilla\Firefox\Profiles\*.default*\cache2"
            Pattern = "*"
            Recursive = $true
        }
        "Old Downloads" = @{
            Path = [Environment]::GetFolderPath('UserProfile') + '\Downloads'
            Pattern = "*"
            Recursive = $true
            AgeFilter = 180 # Days
        }
    }

    # Function to safely remove files with age filter
    function Remove-PathContents {
        param (
            [string]$Path,
            [string]$Pattern,
            [bool]$Recursive,
            [string]$Description,
            [int]$AgeFilter = 0
        )

        if (![System.IO.Directory]::Exists($Path)) {
            Write-Log "Path not found: $Path" -Level Debug
            return
        }

        Write-Host "`nCleaning $Description..." -ForegroundColor Cyan
        Write-Log "Processing cleanup location: $Description ($Path)" -Level Info
        
        if ($AgeFilter -gt 0) {
            Write-Host "Age filter: Removing files older than $AgeFilter days" -ForegroundColor DarkGray
        }

        try {
            # Get all files
            $files = [System.IO.Directory]::GetFiles($Path, $Pattern, 
                $(if ($Recursive) {[System.IO.SearchOption]::AllDirectories} else {[System.IO.SearchOption]::TopDirectoryOnly}))
            
            $filteredFiles = $files | Where-Object {
                $fileInfo = [System.IO.FileInfo]::new($_)
                if ($AgeFilter -gt 0) {
                    return ($fileInfo.LastWriteTime -lt (Get-Date).AddDays(-$AgeFilter))
                }
                return $true
            }
            
            $totalFiles = $filteredFiles.Count
            $processed = 0
            $bytesRemoved = 0

            foreach ($file in $filteredFiles) {
                try {
                    $fileInfo = [System.IO.FileInfo]::new($file)
                    $size = $fileInfo.Length
                    
                    # Skip if file is in use
                    if ((Test-IsFileInUse -Path $file)) {
                        continue
                    }

                    [System.IO.File]::Delete($file)
                    $bytesRemoved += $size
                    $processed++

                    # Update progress
                    $percentComplete = [math]::Round(($processed / $totalFiles) * 100)
                    Write-Progress -Activity "Cleaning $Description" -Status "$percentComplete% Complete" `
                        -PercentComplete $percentComplete
                }
                catch {
                    $script:cleanupStats.Errors += "$Description - $($_.Exception.Message)"
                    continue
                }
            }

            # Clean empty directories if recursive
            if ($Recursive) {
                $dirs = [System.IO.Directory]::GetDirectories($Path, "*", [System.IO.SearchOption]::AllDirectories)
                [Array]::Reverse($dirs) # Process deepest dirs first
                
                foreach ($dir in $dirs) {
                    try {
                        if ([System.IO.Directory]::GetFileSystemEntries($dir).Count -eq 0) {
                            [System.IO.Directory]::Delete($dir, $false)
                        }
                    }
                    catch {
                        $script:cleanupStats.Errors += "Directory: $dir - $($_.Exception.Message)"
                    }
                }
            }

            $script:cleanupStats.TotalBytesRemoved += $bytesRemoved
            $script:cleanupStats.ItemsRemoved += $processed

            Write-Progress -Activity "Cleaning $Description" -Completed
            Write-Host "Removed $([math]::Round($bytesRemoved/1MB, 2)) MB from $Description" -ForegroundColor Green
            Write-Log "Completed cleanup of $Description. Removed $([math]::Round($bytesRemoved/1MB, 2)) MB" -Level Info
        }
        catch {
            Write-Log "Error processing $Description : $($_.Exception.Message)" -Level Error
            $script:cleanupStats.Errors += "$Description - $($_.Exception.Message)"
        }
    }

    # Function to check if file is in use
    function Test-IsFileInUse {
        param([string]$Path)
        try {
            $fileStream = [System.IO.File]::Open($Path, 'Open', 'Read', 'None')
            $fileStream.Close()
            $fileStream.Dispose()
            return $false
        }
        catch {
            return $true
        }
    }

    # Clean Windows Component Store using DISM PowerShell module
    Write-Host "`nCleaning Windows Component Store..." -ForegroundColor Cyan
    try {
        $null = Repair-WindowsImage -Online -StartComponentCleanup -NoRestart
        Write-Host "Successfully cleaned Windows Component Store" -ForegroundColor Green
    }
    catch {
        Write-Log "Error cleaning Component Store: $($_.Exception.Message)" -Level Error
        $cleanupStats.Errors += "Component Store - $($_.Exception.Message)"
    }

    # Process each cleanup path
    foreach ($cleanup in $cleanupPaths.GetEnumerator()) {
        Remove-PathContents -Path $cleanup.Value.Path -Pattern $cleanup.Value.Pattern `
            -Recursive $cleanup.Value.Recursive -Description $cleanup.Name `
            -AgeFilter $($cleanup.Value.AgeFilter ?? 0)
    }

    # Empty Recycle Bin
    Write-Host "`nEmptying Recycle Bin..." -ForegroundColor Cyan
    try {
        $shell = New-Object -ComObject Shell.Application
        $shell.Namespace(0x0A).Items() | ForEach-Object {
            $cleanupStats.TotalBytesRemoved += $_.Size
            $cleanupStats.ItemsRemoved++
        }
        Clear-RecycleBin -Force -ErrorAction Stop
        Write-Host "Successfully emptied Recycle Bin" -ForegroundColor Green
        Write-Log "Recycle Bin emptied successfully" -Level Info
    }
    catch {
        Write-Log "Error emptying Recycle Bin: $($_.Exception.Message)" -Level Error
        $cleanupStats.Errors += "Recycle Bin - $($_.Exception.Message)"
    }

    # Return cleanup statistics
    return $cleanupStats
}

# Replace the old cleanup section with the new implementation
Write-Host "`n===== Starting System Storage Cleanup =====" -ForegroundColor Cyan
$initialDrive = Get-SystemDiskSpace
Show-DriveInfo -Volume $initialDrive -Label "Initial Drive State"

$cleanupResults = Invoke-SystemCleanup
$freedSpace = [math]::Round($cleanupResults.TotalBytesRemoved / 1GB, 2)

Write-Host "`nCleanup Summary:" -ForegroundColor Cyan
Write-Host "------------------------" -ForegroundColor Cyan
Write-Host "Total Space Freed: $freedSpace GB" -ForegroundColor Green
Write-Host "Items Removed: $($cleanupResults.ItemsRemoved)" -ForegroundColor Green

if ($cleanupResults.Errors.Count -gt 0) {
    Write-Host "`nWarnings/Errors:" -ForegroundColor Yellow
    $cleanupResults.Errors | ForEach-Object {
        Write-Host "- $_" -ForegroundColor Yellow
    }
}

$finalDrive = Get-SystemDiskSpace
Show-DriveInfo -Volume $finalDrive -Label "Final Drive State"
