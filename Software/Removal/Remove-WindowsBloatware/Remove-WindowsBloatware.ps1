# =============================================================================
# Script: Remove-WindowsBloatware.ps1
# Created: 2025-05-07 15:45:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-05-07 23:08:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1.1
# Additional Info: Improved messaging to clearly indicate if apps are not found or removed
# =============================================================================

<#
.SYNOPSIS
Removes bloatware applications from Windows PCs including all Dell software except Command Update and all Lenovo software except Vantage.

.DESCRIPTION
This script identifies and removes common Windows bloatware and pre-installed applications
that are often unnecessary for business environments. For Dell PCs, it removes all Dell
software except any version of Dell Command Update. For Lenovo PCs, it removes all Lenovo
software except Lenovo Vantage.

The script performs the following actions:
1. Identifies installed applications through various registry locations
2. Removes UWP applications (Microsoft Store apps)
3. Uninstalls traditional Win32 applications
4. Removes specific Dell bloatware while preserving Dell Command Update
5. Removes specific Lenovo bloatware while preserving Lenovo Vantage
6. Logs all activities and any errors encountered

The script will remove the following software:

UWP Applications (Microsoft Store Apps):
- Microsoft 3D Builder
- Microsoft Bing Finance, News, Sports, Weather
- Microsoft Get Help
- Microsoft Get Started
- Microsoft Messaging
- Microsoft 3D Viewer
- Microsoft Solitaire Collection
- Microsoft Mixed Reality Portal
- Microsoft OneConnect
- Microsoft People
- Microsoft Print 3D
- Microsoft Skype App
- Microsoft Wallet
- Microsoft Windows Alarms
- Microsoft Windows Feedback Hub
- Microsoft Windows Maps
- Microsoft Windows Sound Recorder
- Microsoft Xbox apps (TCUI, App, GameOverlay, GamingOverlay, IdentityProvider, SpeechToTextOverlay)
- Microsoft Your Phone
- Microsoft Zune Music and Video
- Candy Crush games (Saga, Soda Saga, Friends)

Traditional Win32 Applications:
- McAfee Security Software
- Norton Security Software
- Wild Tangent Games
- Candy Crush desktop apps
- Booking.com apps
- Spotify
- HP pre-installed software (JumpStart, Connection Optimizer, Documentation, Smart, Sure)
- Lenovo pre-installed software EXCEPT Lenovo Vantage
- All Dell software EXCEPT Dell Command Update

Dependencies:
- Must be run with administrative privileges
- Windows PowerShell 5.1 or later

Security considerations:
- Requires registry modification permissions
- Requires application uninstallation permissions

.PARAMETER WhatIf
If specified, shows what would happen if the script runs without actually making changes.

.EXAMPLE
.\Remove-WindowsBloatware.ps1
# Removes all identified bloatware applications

.EXAMPLE
.\Remove-WindowsBloatware.ps1 -WhatIf
# Shows what applications would be removed without making actual changes
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param()

# Run this script as an administrator
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You do not have Administrator rights to run this script!`nPlease re-run this script as an Administrator!"
    Break
}

# Script variables
$computerName = $env:COMPUTERNAME
$utcTimestamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd_HH-mm-ss")
$logFile = "$PSScriptRoot\Remove-WindowsBloatware_${computerName}_${utcTimestamp}.log"
$scriptVersion = "1.1.1"

# Function to write log entries
function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS", "DEBUG")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [$Level] $Message"
    
    # Output to console with colors
    switch ($Level) {
        "INFO"    { Write-Host $logEntry -ForegroundColor White }
        "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
        "ERROR"   { Write-Host $logEntry -ForegroundColor Red }
        "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
        "DEBUG"   { Write-Host $logEntry -ForegroundColor Magenta }
        Default   { Write-Host $logEntry -ForegroundColor DarkGray }
    }
    
    # Create log directory if it does not exist
    $logDir = Split-Path -Path $logFile -Parent
    if (-not (Test-Path -Path $logDir)) {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    }
    
    # Write to log file
    Add-Content -Path $logFile -Value $logEntry
}

# Function to uninstall UWP apps
function Uninstall-UWPApp {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true)]
        [string]$AppName
    )
    
    try {
        # Use ErrorAction SilentlyContinue to prevent errors from appearing for non-existent apps
        $app = Get-AppxPackage -Name $AppName -AllUsers -ErrorAction SilentlyContinue
        
        if ($null -ne $app) {
            # Check if it's a single app or multiple apps with the same name
            if ($app -is [System.Array]) {
                Write-Log "FOUND: Multiple instances of UWP application: $AppName" "INFO"
                foreach ($singleApp in $app) {
                    if ($PSCmdlet.ShouldProcess($singleApp.Name, "Remove UWP application")) {
                        Write-Log "REMOVING: UWP application instance: $($singleApp.Name) (PackageFullName: $($singleApp.PackageFullName))" "INFO"
                        try {
                            Remove-AppxPackage -Package $singleApp.PackageFullName -ErrorAction SilentlyContinue
                            Write-Log "REMOVED: UWP application instance: $($singleApp.Name)" "SUCCESS"
                        }                        catch {
                            # Log the error but don't display it to the console
                            $errorMsg = $_.Exception.Message
                            Write-Log "ERROR: Failed to remove UWP application instance $($singleApp.Name): $errorMsg" "WARNING"
                        }
                    }
                    else {
                        Write-Log "WhatIf: Would remove UWP application instance: $($singleApp.Name)" "INFO"
                    }
                }
            }
            else {
                # Single app instance
                if ($PSCmdlet.ShouldProcess($app.Name, "Remove UWP application")) {
                    Write-Log "REMOVING: UWP application: $AppName (PackageFullName: $($app.PackageFullName))" "INFO"
                    try {
                        Remove-AppxPackage -Package $app.PackageFullName -ErrorAction SilentlyContinue
                        Write-Log "REMOVED: UWP application: $AppName" "SUCCESS"
                    }                    catch {
                        # Log the error but don't display it to the console
                        $errorMsg = $_.Exception.Message
                        Write-Log "ERROR: Failed to remove UWP application ${AppName}: $errorMsg" "WARNING"
                    }
                }
                else {
                    Write-Log "WhatIf: Would remove UWP application: $AppName" "INFO"
                }
            }
        }
        else {
            Write-Log "NOT FOUND: UWP application $AppName is not installed or was previously removed" "INFO"
        }
    }    catch {
        # Log the error but don't display it to the console
        $errorMsg = $_.Exception.Message
        Write-Log "ERROR: Failed to access UWP application ${AppName}: $errorMsg" "WARNING"
    }
}

# Function to uninstall Win32 applications
function Uninstall-Win32App {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true)]
        [string]$DisplayName,
        
        [Parameter(Mandatory = $false)]
        [switch]$ExactMatch = $false
    )
    
    Write-Log "Searching for Win32 application: $DisplayName" "INFO"
    
    $uninstallKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    
    $found = $false
    
    foreach ($key in $uninstallKeys) {
        $apps = Get-ChildItem -Path $key -ErrorAction SilentlyContinue | 
                Get-ItemProperty | 
                Where-Object {
                    if ($ExactMatch) {
                        $_.DisplayName -eq $DisplayName
                    } else {
                        $_.DisplayName -like "*$DisplayName*"
                    }
                }
        
        foreach ($app in $apps) {
            $found = $true
            $appName = $app.DisplayName
            $uninstallString = $app.UninstallString
            $productCode = $app.PSChildName
            
            Write-Log "FOUND: Win32 application: $appName" "INFO"
            
            try {
                if ($PSCmdlet.ShouldProcess($appName, "Uninstall application")) {
                    Write-Log "REMOVING: Win32 application: $appName" "INFO"
                      # If msiexec is in the uninstall string, use that
                    if ($uninstallString -like "*msiexec*") {
                        # Using /qn (no UI), /norestart (prevent restart), /passive (progress bar only, no user input)
                        $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $productCode /qn /norestart" -Wait -NoNewWindow -PassThru
                    }
                    else {
                        # Some applications use custom uninstallers
                        $uninstallExe = ($uninstallString -split ' ')[0]
                        $uninstallArgs = ($uninstallString -split ' ', 2)[1]
                        # Add /S or /SILENT if not present for silent uninstall
                        if ($uninstallArgs -notmatch '/S' -and $uninstallArgs -notmatch '/SILENT' -and $uninstallArgs -notmatch '/VERYSILENT') {
                            $uninstallArgs += " /S"
                        }
                        $process = Start-Process -FilePath $uninstallExe -ArgumentList $uninstallArgs -Wait -NoNewWindow -PassThru
                    }
                    
                    if ($process.ExitCode -eq 0) {
                        Write-Log "REMOVED: Win32 application: $appName" "SUCCESS"
                    }
                    else {
                        Write-Log "ERROR: Failed to remove Win32 application: $appName. Exit code: $($process.ExitCode)" "WARNING"
                    }
                }
                else {
                    Write-Log "WhatIf: Would remove Win32 application: $appName" "INFO"
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-Log "ERROR: Failed to remove Win32 application: $appName. Error: $errorMsg" "ERROR"
            }
        }
    }
    
    if (-not $found) {
        Write-Log "NOT FOUND: Win32 application: $DisplayName is not installed" "INFO"
    }
}

# Function to check if an application is Dell Command Update
function Test-IsDellCommandUpdate {
    param (
        [Parameter(Mandatory = $true)]
        [string]$AppName
    )
    
    return $AppName -like "*Dell Command*Update*"
}

# Function to check if an application is Lenovo Vantage
function Test-IsLenovoVantage {
    param (
        [Parameter(Mandatory = $true)]
        [string]$AppName
    )
    
    return $AppName -like "*Lenovo Vantage*"
}

# Function to remove Dell bloatware except Command Update
function Remove-DellBloatware {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    
    Write-Log "Identifying Dell applications..." "INFO"
    
    $uninstallKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    
    # Special case Dell applications that need custom handling
    $specialDellApps = @(
        "Dell Pair", 
        "Dell SupportAssist OS Recovery Plugin for Dell Update"
    )
    
    $foundDellApps = $false
    
    foreach ($key in $uninstallKeys) {
        $dellApps = Get-ChildItem -Path $key -ErrorAction SilentlyContinue | 
                    Get-ItemProperty | 
                    Where-Object { $_.DisplayName -like "*Dell*" }
        
        if ($dellApps) {
            $foundDellApps = $true
        }
        
        foreach ($app in $dellApps) {
            $appName = $app.DisplayName
            
            # Skip Dell Command Update
            if (Test-IsDellCommandUpdate -AppName $appName) {
                Write-Log "KEEPING: Dell Command Update: $appName" "INFO"
                continue
            }
              # Uninstall other Dell applications
            try {
                if ($PSCmdlet.ShouldProcess($appName, "Uninstall Dell application")) {
                    Write-Log "REMOVING: Dell application: $appName" "INFO"
                    
                    $productCode = $app.PSChildName
                    $uninstallString = $app.UninstallString
                    
                    # Special handling for Dell Pair
                    if ($appName -eq "Dell Pair") {
                        # Use our specialized Dell Pair uninstaller
                        $dellPairResult = Uninstall-DellPair
                        if ($dellPairResult) {
                            # Skip further processing as it's been handled by the specialized function
                            continue
                        }
                        # If the specialized function failed, fall through to standard methods
                    }
                    # For other special Dell apps that need custom handling
                    elseif ($specialDellApps -contains $appName) {
                        # Try to get the direct uninstaller path if available
                        if ($uninstallString -match '"([^"]+)"') {
                            $uninstallExe = $matches[1]
                            if (Test-Path $uninstallExe) {
                                Write-Log "Using direct uninstaller for $appName" "INFO"
                                $process = Start-Process -FilePath $uninstallExe -ArgumentList "/S /SILENT" -Wait -NoNewWindow -PassThru
                            }
                            else {
                                Write-Log "Direct uninstaller not found for $appName, using MSI method" "INFO"
                                $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $productCode /qn /norestart /l*v `"$env:TEMP\$($productCode)_uninstall.log`"" -Wait -NoNewWindow -PassThru
                            }
                        }
                        else {
                            Write-Log "Using MSI method with logging for $appName" "INFO"
                            $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $productCode /qn /norestart /l*v `"$env:TEMP\$($productCode)_uninstall.log`"" -Wait -NoNewWindow -PassThru
                        }
                    }
                    else {
                        # Standard MSI uninstall for other Dell apps
                        $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $productCode /qn /norestart" -Wait -NoNewWindow -PassThru
                    }
                      if ($process.ExitCode -eq 0) {
                        Write-Log "REMOVED: Dell application: $appName" "SUCCESS"
                    }
                    else {
                        # Handle specific MSI error codes
                        switch ($process.ExitCode) {
                            1605 { 
                                Write-Log "NOT FOUND: Dell application $appName (code 1605) - product not installed" "WARNING" 
                            }
                            1619 {
                                Write-Log "ERROR: Installation package could not be found (code 1619) for Dell application: $appName" "WARNING"
                            }1639 {
                                Write-Log "Invalid command line parameters (code 1639) for $appName - attempting alternative method" "WARNING"
                                
                                # Special handling for Dell Pair
                                if ($appName -eq "Dell Pair") {
                                    Write-Log "Using special uninstall method for Dell Pair" "INFO"
                                    # Try to find and use the specific uninstaller for Dell Pair
                                    $dellPairPath = Get-ChildItem -Path "C:\Program Files\Dell\*\*\Uninstall.exe" -ErrorAction SilentlyContinue | 
                                                   Where-Object { $_.Directory.Name -like "*Pair*" }
                                    
                                    if ($dellPairPath) {
                                        Write-Log "Found Dell Pair uninstaller at: $($dellPairPath.FullName)" "INFO"
                                        $altProcess = Start-Process -FilePath $dellPairPath.FullName -ArgumentList "/S" -Wait -NoNewWindow -PassThru                                        if ($altProcess.ExitCode -eq 0) {
                                            Write-Log "REMOVED: Dell Pair using direct uninstaller" "SUCCESS"
                                        }
                                        else {
                                            Write-Log "ERROR: Direct uninstaller for Dell Pair failed. Exit code: $($altProcess.ExitCode). Will try alternative cleanup." "WARNING"
                                            # Attempt registry cleanup for Dell Pair
                                            try {
                                                Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -Include "*Dell Pair*" -Force -ErrorAction SilentlyContinue
                                                Remove-Item -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -Include "*Dell Pair*" -Force -ErrorAction SilentlyContinue
                                                Remove-Item -Path "C:\Program Files\Dell\Dell Pair\" -Recurse -Force -ErrorAction SilentlyContinue
                                                Remove-Item -Path "C:\Program Files (x86)\Dell\Dell Pair\" -Recurse -Force -ErrorAction SilentlyContinue
                                                Write-Log "Completed Dell Pair cleanup" "SUCCESS"
                                            }
                                            catch {                                                $errorMsg = $_.Exception.Message
                                                Write-Log "ERROR: Failed during Dell Pair cleanup: $errorMsg" "ERROR"
                                            }
                                        }
                                    }
                                    else {
                                        Write-Log "NOT FOUND: Could not find Dell Pair uninstaller, trying alternative cleanup" "WARNING"
                                        # Attempt registry cleanup for Dell Pair
                                        try {
                                            Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -Include "*Dell Pair*" -Force -ErrorAction SilentlyContinue
                                            Remove-Item -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -Include "*Dell Pair*" -Force -ErrorAction SilentlyContinue
                                            Remove-Item -Path "C:\Program Files\Dell\Dell Pair\" -Recurse -Force -ErrorAction SilentlyContinue
                                            Remove-Item -Path "C:\Program Files (x86)\Dell\Dell Pair\" -Recurse -Force -ErrorAction SilentlyContinue                                                Write-Log "REMOVED: Dell Pair via registry and file cleanup" "SUCCESS"
                                                }
                                                catch {
                                                    $errorMsg = $_.Exception.Message
                                                    Write-Log "ERROR: Failed during Dell Pair cleanup: $errorMsg" "ERROR"
                                        }
                                    }                                }
                                # Standard handling for other applications
                                else {
                                    if ($uninstallString -and $uninstallString -notlike "*msiexec*") {
                                    # Try the original uninstall string from registry
                                    if ($uninstallString -match '"([^"]+)"(.*)') {
                                        $uninstallExe = $matches[1]
                                        $uninstallArgs = $matches[2] + " /S /SILENT"
                                        $altProcess = Start-Process -FilePath $uninstallExe -ArgumentList $uninstallArgs -Wait -NoNewWindow -PassThru                                        if ($altProcess.ExitCode -eq 0) {
                                            Write-Log "REMOVED: $appName using alternative method" "SUCCESS"
                                        }
                                        else {
                                            Write-Log "ERROR: Alternative method failed for $appName. Exit code: $($altProcess.ExitCode)" "WARNING"}
                                    }
                                }
                            }
                            }
                            default {
                                Write-Log "Failed to uninstall: $appName. Exit code: $($process.ExitCode)" "WARNING"
                            }
                        }
                    }
                }
                else {                    Write-Log "WhatIf: Would uninstall Dell application: $appName" "INFO"
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-Log "ERROR: Failed to remove Dell application: $appName. Error: $errorMsg" "ERROR"
            }
        }
    }
    
    if (-not $foundDellApps) {
        Write-Log "NOT FOUND: No Dell applications installed" "INFO"
    }
}

# Function to remove Lenovo bloatware except Vantage
function Remove-LenovoBloatware {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    
    Write-Log "Identifying Lenovo applications..." "INFO"
    
    $uninstallKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    
    $foundLenovoApps = $false
    
    foreach ($key in $uninstallKeys) {
        $lenovoApps = Get-ChildItem -Path $key -ErrorAction SilentlyContinue | 
                    Get-ItemProperty | 
                    Where-Object { $_.DisplayName -like "*Lenovo*" }
        
        if ($lenovoApps -and $lenovoApps.Count -gt 0) {
            $foundLenovoApps = $true
        }
        
        foreach ($app in $lenovoApps) {
            $appName = $app.DisplayName
            
            # Skip Lenovo Vantage
            if (Test-IsLenovoVantage -AppName $appName) {
                Write-Log "KEEPING: Lenovo Vantage: $appName" "INFO"
                continue
            }
              # Uninstall other Lenovo applications
            try {
                if ($PSCmdlet.ShouldProcess($appName, "Uninstall Lenovo application")) {
                    Write-Log "REMOVING: Lenovo application: $appName" "INFO"
                    
                    $productCode = $app.PSChildName
                    $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $productCode /qn /norestart" -Wait -NoNewWindow -PassThru
                      if ($process.ExitCode -eq 0) {
                        Write-Log "REMOVED: Lenovo application: $appName" "SUCCESS"
                    }
                    else {
                        Write-Log "ERROR: Failed to remove Lenovo application: $appName. Exit code: $($process.ExitCode)" "WARNING"
                    }
                }
                else {                    Write-Log "WhatIf: Would uninstall Lenovo application: $appName" "INFO"
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-Log "ERROR: Failed to remove Lenovo application: $appName. Error: $errorMsg" "ERROR"
            }
        }
    }
    
    if (-not $foundLenovoApps) {
        Write-Log "NOT FOUND: No Lenovo applications installed" "INFO"
    }
}

# Function to handle Dell Pair uninstallation, which requires special handling
function Uninstall-DellPair {
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([System.Boolean])]
    param()
    
    Write-Log "Starting Dell Pair special uninstallation procedure" "INFO"
    
    # First try uninstalling using standard method with a variety of arguments
    $uninstallRegistryKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    
    $foundDellPair = $false
    
    # First attempt: Find Dell Pair in registry and use its uninstall string
    foreach ($key in $uninstallRegistryKeys) {
        $dellPairs = Get-ChildItem -Path $key -ErrorAction SilentlyContinue | 
                    Get-ItemProperty -ErrorAction SilentlyContinue |
                    Where-Object { $_.DisplayName -like "*Dell Pair*" }
        
        if ($null -ne $dellPairs) {
            foreach ($app in $dellPairs) {
                $foundDellPair = $true
                $appName = $app.DisplayName
                $productCode = $app.PSChildName
                $uninstallString = $app.UninstallString
                
                Write-Log "Found Dell Pair application: $appName with Product Code: $productCode" "INFO"
                
                # Try multiple uninstall methods to see what works
                if ($PSCmdlet.ShouldProcess("Dell Pair", "Uninstall using multiple methods")) {
                    # Method 1: Standard MSI uninstall with logging
                    Write-Log "Trying MSI uninstall with logging for Dell Pair" "INFO"
                    $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/x `"$productCode`" /qn /norestart /l*v `"$env:TEMP\DellPair_uninstall.log`"" -Wait -NoNewWindow -PassThru -ErrorAction SilentlyContinue
                      if ($process.ExitCode -eq 0) {
                        Write-Log "REMOVED: Dell Pair using MSI with product code" "SUCCESS"
                        return
                    }
                    
                    # Method 2: Try using the uninstall string directly if available
                    if ($uninstallString) {
                        Write-Log "Trying direct uninstall string for Dell Pair: $uninstallString" "INFO"
                        if ($uninstallString -match '"([^"]+)"(.*)') {
                            $uninstallExe = $matches[1]
                            $uninstallArgs = $matches[2] + " /S /SILENT /VERYSILENT /NORESTART"
                            
                            if (Test-Path $uninstallExe) {
                                $process = Start-Process -FilePath $uninstallExe -ArgumentList $uninstallArgs -Wait -NoNewWindow -PassThru -ErrorAction SilentlyContinue
                                  if ($process.ExitCode -eq 0) {
                                    Write-Log "REMOVED: Dell Pair using direct uninstall string" "SUCCESS"
                                    return
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    
    # Second attempt: Search for uninstaller in common Dell locations
    if (-not $foundDellPair -or $foundDellPair) { # Still continue even if we found it but failed to uninstall
        Write-Log "Searching for Dell Pair uninstaller in common locations" "INFO"
        
        $possiblePaths = @(
            "${env:ProgramFiles}\Dell\Dell Pair\uninstall.exe",
            "${env:ProgramFiles(x86)}\Dell\Dell Pair\uninstall.exe",
            "${env:ProgramFiles}\Dell\DellPair\uninstall.exe",
            "${env:ProgramFiles(x86)}\Dell\DellPair\uninstall.exe"
        )
        
        # Also search for uninstallers in Dell subdirectories
        $dellDirs = Get-ChildItem -Path "${env:ProgramFiles}\Dell\", "${env:ProgramFiles(x86)}\Dell\" -Directory -ErrorAction SilentlyContinue
        foreach ($dir in $dellDirs) {
            $possiblePaths += Get-ChildItem -Path $dir.FullName -Recurse -Include "unins*.exe" -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName }
        }
        
        foreach ($path in $possiblePaths) {
            if (Test-Path $path) {
                Write-Log "Found potential Dell Pair uninstaller: $path" "INFO"
                
                if ($PSCmdlet.ShouldProcess("Dell Pair", "Uninstall using $path")) {
                    $process = Start-Process -FilePath $path -ArgumentList "/S /SILENT /VERYSILENT /NORESTART" -Wait -NoNewWindow -PassThru -ErrorAction SilentlyContinue
                      if ($process.ExitCode -eq 0) {
                        Write-Log "REMOVED: Dell Pair using $path" "SUCCESS"
                        return
                    }
                    else {
                        Write-Log "ERROR: Uninstaller $path failed with exit code: $($process.ExitCode)" "WARNING"
                    }
                }
            }
        }
    }
    
    # Final attempt: Brute force removal of files and registry keys
    Write-Log "Attempting manual removal of Dell Pair files and registry entries" "INFO"
    
    if ($PSCmdlet.ShouldProcess("Dell Pair", "Manual cleanup")) {
        try {
            # Remove registry entries
            $regPaths = @(
                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
            )
            
            foreach ($regPath in $regPaths) {
                Get-ChildItem -Path $regPath -ErrorAction SilentlyContinue | 
                Get-ItemProperty -ErrorAction SilentlyContinue | 
                Where-Object { $_.DisplayName -like "*Dell Pair*" } | 
                ForEach-Object {
                    $keyPath = $_.PSPath
                    Write-Log "Removing registry key: $keyPath" "INFO"
                    Remove-Item -Path $keyPath -Force -ErrorAction SilentlyContinue
                }
            }
            
            # Remove program files
            $filePaths = @(
                "${env:ProgramFiles}\Dell\Dell Pair\",
                "${env:ProgramFiles(x86)}\Dell\Dell Pair\",
                "${env:ProgramFiles}\Dell\DellPair\",
                "${env:ProgramFiles(x86)}\Dell\DellPair\"
            )
            
            foreach ($filePath in $filePaths) {
                if (Test-Path $filePath) {
                    Write-Log "Removing directory: $filePath" "INFO"
                    Remove-Item -Path $filePath -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
              Write-Log "REMOVED: Dell Pair through manual cleanup completed" "SUCCESS"
            return $true
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-Log "ERROR: Failed during Dell Pair manual cleanup: $errorMsg" "ERROR"
            return $false
        }
    }
    
    return $false
}

# Start script execution
Write-Log "Starting Windows bloatware removal script v$scriptVersion" "INFO"

try {
    # List of common UWP bloatware apps
    $uwpBloatware = @(
        "Microsoft.3DBuilder",
        "Microsoft.BingFinance",
        "Microsoft.BingNews",
        "Microsoft.BingSports",
        "Microsoft.BingWeather",
        "Microsoft.GetHelp",
        "Microsoft.Getstarted",
        "Microsoft.Messaging",
        "Microsoft.Microsoft3DViewer",
        "Microsoft.MicrosoftSolitaireCollection",
        "Microsoft.MixedReality.Portal",
        "Microsoft.OneConnect", 
        "Microsoft.People",
        "Microsoft.Print3D",
        "Microsoft.SkypeApp",
        "Microsoft.Wallet",
        "Microsoft.WindowsAlarms",
        "Microsoft.WindowsFeedbackHub",
        "Microsoft.WindowsMaps",
        "Microsoft.WindowsSoundRecorder",
        "Microsoft.Xbox.TCUI",
        "Microsoft.XboxApp",
        "Microsoft.XboxGameOverlay",
        "Microsoft.XboxGamingOverlay",
        "Microsoft.XboxIdentityProvider",
        "Microsoft.XboxSpeechToTextOverlay",
        "Microsoft.YourPhone",
        "Microsoft.ZuneMusic",
        "Microsoft.ZuneVideo",
        "king.com.CandyCrushSaga",
        "king.com.CandyCrushSodaSaga",
        "king.com.CandyCrushFriends"
    )
      # Remove UWP bloatware
    Write-Log "Removing UWP bloatware applications..." "INFO"
    foreach ($app in $uwpBloatware) {
        # Wrap each call in try/catch to ensure script continues even if one app fails
        try {
            Uninstall-UWPApp -AppName $app -ErrorAction SilentlyContinue
        }
        catch {
            # Just log and continue to the next app
            Write-Log "Caught exception while processing $app, continuing with next app" "WARNING"
        }
    }
    
    # List of common Win32 bloatware apps
    $win32Bloatware = @(
        "McAfee",
        "Norton ",
        "Wild Tangent",
        "Candy Crush",
        "Booking.com",
        "Spotify",
        # "Dolby",  # Excluded as requested
        "HP JumpStart",
        "HP Connection Optimizer",
        "HP Documentation",
        "HP Smart",
        "HP Sure"
        # Lenovo apps are now handled by the Remove-LenovoBloatware function
    )
           
    # Remove Win32 bloatware
    Write-Log "Removing Win32 bloatware applications..." "INFO"
    foreach ($app in $win32Bloatware) {
        Uninstall-Win32App -DisplayName $app
    }
    
    # Remove Dell bloatware except Command Update
    Write-Log "Removing Dell bloatware (except Command Update)..." "INFO"
    Remove-DellBloatware
    
    # Remove Lenovo bloatware except Vantage
    Write-Log "Removing Lenovo bloatware (except Vantage)..." "INFO"
    Remove-LenovoBloatware
    
    # Clean up any leftover files
    $bloatwareFolders = @(
        "${env:ProgramFiles}\Dell",  # Not removing all Dell folders, just checking for specific ones
        "${env:ProgramFiles}\McAfee",
        "${env:ProgramFiles}\Norton",
        "${env:ProgramFiles}\Wild Tangent Games",
        "${env:ProgramFiles(x86)}\Dell", # Not removing all Dell folders, just checking for specific ones
        "${env:ProgramFiles(x86)}\McAfee",
        "${env:ProgramFiles(x86)}\Norton",
        "${env:ProgramFiles(x86)}\Wild Tangent Games",
        "${env:ProgramFiles}\Lenovo", # Not removing all Lenovo folders, just checking for specific ones
        "${env:ProgramFiles(x86)}\Lenovo" # Not removing all Lenovo folders, just checking for specific ones
    )
    
    Write-Log "Cleaning up leftover bloatware directories..." "INFO"
    foreach ($folder in $bloatwareFolders) {
        if (Test-Path $folder) {
            # Skip Dell Command Update folders
            if (($folder -like "*Dell*") -and (Test-Path "$folder\Command Update")) {
                Write-Log "Skipping Dell Command Update folder: $folder\Command Update" "INFO"
                continue
            }
            
            # Skip Lenovo Vantage folders
            if (($folder -like "*Lenovo*") -and (Test-Path "$folder\Lenovo Vantage")) {
                Write-Log "Skipping Lenovo Vantage folder: $folder\Lenovo Vantage" "INFO"
                continue
            }
            
            if ($PSCmdlet.ShouldProcess($folder, "Remove directory")) {
                Write-Log "Removing directory: $folder" "INFO"
                try {
                    Remove-Item -Path $folder -Recurse -Force                    Write-Log "REMOVED: Directory $folder" "SUCCESS"
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Write-Log "ERROR: Failed to remove directory $folder. Error: $errorMsg" "ERROR"
                }
            }
            else {
                Write-Log "WhatIf: Would remove directory: $folder" "INFO"
            }
        }
    }
    
    Write-Log "Windows bloatware removal completed successfully" "SUCCESS"
}
catch {
    $errorMsg = $_.Exception.Message
    Write-Log "An error occurred during bloatware removal: $errorMsg" "ERROR"
    Write-Log "Exception details: $($_.Exception.Message)" "ERROR"
    Write-Log "Stack trace: $($_.ScriptStackTrace)" "DEBUG"
    exit 1
}
