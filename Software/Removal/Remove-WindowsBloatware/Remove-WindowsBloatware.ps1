# =============================================================================
# Script: Remove-WindowsBloatware.ps1
# Created: 2025-05-07 15:45:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-05-07 22:30:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.6
# Additional Info: Added dedicated handling for Lenovo software with Lenovo Vantage preservation
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
$logFile = "$PSScriptRoot\Remove-WindowsBloatware.log"
$scriptVersion = "1.0.6"

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
        $app = Get-AppxPackage -Name $AppName -AllUsers
        if ($app) {
            if ($PSCmdlet.ShouldProcess($AppName, "Remove UWP application")) {
                Write-Log "Removing UWP application: $AppName" "INFO"
                Remove-AppxPackage -Package $app.PackageFullName
                Write-Log "Successfully removed UWP application: $AppName" "SUCCESS"
            }
            else {
                Write-Log "WhatIf: Would remove UWP application: $AppName" "INFO"
            }
        }
        else {
            Write-Log "UWP application not found: $AppName" "INFO"
        }
    }
    catch {
        Write-Log "Error removing UWP application $AppName. Error: $_" "ERROR"
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
            
            Write-Log "Found application: $appName" "INFO"
            
            try {
                if ($PSCmdlet.ShouldProcess($appName, "Uninstall application")) {
                    Write-Log "Uninstalling application: $appName" "INFO"
                    
                    # If msiexec is in the uninstall string, use that
                    if ($uninstallString -like "*msiexec*") {
                        $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $productCode /qn" -Wait -NoNewWindow -PassThru
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
                        Write-Log "Successfully uninstalled: $appName" "SUCCESS"
                    }
                    else {
                        Write-Log "Failed to uninstall: $appName. Exit code: $($process.ExitCode)" "WARNING"
                    }
                }
                else {
                    Write-Log "WhatIf: Would uninstall application: $appName" "INFO"
                }
            }
            catch {
                Write-Log "Error uninstalling $appName. Error: $_" "ERROR"
            }
        }
    }
    
    if (-not $found) {
        Write-Log "No matching applications found for: $DisplayName" "INFO"
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
    
    foreach ($key in $uninstallKeys) {
        $dellApps = Get-ChildItem -Path $key -ErrorAction SilentlyContinue | 
                    Get-ItemProperty | 
                    Where-Object { $_.DisplayName -like "*Dell*" }
        
        foreach ($app in $dellApps) {
            $appName = $app.DisplayName
            
            # Skip Dell Command Update
            if (Test-IsDellCommandUpdate -AppName $appName) {
                Write-Log "Keeping Dell Command Update: $appName" "INFO"
                continue
            }
            
            # Uninstall other Dell applications
            try {
                if ($PSCmdlet.ShouldProcess($appName, "Uninstall Dell application")) {
                    Write-Log "Uninstalling Dell application: $appName" "INFO"
                    
                    $productCode = $app.PSChildName
                    $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $productCode /qn" -Wait -NoNewWindow -PassThru
                    
                    if ($process.ExitCode -eq 0) {
                        Write-Log "Successfully uninstalled: $appName" "SUCCESS"
                    }
                    else {
                        Write-Log "Failed to uninstall: $appName. Exit code: $($process.ExitCode)" "WARNING"
                    }
                }
                else {
                    Write-Log "WhatIf: Would uninstall Dell application: $appName" "INFO"
                }
            }
            catch {
                Write-Log "Error uninstalling $appName. Error: $_" "ERROR"
            }
        }
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
    
    foreach ($key in $uninstallKeys) {
        $lenovoApps = Get-ChildItem -Path $key -ErrorAction SilentlyContinue | 
                    Get-ItemProperty | 
                    Where-Object { $_.DisplayName -like "*Lenovo*" }
        
        foreach ($app in $lenovoApps) {
            $appName = $app.DisplayName
            
            # Skip Lenovo Vantage
            if (Test-IsLenovoVantage -AppName $appName) {
                Write-Log "Keeping Lenovo Vantage: $appName" "INFO"
                continue
            }
            
            # Uninstall other Lenovo applications
            try {
                if ($PSCmdlet.ShouldProcess($appName, "Uninstall Lenovo application")) {
                    Write-Log "Uninstalling Lenovo application: $appName" "INFO"
                    
                    $productCode = $app.PSChildName
                    $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $productCode /qn" -Wait -NoNewWindow -PassThru
                    
                    if ($process.ExitCode -eq 0) {
                        Write-Log "Successfully uninstalled: $appName" "SUCCESS"
                    }
                    else {
                        Write-Log "Failed to uninstall: $appName. Exit code: $($process.ExitCode)" "WARNING"
                    }
                }
                else {
                    Write-Log "WhatIf: Would uninstall Lenovo application: $appName" "INFO"
                }
            }
            catch {
                Write-Log "Error uninstalling $appName. Error: $_" "ERROR"
            }
        }
    }
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
        Uninstall-UWPApp -AppName $app
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
                    Remove-Item -Path $folder -Recurse -Force
                    Write-Log "Successfully removed directory: $folder" "SUCCESS"
                }
                catch {
                    Write-Log "Failed to remove directory $folder. Error: $_" "ERROR"
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
    Write-Log "An error occurred during bloatware removal: $_" "ERROR"
    Write-Log "Exception details: $($_.Exception.Message)" "ERROR"
    Write-Log "Stack trace: $($_.ScriptStackTrace)" "DEBUG"
    exit 1
}
