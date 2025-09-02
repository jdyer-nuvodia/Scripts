# =============================================================================
# Script: Remove-WindowsBloatware.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-08-28 16:04:12 UTC
# Updated By: jdyer-nuvodia
# Version: 3.3.3
# Additional Info: Added HP Wolf Security to the removal list
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
7. Disables and stops unnecessary system services

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
- Microsoft OneNote for Windows 10
- Candy Crush games (Saga, Soda Saga, Friends)

Traditional Win32 Applications:
- McAfee Security Software
- Norton Security Software
- Wild Tangent Games
- Candy Crush desktop apps
- Booking.com apps
- Spotify
- HP pre-installed software (JumpStart, Connection Optimizer, Documentation, Smart, Sure, Wolf Security)
- Lenovo pre-installed software EXCEPT Lenovo Vantage
- All Dell software EXCEPT Dell Command Update

Windows Services:
- Elliptic Virtual Lock Sensor Service
- Intel Context Sensing Service

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

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param(
    [Parameter(Mandatory = $false)]
    [switch]$Force
)

# Run this script as an administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You do not have Administrator rights to run this script!`nPlease re-run this script as an Administrator!"
    break
}

# Script variables
$computerName = $env:COMPUTERNAME
$utcTimestamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd_HH-mm-ss")
$logFile = $PSScriptRoot + "\Remove-WindowsBloatware_" + $computerName + "_" + $utcTimestamp + ".log"
$scriptVersion = "1.1.11"

# Function to write log entries
function Write-LogEntry {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS", "DEBUG")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [$Level] $Message"

    # Output to console without using Write-Host
    switch ($Level) {
        "INFO" { Write-Output $logEntry }
        "WARNING" { Write-Warning $logEntry }
        "ERROR" { Write-Error $logEntry }
        "SUCCESS" { Write-Output $logEntry }
        "DEBUG" { Write-Verbose $logEntry }
        default { Write-Output $logEntry }
    }

    # Create log directory if it does not exist
    $logDir = Split-Path -Path $logFile -Parent
    if (-not (Test-Path -Path $logDir)) {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    }

    # Write to log file
    Add-Content -Path $logFile -Value $logEntry
}

# Helper function to remove directories without confirmation
function Remove-DirectorySilently {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([System.Boolean])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $false)]
        [string]$LogPrefix = "",

        [Parameter(Mandatory = $false)]
        [switch]$SkipProtectedFolders,

        [Parameter(Mandatory = $false)]
        [switch]$Force
    )

    # If Force parameter is passed, use it
    $useForce = $Force.IsPresent

    if (Test-Path $Path) {
        if ($PSCmdlet.ShouldProcess($Path, "Remove directory")) {
            Write-LogEntry "$LogPrefix Removing directory: $Path" "INFO"

            # Check if this is a manufacturer system folder that should be handled carefully
            $isManufacturerFolder = ($Path -like "*\Dell*" -or $Path -like "*\Lenovo*" -or $Path -like "*\HP*")

            # If this is a manufacturer folder and we're not skipping protected items
            # and force isn't enabled, use selective removal
            if ($isManufacturerFolder -and -not $SkipProtectedFolders -and -not $useForce) {
                try {
                    # Get list of files/folders in the directory
                    $items = Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue

                    # Count how many items might be in use
                    $protectedItems = @()
                    $deletableItems = @()

                    # Separate items into protected (likely in use) and deletable
                    foreach ($item in $items) {
                        if ($item.PSIsContainer) {
                            # It's a directory, check if it contains special protected folders
                            if ($item.Name -like "*Command Update*" -or
                                $item.Name -like "*Vantage*" -or
                                $item.Name -like "*UpdateService*" -or
                                $item.Name -like "*Service*") {
                                $protectedItems += $item.FullName
                            } else {
                                $deletableItems += $item.FullName
                            }
                        } else {
                            # It's a file, check if it's likely to be in use
                            if ($item.Extension -eq ".dll" -or
                                $item.Extension -eq ".exe" -or
                                $item.Extension -eq ".sys" -or
                                $item.Name -like "*lock*") {
                                # These files might be locked, skip them
                                $protectedItems += $item.FullName
                            } else {
                                $deletableItems += $item.FullName
                            }
                        }
                    }

                    # Log information about protected items
                    if ($protectedItems.Count -gt 0) {
                        Write-LogEntry "$LogPrefix INFO: Skipping $($protectedItems.Count) protected items in $Path" "INFO"
                    }

                    # Try to delete non-protected items
                    foreach ($item in $deletableItems) {
                        try {
                            if (Test-Path $item) {
                                # Use -Force and -Confirm:$false to prevent prompting
                                Remove-Item -Path $item -Force -Confirm:$false -ErrorAction SilentlyContinue
                            }
                        } catch {
                            # Log the error but continue processing
                            Write-LogEntry "$LogPrefix WARNING: Could not remove item $item - $($_.Exception.Message)" "DEBUG"
                        }
                    }

                    Write-LogEntry "$LogPrefix PARTIAL: Directory cleanup of $Path completed with some items skipped" "WARNING"
                    return $true
                } catch {
                    $errorMsg = $_.Exception.Message
                    Write-LogEntry "$LogPrefix ERROR: Failed selective directory cleanup for $Path. Error: $errorMsg" "ERROR"
                    return $false
                }
            } else {
                # Standard removal for non-manufacturer folders or when force/skip is enabled
                try {
                    # Use -Force to override read-only attributes, -Recurse to remove subdirectories,
                    # -Confirm:$false to suppress confirmation, and -ErrorAction Stop to catch errors
                    Remove-Item -Path $Path -Recurse -Force -Confirm:$false -ErrorAction Stop
                    Write-LogEntry "$LogPrefix REMOVED: Directory $Path" "SUCCESS"
                    return $true
                } catch {
                    $errorMsg = $_.Exception.Message
                    Write-LogEntry "$LogPrefix ERROR: Failed to remove directory $Path. Error: $errorMsg" "ERROR"
                    return $false
                }
            }
        } else {
            Write-LogEntry "$LogPrefix WhatIf: Would remove directory: $Path" "INFO"
            return $true
        }
    }
    return $false
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
                Write-LogEntry "FOUND: Multiple instances of UWP application: $AppName" "INFO"
                foreach ($singleApp in $app) {
                    if ($PSCmdlet.ShouldProcess($singleApp.Name, "Remove UWP application")) {
                        Write-LogEntry "REMOVING: UWP application instance: $($singleApp.Name) (PackageFullName: $($singleApp.PackageFullName))" "INFO"
                        try {
                            Remove-AppxPackage -Package $singleApp.PackageFullName -ErrorAction SilentlyContinue
                            Write-LogEntry "REMOVED: UWP application instance: $($singleApp.Name)" "SUCCESS"
                        } catch {
                            # Log the error but don't display it to the console
                            $errorMsg = $_.Exception.Message
                            Write-LogEntry "ERROR: Failed to remove UWP application instance $($singleApp.Name): $errorMsg" "WARNING"
                        }
                    } else {
                        Write-LogEntry "WhatIf: Would remove UWP application instance: $($singleApp.Name)" "INFO"
                    }
                }
            } else {
                # Single app instance
                if ($PSCmdlet.ShouldProcess($app.Name, "Remove UWP application")) {
                    Write-LogEntry "REMOVING: UWP application: $AppName (PackageFullName: $($app.PackageFullName))" "INFO"
                    try {
                        Remove-AppxPackage -Package $app.PackageFullName -ErrorAction SilentlyContinue
                        Write-LogEntry "REMOVED: UWP application: $AppName" "SUCCESS"
                    } catch {
                        # Log the error but don't display it to the console
                        $errorMsg = $_.Exception.Message
                        Write-LogEntry "ERROR: Failed to remove UWP application ${ AppName}: $errorMsg" "WARNING"
                    }
                } else {
                    Write-LogEntry "WhatIf: Would remove UWP application: $AppName" "INFO"
                }
            }
        } else {
            Write-LogEntry "NOT FOUND: UWP application $AppName is not installed or was previously removed" "INFO"
        }
    } catch {
        # Log the error but don't display it to the console
        $errorMsg = $_.Exception.Message
        Write-LogEntry "ERROR: Failed to access UWP application ${ AppName}: $errorMsg" "WARNING"
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

    Write-LogEntry "Searching for Win32 application: $DisplayName" "INFO"

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

            Write-LogEntry "FOUND: Win32 application: $appName" "INFO"

            try {
                if ($PSCmdlet.ShouldProcess($appName, "Uninstall application")) {
                    Write-LogEntry "REMOVING: Win32 application: $appName" "INFO"
                    # If msiexec is in the uninstall string, use that
                    if ($uninstallString -like "*msiexec*") {
                        # Using /qn (no UI), /norestart (prevent restart), /passive (progress bar only, no user input)
                        $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $productCode /qn /norestart" -Wait -NoNewWindow -PassThru
                    } else {
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
                        Write-LogEntry "REMOVED: Win32 application: $appName" "SUCCESS"
                    } else {
                        Write-LogEntry "ERROR: Failed to remove Win32 application: $appName. Exit code: $($process.ExitCode)" "WARNING"
                    }
                } else {
                    Write-LogEntry "WhatIf: Would remove Win32 application: $appName" "INFO"
                }
            } catch {
                $errorMsg = $_.Exception.Message
                Write-LogEntry "ERROR: Failed to remove Win32 application: $appName. Error: $errorMsg" "ERROR"
            }
        }
    }

    if (-not $found) {
        Write-LogEntry "NOT FOUND: Win32 application: $DisplayName is not installed" "INFO"
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

    Write-LogEntry "Identifying Dell applications..." "INFO"

    $uninstallKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    # Special case Dell applications that need custom handling
    $specialDellApps = @(
        "Dell Pair",
        "Dell SupportAssist OS Recovery Plugin for Dell Update",
        "Dell Digital Delivery",
        "Partner Promo"
    )

    $foundDellApps = $false
    foreach ($key in $uninstallKeys) {
        # Look for Dell apps and specific apps like Partner Promo that might not have "Dell" in the name
        $dellApps = Get-ChildItem -Path $key -ErrorAction SilentlyContinue |
            Get-ItemProperty |
            Where-Object {
                $_.DisplayName -like "*Dell*" -or
                $_.DisplayName -like "*Partner Promo*" -or
                $_.DisplayName -like "*SupportAssist OS Recovery*" -or
                $_.Publisher -like "*Dell*"
            }

        if ($dellApps) {
            $foundDellApps = $true
        }

        foreach ($app in $dellApps) {
            $appName = $app.DisplayName

            # Skip Dell Command Update
            if (Test-IsDellCommandUpdate -AppName $appName) {
                Write-LogEntry "KEEPING: Dell Command Update: $appName" "INFO"
                continue
            }
            # Uninstall other Dell applications
            try {
                if ($PSCmdlet.ShouldProcess($appName, "Uninstall Dell application")) {
                    Write-LogEntry "REMOVING: Dell application: $appName" "INFO"

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
                                Write-LogEntry "Using direct uninstaller for $appName" "INFO"
                                $process = Start-Process -FilePath $uninstallExe -ArgumentList "/S /SILENT" -Wait -NoNewWindow -PassThru
                            } else {
                                Write-LogEntry "Direct uninstaller not found for $appName, using MSI method" "INFO"
                                $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $productCode /qn /norestart /l*v `"$env:TEMP\$($productCode)_uninstall.log`"" -Wait -NoNewWindow -PassThru
                            }
                        } else {
                            Write-LogEntry "Using MSI method with logging for $appName" "INFO"
                            $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $productCode /qn /norestart /l*v `"$env:TEMP\$($productCode)_uninstall.log`"" -Wait -NoNewWindow -PassThru
                        }
                    } else {
                        # Standard MSI uninstall for other Dell apps
                        $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $productCode /qn /norestart" -Wait -NoNewWindow -PassThru
                    }
                    if ($process.ExitCode -eq 0) {
                        Write-LogEntry "REMOVED: Dell application: $appName" "SUCCESS"
                    } else {
                        # Handle specific MSI error codes
                        switch ($process.ExitCode) {
                            1605 {
                                Write-LogEntry "NOT FOUND: Dell application $appName (code 1605) - product not installed" "WARNING"
                            }
                            1619 {
                                Write-LogEntry "ERROR: Installation package could not be found (code 1619) for Dell application: $appName" "WARNING"
                            }
                            1639 {
                                Write-LogEntry "Invalid command line parameters (code 1639) for $appName - attempting alternative method" "WARNING"

                                # Special handling for Dell Pair
                                if ($appName -eq "Dell Pair") {
                                    Write-LogEntry "Using special uninstall method for Dell Pair" "INFO"
                                    # Try to find and use the specific uninstaller for Dell Pair
                                    $dellPairPath = Get-ChildItem -Path "C:\Program Files\Dell\*\*\Uninstall.exe" -ErrorAction SilentlyContinue |
                                        Where-Object { $_.Directory.Name -like "*Pair*" }

                                    if ($dellPairPath) {
                                        Write-LogEntry "Found Dell Pair uninstaller at: $($dellPairPath.FullName)" "INFO"
                                        $altProcess = Start-Process -FilePath $dellPairPath.FullName -ArgumentList "/S" -Wait -NoNewWindow -PassThru
                                        if ($altProcess.ExitCode -eq 0) {
                                            Write-LogEntry "REMOVED: Dell Pair using direct uninstaller" "SUCCESS"
                                        }

                                        else {
                                            # Attempt registry cleanup for Dell Pair
                                            Write-LogEntry "ERROR: Direct uninstaller for Dell Pair failed. Exit code: $($altProcess.ExitCode). Will try alternative cleanup." "WARNING"
                                            try {
                                                Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -Include "*Dell Pair*" -Force -Confirm:$false -ErrorAction SilentlyContinue
                                                Remove-Item -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -Include "*Dell Pair*" -Force -Confirm:$false -ErrorAction SilentlyContinue
                                                Remove-Item -Path "C:\Program Files\Dell\Dell Pair\" -Recurse -Force -Confirm:$false -ErrorAction SilentlyContinue
                                                Remove-Item -Path "C:\Program Files (x86)\Dell\Dell Pair\" -Recurse -Force -Confirm:$false -ErrorAction SilentlyContinue
                                                Write-LogEntry "REMOVED: Dell Pair via registry and file cleanup" "SUCCESS"
                                            } catch {
                                                $errorMsg = $_.Exception.Message
                                                Write-LogEntry "ERROR: Failed during Dell Pair cleanup: $errorMsg" "ERROR"
                                            }
                                        }
                                    } else {
                                        # Attempt registry cleanup for Dell Pair
                                        Write-LogEntry "NOT FOUND: Could not find Dell Pair uninstaller, trying alternative cleanup" "WARNING"
                                        try {
                                            Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -Include "*Dell Pair*" -Force -Confirm:$false -ErrorAction SilentlyContinue
                                            Remove-Item -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -Include "*Dell Pair*" -Force -Confirm:$false -ErrorAction SilentlyContinue
                                            Remove-Item -Path "C:\Program Files\Dell\Dell Pair\" -Recurse -Force -Confirm:$false -ErrorAction SilentlyContinue
                                            Remove-Item -Path "C:\Program Files (x86)\Dell\Dell Pair\" -Recurse -Force -Confirm:$false -ErrorAction SilentlyContinue
                                            Write-LogEntry "REMOVED: Dell Pair via registry and file cleanup" "SUCCESS"
                                        } catch {
                                            $errorMsg = $_.Exception.Message
                                            Write-LogEntry "ERROR: Failed during Dell Pair cleanup: $errorMsg" "ERROR"
                                        }
                                    }
                                }
                                # Standard handling for other applications
                                else {
                                    if ($uninstallString -and $uninstallString -notlike "*msiexec*") {
                                        # Try the original uninstall string from registry
                                        if ($uninstallString -match '"([^"]+)"(.*)') {
                                            $uninstallExe = $matches[1]
                                            $uninstallArgs = $matches[2] + " /S /SILENT"
                                            $altProcess = Start-Process -FilePath $uninstallExe -ArgumentList $uninstallArgs -Wait -NoNewWindow -PassThru
                                            if ($altProcess.ExitCode -eq 0) {
                                                Write-LogEntry "REMOVED: $appName using alternative method" "SUCCESS"
                                            }

                                            else {
                                                Write-LogEntry "ERROR: Alternative method failed for $appName. Exit code: $($altProcess.ExitCode)" "WARNING"
                                            }
                                        }
                                    }
                                }
                            }

                            default {
                                Write-LogEntry "Failed to uninstall: $appName. Exit code: $($process.ExitCode)" "WARNING"
                            }
                        }
                    }
                } else {
                    Write-LogEntry "WhatIf: Would uninstall Dell application: $appName" "INFO"
                }
            } catch {
                $errorMsg = $_.Exception.Message
                Write-LogEntry "ERROR: Failed to remove Dell application: $appName. Error: $errorMsg" "ERROR"
            }
        }
    }

    if (-not $foundDellApps) {
        Write-LogEntry "NOT FOUND: No Dell applications installed" "INFO"
    }
}

# Function to remove Lenovo bloatware except Vantage
function Remove-LenovoBloatware {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()

    Write-LogEntry "Identifying Lenovo applications..." "INFO"

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
                Write-LogEntry "KEEPING: Lenovo Vantage: $appName" "INFO"
                continue
            }
            # Uninstall other Lenovo applications
            try {
                if ($PSCmdlet.ShouldProcess($appName, "Uninstall Lenovo application")) {
                    Write-LogEntry "REMOVING: Lenovo application: $appName" "INFO"

                    $productCode = $app.PSChildName
                    $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $productCode /qn /norestart" -Wait -NoNewWindow -PassThru
                    if ($process.ExitCode -eq 0) {
                        Write-LogEntry "REMOVED: Lenovo application: $appName" "SUCCESS"
                    } else {
                        Write-LogEntry "ERROR: Failed to remove Lenovo application: $appName. Exit code: $($process.ExitCode)" "WARNING"
                    }
                } else {
                    Write-LogEntry "WhatIf: Would uninstall Lenovo application: $appName" "INFO"
                }
            } catch {
                $errorMsg = $_.Exception.Message
                Write-LogEntry "ERROR: Failed to remove Lenovo application: $appName. Error: $errorMsg" "ERROR"
            }
        }
    }

    if (-not $foundLenovoApps) {
        Write-LogEntry "NOT FOUND: No Lenovo applications installed" "INFO"
    }
}

# Function to handle Dell Pair uninstallation, which requires special handling
function Uninstall-DellPair {
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([System.Boolean])]
    param()

    Write-LogEntry "Starting Dell Pair special uninstallation procedure" "INFO"

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

                Write-LogEntry "Found Dell Pair application: $appName with Product Code: $productCode" "INFO"

                # Try multiple uninstall methods to see what works
                if ($PSCmdlet.ShouldProcess("Dell Pair", "Uninstall using multiple methods")) {
                    # Method 1: Standard MSI uninstall with logging
                    Write-LogEntry "Trying MSI uninstall with logging for Dell Pair" "INFO"
                    $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/x `"$productCode`" /qn /norestart /l*v `"$env:TEMP\DellPair_uninstall.log`"" -Wait -NoNewWindow -PassThru -ErrorAction SilentlyContinue
                    if ($process.ExitCode -eq 0) {
                        Write-LogEntry "REMOVED: Dell Pair using MSI with product code" "SUCCESS"
                        return
                    }

                    # Method 2: Try using the uninstall string directly if available
                    if ($uninstallString) {
                        Write-LogEntry "Trying direct uninstall string for Dell Pair: $uninstallString" "INFO"
                        if ($uninstallString -match '"([^"]+)"(.*)') {
                            $uninstallExe = $matches[1]
                            $uninstallArgs = $matches[2] + " /S /SILENT /VERYSILENT /NORESTART"

                            if (Test-Path $uninstallExe) {
                                $process = Start-Process -FilePath $uninstallExe -ArgumentList $uninstallArgs -Wait -NoNewWindow -PassThru -ErrorAction SilentlyContinue
                                if ($process.ExitCode -eq 0) {
                                    Write-LogEntry "REMOVED: Dell Pair using direct uninstall string" "SUCCESS"
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
    # Still continue even if we found it but failed to uninstall
    if (-not $foundDellPair -or $foundDellPair) {
        Write-LogEntry "Searching for Dell Pair uninstaller in common locations" "INFO"

        $possiblePaths = @(
            "${ env:ProgramFiles}\Dell\Dell Pair\uninstall.exe",
            "${ env:ProgramFiles(x86)}\Dell\Dell Pair\uninstall.exe",
            "${ env:ProgramFiles}\Dell\DellPair\uninstall.exe",
            "${ env:ProgramFiles(x86)}\Dell\DellPair\uninstall.exe"
        )

        # Also search for uninstallers in Dell subdirectories
        $dellDirs = Get-ChildItem -Path "${ env:ProgramFiles}\Dell\", "${ env:ProgramFiles(x86)}\Dell\" -Directory -ErrorAction SilentlyContinue
        foreach ($dir in $dellDirs) {
            $possiblePaths += Get-ChildItem -Path $dir.FullName -Recurse -Include "unins*.exe" -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName }
        }

        foreach ($path in $possiblePaths) {
            if (Test-Path $path) {
                Write-LogEntry "Found potential Dell Pair uninstaller: $path" "INFO"

                if ($PSCmdlet.ShouldProcess("Dell Pair", "Uninstall using $path")) {
                    $process = Start-Process -FilePath $path -ArgumentList "/S /SILENT /VERYSILENT /NORESTART" -Wait -NoNewWindow -PassThru -ErrorAction SilentlyContinue
                    if ($process.ExitCode -eq 0) {
                        Write-LogEntry "REMOVED: Dell Pair using $path" "SUCCESS"
                        return
                    } else {
                        Write-LogEntry "ERROR: Uninstaller $path failed with exit code: $($process.ExitCode)" "WARNING"
                    }
                }
            }
        }
    }

    # Final attempt: Brute force removal of files and registry keys
    Write-LogEntry "Attempting manual removal of Dell Pair files and registry entries" "INFO"

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
                        Write-LogEntry "Removing registry key: $keyPath" "INFO"
                        Remove-Item -Path $keyPath -Force -Confirm:$false -ErrorAction SilentlyContinue
                    }
            }
            # Remove program files
            $filePaths = @(
                "${ env:ProgramFiles}\Dell\Dell Pair\",
                "${ env:ProgramFiles(x86)}\Dell\Dell Pair\",
                "${ env:ProgramFiles}\Dell\DellPair\",
                "${ env:ProgramFiles(x86)}\Dell\DellPair\"
            )

            foreach ($filePath in $filePaths) {
                # Use our silent removal helper function
                Remove-DirectorySilently -Path $filePath -LogPrefix "Dell Pair cleanup:" -Force
            }
            Write-LogEntry "REMOVED: Dell Pair through manual cleanup completed" "SUCCESS"
            return $true
        } catch {
            $errorMsg = $_.Exception.Message
            Write-LogEntry "ERROR: Failed during Dell Pair manual cleanup: $errorMsg" "ERROR"
            return $false
        }
    }
    return $false
}

# Function to stop and disable unwanted services
function Stop-DisableBloatwareService {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServiceName,

        [Parameter(Mandatory = $false)]
        [string]$DisplayName = "",

        [Parameter(Mandatory = $false)]
        [switch]$Force
    )

    if ([string]::IsNullOrEmpty($DisplayName)) {
        $DisplayName = $ServiceName
    }

    Write-LogEntry "Checking for service: $DisplayName" "INFO"

    # First try to find service by name
    try {
        $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue

        # If not found by name, try to find by display name using a safer approach with WMI
        if ($null -eq $service -and -not [string]::IsNullOrEmpty($DisplayName)) {
            Write-LogEntry "Service not found by name, searching by display name: $DisplayName" "INFO"
            try {
                # Use WMI/CIM to query services instead of Get-Service which can have permission issues
                $foundServices = Get-CimInstance -ClassName Win32_Service -Filter "DisplayName LIKE '%$DisplayName%'" -ErrorAction SilentlyContinue

                if ($null -ne $foundServices -and ($foundServices | Measure-Object).Count -gt 0) {
                    # Take the first matching service
                    $foundService = $foundServices | Select-Object -First 1
                    Write-LogEntry "FOUND: Service with display name matching '$DisplayName' (Name: $($foundService.Name))" "INFO"

                    # Now get the service object using the name we found
                    $service = Get-Service -Name $foundService.Name -ErrorAction SilentlyContinue
                    if ($null -ne $service) {
                        # Update the ServiceName variable to match what was found
                        $ServiceName = $service.Name
                    }
                }
            } catch {
                $errorMsg = $_.Exception.Message
                Write-LogEntry "ERROR: Failed to search for service by display name: $errorMsg" "WARNING"
            }
        }

        if ($null -eq $service) {
            Write-LogEntry "NOT FOUND: Service $DisplayName is not installed" "INFO"
            return
        }

        Write-LogEntry "FOUND: Service $ServiceName (DisplayName: $($service.DisplayName), Status: $($service.Status))" "INFO"

        # First, try to stop the service if it's running
        if ($PSCmdlet.ShouldProcess($DisplayName, "Stop and disable service")) {
            if ($service.Status -eq "Running") {
                Write-LogEntry "STOPPING: Service $ServiceName" "INFO"
                try {
                    # Try multiple methods to stop the service
                    $stopSuccess = $false

                    # Method 1: Standard Stop-Service cmdlet
                    try {
                        if ($Force) {
                            Stop-Service -Name $ServiceName -Force -ErrorAction Stop
                        } else {
                            Stop-Service -Name $ServiceName -ErrorAction Stop
                        }
                        Write-LogEntry "STOPPED: Service $ServiceName using standard method" "SUCCESS"
                        $stopSuccess = $true
                    } catch {
                        $errorMsg = $_.Exception.Message
                        Write-LogEntry "WARNING: Standard stop failed for ${ ServiceName}: $errorMsg. Trying alternative methods." "WARNING"
                    }

                    # Method 2: Use WMI/CIM to stop the service if standard method failed
                    if (-not $stopSuccess) {
                        try {
                            $wmiService = Get-CimInstance -ClassName Win32_Service -Filter "Name = '$ServiceName'" -ErrorAction SilentlyContinue

                            if ($null -ne $wmiService) {
                                $result = $wmiService | Invoke-CimMethod -MethodName StopService
                                if ($result.ReturnValue -eq 0) {
                                    Write-LogEntry "STOPPED: Service $ServiceName using WMI method" "SUCCESS"
                                    $stopSuccess = $true
                                } else {
                                    Write-LogEntry "WARNING: Failed to stop service via WMI. Return code: $($result.ReturnValue)" "WARNING"
                                }
                            }
                        } catch {
                            $errorMsg = $_.Exception.Message
                            Write-LogEntry "WARNING: WMI stop failed for ${ ServiceName}: $errorMsg" "WARNING"
                        }
                    }

                    # Method 3: Use SC.exe command if both other methods failed
                    if (-not $stopSuccess) {
                        try {
                            $scResult = Start-Process -FilePath "sc.exe" -ArgumentList "stop $ServiceName" -NoNewWindow -Wait -PassThru
                            if ($scResult.ExitCode -eq 0) {
                                Write-LogEntry "STOPPED: Service $ServiceName using SC.exe" "SUCCESS"
                                $stopSuccess = $true
                            } else {
                                Write-LogEntry "WARNING: SC.exe stop failed for $ServiceName. Exit code: $($scResult.ExitCode)" "WARNING"
                            }
                        } catch {
                            $errorMsg = $_.Exception.Message
                            Write-LogEntry "WARNING: SC.exe stop failed for ${ ServiceName}: $errorMsg" "WARNING"
                        }
                    }

                    if (-not $stopSuccess) {
                        Write-LogEntry "ERROR: Could not stop service $ServiceName after trying multiple methods" "ERROR"
                    }
                } catch {
                    $errorMsg = $_.Exception.Message
                    Write-LogEntry "ERROR: Failed to stop service ${ ServiceName}. Error: $errorMsg" "WARNING"
                }
            }
            # Then, set the service to disabled
            Write-LogEntry "DISABLING: Service $ServiceName" "INFO"

            # Try multiple methods to disable the service
            $disableSuccess = $false

            try {
                # Method 1: Standard Set-Service cmdlet
                try {
                    Set-Service -Name $ServiceName -StartupType Disabled -ErrorAction Stop
                    Write-LogEntry "DISABLED: Service $ServiceName using standard method" "SUCCESS"
                    $disableSuccess = $true
                } catch {
                    $errorMsg = $_.Exception.Message
                    Write-LogEntry "WARNING: Standard disable failed for ${ ServiceName}: $errorMsg. Trying alternative methods." "WARNING"
                }

                # Method 2: Use the registry directly if method 1 failed
                if (-not $disableSuccess) {
                    try {
                        $regPath = "HKLM:\SYSTEM\CurrentControlSet\Services\$ServiceName"
                        if (Test-Path $regPath) {
                            Set-ItemProperty -Path $regPath -Name "Start" -Value 4 -Type DWord -Force
                            Write-LogEntry "DISABLED: Service $ServiceName using registry method" "SUCCESS"
                            $disableSuccess = $true
                        } else {
                            Write-LogEntry "WARNING: Registry key not found for $ServiceName" "WARNING"
                        }
                    } catch {
                        $errorMsg = $_.Exception.Message
                        Write-LogEntry "WARNING: Registry disable failed for ${ ServiceName}: $errorMsg" "WARNING"
                    }
                }

                # Method 3: Use SC.exe command if both other methods failed
                if (-not $disableSuccess) {
                    try {
                        $scResult = Start-Process -FilePath "sc.exe" -ArgumentList "config $ServiceName start = disabled" -NoNewWindow -Wait -PassThru
                        if ($scResult.ExitCode -eq 0) {
                            Write-LogEntry "DISABLED: Service $ServiceName using SC.exe" "SUCCESS"
                            $disableSuccess = $true
                        } else {
                            Write-LogEntry "WARNING: SC.exe disable failed for $ServiceName. Exit code: $($scResult.ExitCode)" "WARNING"
                        }
                    } catch {
                        $errorMsg = $_.Exception.Message
                        Write-LogEntry "WARNING: SC.exe disable failed for ${ ServiceName}: $errorMsg" "WARNING"
                    }
                }

                if (-not $disableSuccess) {
                    Write-LogEntry "ERROR: Could not disable service $ServiceName after trying multiple methods" "ERROR"
                }
            } catch {
                $errorMsg = $_.Exception.Message
                Write-LogEntry "ERROR: Failed to disable service ${ ServiceName}. Error: $errorMsg" "WARNING"
            }
        } else {
            Write-LogEntry "WhatIf: Would stop and disable service: $ServiceName" "INFO"
        }
    } catch {
        $errorMsg = $_.Exception.Message
        Write-LogEntry "ERROR: An error occurred while processing service ${ ServiceName}. Error: $errorMsg" "ERROR"
    }
}

# Start script execution
Write-LogEntry "Starting Windows bloatware removal script v$scriptVersion" "INFO"

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
        "Microsoft.XboxSpeechToTextOverlay", "Microsoft.YourPhone",
        "Microsoft.ZuneMusic",
        "Microsoft.ZuneVideo",
        "Microsoft.Office.OneNote",
        "king.com.CandyCrushSaga",
        "king.com.CandyCrushSodaSaga",
        "king.com.CandyCrushFriends"
    )
    # Remove UWP bloatware
    Write-LogEntry "Removing UWP bloatware applications..." "INFO"
    foreach ($app in $uwpBloatware) {
        # Wrap each call in try/catch to ensure script continues even if one app fails
        try {
            Uninstall-UWPApp -AppName $app -ErrorAction SilentlyContinue
        } catch {
            # Just log and continue to the next app
            Write-LogEntry "Caught exception while processing $app, continuing with next app" "WARNING"
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
        "HP Sure",
        "HP Wolf Security"
        # Lenovo apps are now handled by the Remove-LenovoBloatware function
    )

    # Remove Win32 bloatware
    Write-LogEntry "Removing Win32 bloatware applications..." "INFO"
    foreach ($app in $win32Bloatware) {
        Uninstall-Win32App -DisplayName $app
    }

    # Remove Dell bloatware except Command Update
    Write-LogEntry "Removing Dell bloatware (except Command Update)..." "INFO"
    Remove-DellBloatware

    # Remove Lenovo bloatware except Vantage
    Write-LogEntry "Removing Lenovo bloatware (except Vantage)..." "INFO"
    Remove-LenovoBloatware

    # Remove problematic and unnecessary services
    # List of services to stop and disable
    Write-LogEntry "Stopping and disabling bloatware services..." "INFO"
    $bloatwareServices = @(
        # Service name, Display name (for logs)
        # Using common service names but also listing full display names for better matching
        @{ Name = "ENSS"; DisplayName = "Elliptic Virtual Lock Sensor Service" },
        @{ Name = "EllipticVS"; DisplayName = "Elliptic Virtual Lock Sensor Service" },
        @{ Name = "ICSS"; DisplayName = "Intel Context Sensing Service" },
        @{ Name = "IntelCSS"; DisplayName = "Intel Context Sensing Service" }
    )

    foreach ($service in $bloatwareServices) {
        Stop-DisableBloatwareService -ServiceName $service.Name -DisplayName $service.DisplayName -Force:$Force
    }

    # Clean up any leftover files
    $bloatwareFolders = @(
        "${ env:ProgramFiles}\McAfee",
        "${ env:ProgramFiles}\Norton",
        "${ env:ProgramFiles}\Wild Tangent Games",
        "${ env:ProgramFiles(x86)}\McAfee",
        "${ env:ProgramFiles(x86)}\Norton",
        "${ env:ProgramFiles(x86)}\Wild Tangent Games"
    )

    # Manufacturer folders require special handling
    $manufacturerFolders = @(
        "${ env:ProgramFiles}\Dell",
        "${ env:ProgramFiles(x86)}\Dell",
        "${ env:ProgramFiles}\Lenovo",
        "${ env:ProgramFiles(x86)}\Lenovo",
        "${ env:ProgramFiles}\HP",
        "${ env:ProgramFiles(x86)}\HP"
    )

    Write-LogEntry "Cleaning up leftover bloatware directories..." "INFO"

    # First handle non-manufacturer folders (full removal)
    foreach ($folder in $bloatwareFolders) {
        if (Test-Path $folder) {
            # Use our helper function that guarantees no confirmation prompt
            if ($WhatIfPreference) {
                Write-LogEntry "WhatIf: Would remove directory: $folder" "INFO"
            } else {
                Remove-DirectorySilently -Path $folder -Force
            }
        }
    }

    # Then handle manufacturer folders (selective removal)
    foreach ($folder in $manufacturerFolders) {
        if (Test-Path $folder) {
            # Skip Dell Command Update folders
            if (($folder -like "*Dell*") -and (Test-Path "$folder\Command Update")) {
                Write-LogEntry "Skipping Dell Command Update folder: $folder\Command Update" "INFO"
                continue
            }

            # Skip Lenovo Vantage folders
            if (($folder -like "*Lenovo*") -and (Test-Path "$folder\Lenovo Vantage")) {
                Write-LogEntry "Skipping Lenovo Vantage folder: $folder\Lenovo Vantage" "INFO"
                continue
            }

            # Use our enhanced helper function that handles locked files
            if ($WhatIfPreference) {
                Write-LogEntry "WhatIf: Would selectively clean directory: $folder" "INFO"
            } else {
                Remove-DirectorySilently -Path $folder -LogPrefix "SELECTIVE:" -Force
            }
        }
    }

    Write-LogEntry "Windows bloatware removal completed successfully" "SUCCESS"
} catch {
    $errorMsg = $_.Exception.Message
    Write-LogEntry "An error occurred during bloatware removal: $errorMsg" "ERROR"
    Write-LogEntry "Exception details: $($_.Exception.Message)" "ERROR"
    Write-LogEntry "Stack trace: $($_.ScriptStackTrace)" "DEBUG"
    exit 1
}
