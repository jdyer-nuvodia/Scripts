# =============================================================================
# Script: Get-SetInactivityTimers.ps1
# Created: 2025-04-08 21:45:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-05-09 21:45:00 UTC
# Updated By: GitHub Copilot
# Version: 1.5.0
# Additional Info: Added functionality to detect Intel Context Sensing, Dell Optimizer, Elliptic Virtual Lock Sensor, and Windows Dynamic Lock.
# =============================================================================

<#
.SYNOPSIS
Gets and optionally sets Windows system inactivity timers.

.DESCRIPTION
This script retrieves all available system inactivity settings including screen timeout,
sleep settings, power management configurations, screen saver timeout, and security policies related to machine locking.
It displays these settings to the user and provides the option to modify them. All changes support -WhatIf functionality for safety.

.EXAMPLE
.\\Get-SetInactivityTimers.ps1
Displays current inactivity settings and prompts for changes

.EXAMPLE
.\\Get-SetInactivityTimers.ps1 -WhatIf
Shows what changes would be made without actually making them
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param()

# Function to safely stop transcript
function Stop-TranscriptSafely {
    Write-Debug "Entering Stop-TranscriptSafely function."
    # Check the script-scoped flag to see if transcript was started by this script
    if ($script:transcriptActive) {
        Write-Debug "Transcript was active, attempting to stop."
        try {
            # Use Stop to ensure catch block executes on error
                        Stop-Transcript -ErrorAction Stop
            Write-Debug "Stop-Transcript command executed."
            # Give the system a moment to release the file handle
            Start-Sleep -Milliseconds 500
            Write-Debug "Slept for 500ms after Stop-Transcript."
            # Force garbage collection to potentially release file handles
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            Write-Debug "Garbage collection triggered after Stop-Transcript."
            # Set the flag to inactive *after* successful stop
            $script:transcriptActive = $false
            Write-Debug "Transcript marked as inactive."
        }
        catch {
            Write-Warning "Error stopping transcript: $_"
            # Even if stopping failed, mark as inactive to prevent retry loops if applicable
            $script:transcriptActive = $false
            Write-Debug "Transcript marked as inactive despite error during stop."
        }
    } else {
        Write-Debug "Transcript was not marked as active by this script, skipping Stop-Transcript."
    }
    Write-Debug "Exiting Stop-TranscriptSafely function."
}

# Function to format minutes into a readable string
function Format-Minutes {
    # Use double for precision
        param([double]$TotalMinutes)

    if ($TotalMinutes -eq 0) { return "Never" }

    # Handle large values (over 168 hours / 7 days)
    # 168 hours * 60 minutes/hour
        if ($TotalMinutes -gt 10080) {
         return "Effectively Disabled" # Changed from "Effectively Never"
    }

    # Use integer part for calculations involving days, hours, minutes
    $minutesInt = [int][math]::Floor($TotalMinutes)

    if ($minutesInt -lt 1) {
        # Handle cases less than 1 minute if needed, e.g., show seconds or round up
        # Or adjust as needed
                return "Less than 1 minute"
    }

    $days = [math]::Floor($minutesInt / 1440)
    $remainingMinutesAfterDays = $minutesInt % 1440
    $hours = [math]::Floor($remainingMinutesAfterDays / 60)
    $minutes = $remainingMinutesAfterDays % 60

    $parts = @()
    if ($days -gt 0) { $parts += "$days day$(if ($days -gt 1) {'s'} else {''})" }
    if ($hours -gt 0) { $parts += "$hours hour$(if ($hours -gt 1) {'s'} else {''})" }
    if ($minutes -gt 0) { $parts += "$minutes minute$(if ($minutes -gt 1) {'s'} else {''})" }

    # Fallback if calculation results in empty parts (should not happen with Floor)
    if ($parts.Count -eq 0) {
        return "$minutesInt minute$(if ($minutesInt -gt 1) {'s'} else {''})"
    }

    return $parts -join ', '
}

function Get-PowerSettings {
    Write-Host "Retrieving current power settings..." -ForegroundColor Cyan
    Write-Debug "Starting power settings retrieval"

    # Get current power scheme info
    $powerSchemeInfo = powercfg /getactivescheme
    Write-Debug "Raw power scheme info: $powerSchemeInfo"

    if ([string]::IsNullOrWhiteSpace($powerSchemeInfo)) {
        Write-Warning "Could not retrieve power scheme information"
        return $null
    }
    # Extract the GUID robustly
    $schemeGuid = ($powerSchemeInfo -split ' ' | Where-Object { $_ -match '^[{(]?[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}[)}]?$' } | Select-Object -First 1)
    # Extract the Scheme Name
    $schemeName = if ($powerSchemeInfo -match '\((.*?)\)') { $matches[1] } else { "Unknown" }
    Write-Debug "Active Power Scheme Name: $schemeName"

    if (-not $schemeGuid) {
        Write-Warning "Could not determine active power scheme GUID from '$powerSchemeInfo'"
        return $null
    }
    Write-Debug "Active Power Scheme GUID: $schemeGuid"

    # Get all power settings using powercfg /query
    $powerSettings = powercfg /query $schemeGuid
    if ([string]::IsNullOrWhiteSpace($powerSettings)) {
        Write-Warning "Could not retrieve power settings for scheme $schemeGuid"
        return $null
    }

    Write-Debug "Raw power settings output:"
    $powerSettings | Out-String | Write-Debug

    # Initialize timeout variables (in seconds)
    $monitorTimeoutAC_Seconds = 0
    $monitorTimeoutDC_Seconds = 0
    $sleepTimeoutAC_Seconds = 0
    $sleepTimeoutDC_Seconds = 0
    $hibernateTimeoutAC_Seconds = 0
    $hibernateTimeoutDC_Seconds = 0
    # Added for screen saver
        $screenSaverTimeout_Seconds = 0

    # Split the output into lines for parsing
    $powerSettingsLines = $powerSettings -split [System.Environment]::NewLine

    # Define GUIDs for common settings (these might vary slightly by Windows version, but are generally stable)
    # SUB_VIDEO
        $displaySubgroupGuid = "7516b95f-f776-4464-8c53-06167f40cc99"
    # SUB_SLEEP
        $sleepSubgroupGuid = "238c9fa8-0aad-41ed-83f4-97be242c8f20"
    # VIDEOIDLE (Corrected last digit)
        $displayTimeoutSettingGuid = "3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e"
    # STANDBYIDLE
        $sleepTimeoutSettingGuid = "29f6c1db-86da-48c5-9fdb-f2b67b1f44da"
    # HIBERNATEIDLE
        $hibernateTimeoutSettingGuid = "9d7815a6-7ee4-497e-8888-515a05f02364"

    $currentSubgroupGuid = $null
    $currentSettingGuid = $null

    foreach ($line in $powerSettingsLines) {
        # Identify current Subgroup
        if ($line -match "Subgroup GUID: ([0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12})") {
            $currentSubgroupGuid = $matches[1]
            # Reset setting when subgroup changes
                        $currentSettingGuid = $null
            Write-Debug "Processing Subgroup: $currentSubgroupGuid"
            continue
        }

        # Identify current Power Setting GUID within the subgroup
        if ($line -match "Power Setting GUID: ([0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12})") {
            $currentSettingGuid = $matches[1]
            Write-Debug "Processing Setting: $currentSettingGuid within Subgroup: $currentSubgroupGuid"
            continue
        }

        # Extract AC value if within the correct subgroup and setting
        if ($line -match "Current AC Power Setting Index:\s+(0x[0-9a-fA-F]+)") {
            $hexValue = $matches[1]
            $seconds = [Convert]::ToInt32($hexValue, 16)
            Write-Debug "Found AC Value (Hex: $hexValue, Seconds: $seconds) for Setting: $currentSettingGuid"
            if ($currentSubgroupGuid -eq $displaySubgroupGuid -and $currentSettingGuid -eq $displayTimeoutSettingGuid) {
                $monitorTimeoutAC_Seconds = $seconds
                Write-Debug "Assigned AC Monitor Timeout: $seconds seconds"
            }
            elseif ($currentSubgroupGuid -eq $sleepSubgroupGuid -and $currentSettingGuid -eq $sleepTimeoutSettingGuid) {
                $sleepTimeoutAC_Seconds = $seconds
                Write-Debug "Assigned AC Sleep Timeout: $seconds seconds"
            }
            elseif ($currentSubgroupGuid -eq $sleepSubgroupGuid -and $currentSettingGuid -eq $hibernateTimeoutSettingGuid) {
                $hibernateTimeoutAC_Seconds = $seconds
                Write-Debug "Assigned AC Hibernate Timeout: $seconds seconds"
            }
        }

        # Extract DC value if within the correct subgroup and setting
        if ($line -match "Current DC Power Setting Index:\s+(0x[0-9a-fA-F]+)") {
            $hexValue = $matches[1]
            $seconds = [Convert]::ToInt32($hexValue, 16)
            Write-Debug "Found DC Value (Hex: $hexValue, Seconds: $seconds) for Setting: $currentSettingGuid"
            if ($currentSubgroupGuid -eq $displaySubgroupGuid -and $currentSettingGuid -eq $displayTimeoutSettingGuid) {
                $monitorTimeoutDC_Seconds = $seconds
                Write-Debug "Assigned DC Monitor Timeout: $seconds seconds"
            }
            elseif ($currentSubgroupGuid -eq $sleepSubgroupGuid -and $currentSettingGuid -eq $sleepTimeoutSettingGuid) {
                $sleepTimeoutDC_Seconds = $seconds
                Write-Debug "Assigned DC Sleep Timeout: $seconds seconds"
            }
            elseif ($currentSubgroupGuid -eq $sleepSubgroupGuid -and $currentSettingGuid -eq $hibernateTimeoutSettingGuid) {
                $hibernateTimeoutDC_Seconds = $seconds
                Write-Debug "Assigned DC Hibernate Timeout: $seconds seconds"
            }
        }
    }

    # Get screen saver settings from registry
    try {
        $screenSaverTimeoutReg = Get-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "ScreenSaveTimeout" -ErrorAction SilentlyContinue
        if ($null -ne $screenSaverTimeoutReg.ScreenSaveTimeout) {
            $screenSaverTimeout_Seconds = [int]$screenSaverTimeoutReg.ScreenSaveTimeout
            Write-Debug "Found Screen Saver Timeout: $screenSaverTimeout_Seconds seconds"
        } else {
            Write-Debug "Screen Saver Timeout registry key not found or value is null."
        }
    } catch {
        Write-Warning "Could not retrieve screen saver timeout from registry: $_"
    }

    # Convert seconds to minutes for display (0 seconds means Never)
    # Use 60.0 for double division
        $monitorTimeoutAC_Minutes = if ($monitorTimeoutAC_Seconds -eq 0) { 0 } else { $monitorTimeoutAC_Seconds / 60.0 }
    $monitorTimeoutDC_Minutes = if ($monitorTimeoutDC_Seconds -eq 0) { 0 } else { $monitorTimeoutDC_Seconds / 60.0 }
    $sleepTimeoutAC_Minutes = if ($sleepTimeoutAC_Seconds -eq 0) { 0 } else { $sleepTimeoutAC_Seconds / 60.0 }
    $sleepTimeoutDC_Minutes = if ($sleepTimeoutDC_Seconds -eq 0) { 0 } else { $sleepTimeoutDC_Seconds / 60.0 }
    $hibernateTimeoutAC_Minutes = if ($hibernateTimeoutAC_Seconds -eq 0) { 0 } else { $hibernateTimeoutAC_Seconds / 60.0 }
    $hibernateTimeoutDC_Minutes = if ($hibernateTimeoutDC_Seconds -eq 0) { 0 } else { $hibernateTimeoutDC_Seconds / 60.0 }
    $screenSaverTimeout_Minutes = if ($screenSaverTimeout_Seconds -eq 0) { 0 } else { $screenSaverTimeout_Seconds / 60.0 }

    Write-Debug "Calculated Minutes - Monitor AC: $monitorTimeoutAC_Minutes, DC: $monitorTimeoutDC_Minutes"
    Write-Debug "Calculated Minutes - Sleep AC: $sleepTimeoutAC_Minutes, DC: $sleepTimeoutDC_Minutes"
    Write-Debug "Calculated Minutes - Hibernate AC: $hibernateTimeoutAC_Minutes, DC: $hibernateTimeoutDC_Minutes"
    Write-Debug "Calculated Minutes - Screen Saver: $screenSaverTimeout_Minutes"

    # Return results as a hashtable
    return @{
        # Use extracted scheme name
                PowerPlanName = $schemeName
        PowerPlanGuid = $schemeGuid
        MonitorTimeoutAC = $monitorTimeoutAC_Minutes
        MonitorTimeoutDC = $monitorTimeoutDC_Minutes
        SleepTimeoutAC = $sleepTimeoutAC_Minutes
        SleepTimeoutDC = $sleepTimeoutDC_Minutes
        HibernateTimeoutAC = $hibernateTimeoutAC_Minutes
        HibernateTimeoutDC = $hibernateTimeoutDC_Minutes
        ScreenSaverTimeout = $screenSaverTimeout_Minutes
    }
}

function Get-LockPolicySettings {    Write-Host "Checking Group Policy and security settings..." -ForegroundColor Cyan
    $settings = [PSCustomObject]@{
        PSTypeName = 'LockPolicySettings'
        ScreenSaverForced = $false
        ScreenSaverSecure = $false
        AutoLockEnabled = $false
        AutoLockTimeout = $null
    }

    try {
        # Check screen saver policy settings
        $screenSaverPolicy = Get-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Windows\Control Panel\Desktop" -ErrorAction SilentlyContinue
        if ($screenSaverPolicy -and $null -ne $screenSaverPolicy.ScreenSaverIsSecure) {
            $settings.ScreenSaverForced = $screenSaverPolicy.ScreenSaverIsSecure -eq 1
        }


        # Check workstation lock settings
        $lockSettings = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\System" -ErrorAction SilentlyContinue
        if ($lockSettings) {
            $settings.AutoLockEnabled = $lockSettings.DisableLockWorkstation -ne 1
        }

        # Check machine inactivity limit
        $inactivityLimit = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "InactivityTimeoutSecs" -ErrorAction SilentlyContinue
        if ($inactivityLimit) {
            $settings.AutoLockTimeout = [math]::Round($inactivityLimit.InactivityTimeoutSecs / 60)
        }
    }
    catch {
        Write-Warning "Error checking security policies: $_"
    }

    return $settings
}

function Get-IntelContextSensing {
    Write-Debug "Checking for Intel Context Sensing..."
    $contextSensingInfo = @{
        Installed = $false
        Version = "Not installed"
        Path = $null
        Status = "Not present"
    }

    try {
        # Check for Intel Context Sensing service
        $service = Get-Service -Name "IntelSenseS*" -ErrorAction SilentlyContinue |
                   Where-Object { $_.DisplayName -like "*Context*" -or $_.DisplayName -like "*Sens*" }
        if ($service) {
            $contextSensingInfo.Installed = $true
            $contextSensingInfo.Status = $service.Status

            # Try to get version info
            $path = (Get-WmiObject -Class Win32_Service -Filter "Name='$($service.Name)'" -ErrorAction SilentlyContinue).PathName
            if ($path) {
                $contextSensingInfo.Path = $path -replace '^"([^"]+)".*$', '$1'
                if (Test-Path $contextSensingInfo.Path) {
                    $versionInfo = (Get-Item $contextSensingInfo.Path).VersionInfo
                    $contextSensingInfo.Version = $versionInfo.ProductVersion
                }
            }
        }

        # Alternative check for Intel Context Sensing through registry
        if (-not $contextSensingInfo.Installed) {
            $regPaths = @(
                "HKLM:\SOFTWARE\Intel\Intel Context Sensing",
                "HKLM:\SOFTWARE\Intel\Sensing",
                "HKLM:\SOFTWARE\Intel\HPSS"
            )

            foreach ($regPath in $regPaths) {
                if (Test-Path $regPath) {
                    $contextSensingInfo.Installed = $true
                    $contextSensingInfo.Status = "Installed (service not running or not detected)"
                    break
                }
            }
        }

        # Check for installed applications
        if (-not $contextSensingInfo.Installed) {
            $apps = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |
                    Where-Object { $_.DisplayName -like "*Intel*" -and ($_.DisplayName -like "*Context*" -or $_.DisplayName -like "*Sens*") }
            if ($apps) {
                $contextSensingInfo.Installed = $true
                $contextSensingInfo.Status = "Application installed"
                $contextSensingInfo.Version = $apps[0].DisplayVersion
            }
        }
    }
    catch {
        Write-Debug "Error checking for Intel Context Sensing: $_"
    }

    return $contextSensingInfo
}

function Get-DellOptimizer {
    Write-Debug "Checking for Dell Optimizer..."
    $dellOptimizerInfo = @{
        Installed = $false
        Version = "Not installed"
        Features = @()
        Status = "Not present"
    }

    try {
        # Check for Dell Optimizer application
        $appPaths = @(
            "${env:ProgramFiles}\Dell\DellOptimizer\DellOptimizer.exe",
            "${env:ProgramFiles(x86)}\Dell\DellOptimizer\DellOptimizer.exe"
        )

        foreach ($path in $appPaths) {
            if (Test-Path $path) {
                $dellOptimizerInfo.Installed = $true
                $versionInfo = (Get-Item $path).VersionInfo
                $dellOptimizerInfo.Version = $versionInfo.ProductVersion
                $dellOptimizerInfo.Status = "Installed"
                break
            }
        }

        # Check for Dell Optimizer service
        if (-not $dellOptimizerInfo.Installed) {
            $service = Get-Service -Name "*DellOptimizer*" -ErrorAction SilentlyContinue
            if ($service) {
                $dellOptimizerInfo.Installed = $true
                $dellOptimizerInfo.Status = $service.Status
            }
        }

        # Check registry for Dell Optimizer features
        $regPath = "HKLM:\SOFTWARE\Dell\DellOptimizer"
        if (Test-Path $regPath) {
            $dellOptimizerInfo.Installed = $true

            # Check for specific features
            $featurePaths = @{
                "Presence Detection" = "HKLM:\SOFTWARE\Dell\DellOptimizer\PresenceDetection"
                "Walk Away Lock" = "HKLM:\SOFTWARE\Dell\DellOptimizer\WalkAwayLock"
                "Wake on Approach" = "HKLM:\SOFTWARE\Dell\DellOptimizer\WakeOnApproach"
                "Intelligent Audio" = "HKLM:\SOFTWARE\Dell\DellOptimizer\IntelligentAudio"
            }

            foreach ($feature in $featurePaths.Keys) {
                if (Test-Path $featurePaths[$feature]) {
                    try {
                        $enabled = Get-ItemProperty -Path $featurePaths[$feature] -Name "Enabled" -ErrorAction SilentlyContinue
                        if ($null -ne $enabled -and $enabled.Enabled -eq 1) {
                            $dellOptimizerInfo.Features += "$feature (Enabled)"
                        }
                        else {
                            $dellOptimizerInfo.Features += "$feature (Disabled)"
                        }
                    }
                    catch {
                        $dellOptimizerInfo.Features += "$feature (Status Unknown)"
                    }
                }
            }

            # If features are empty but Dell Optimizer is installed
            if ($dellOptimizerInfo.Features.Count -eq 0) {
                $dellOptimizerInfo.Features = @("No auto-lock features detected")
            }
        }

        # Check installed applications as fallback
        if (-not $dellOptimizerInfo.Installed) {
            $apps = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |
                    Where-Object { $_.DisplayName -like "*Dell Optimizer*" }
            if ($apps) {
                $dellOptimizerInfo.Installed = $true
                $dellOptimizerInfo.Status = "Application installed"
                $dellOptimizerInfo.Version = $apps[0].DisplayVersion
            }
        }
    }
    catch {
        Write-Debug "Error checking for Dell Optimizer: $_"
    }

    return $dellOptimizerInfo
}

function Get-EllipticSensor {
    Write-Debug "Checking for Elliptic Virtual Lock Sensor..."
    $ellipticInfo = @{
        Installed = $false
        Version = "Not installed"
        Status = "Not present"
    }

    try {
        # Check for Elliptic Sensor service
        $service = Get-Service -Name "*Elliptic*" -ErrorAction SilentlyContinue
        if ($service) {
            $ellipticInfo.Installed = $true
            $ellipticInfo.Status = $service.Status

            # Try to get version info
            $path = (Get-WmiObject -Class Win32_Service -Filter "Name='$($service.Name)'" -ErrorAction SilentlyContinue).PathName
            if ($path) {
                $ellipticInfo.Path = $path -replace '^"([^"]+)".*$', '$1'
                if (Test-Path $ellipticInfo.Path) {
                    $versionInfo = (Get-Item $ellipticInfo.Path).VersionInfo
                    $ellipticInfo.Version = $versionInfo.ProductVersion
                }
            }
        }

        # Check registry for Elliptic Sensor
        $regPaths = @(
            "HKLM:\SOFTWARE\Elliptic",
            "HKLM:\SOFTWARE\Elliptic Labs"
        )

        foreach ($regPath in $regPaths) {
            if (Test-Path $regPath) {
                $ellipticInfo.Installed = $true
                if ($ellipticInfo.Status -eq "Not present") {
                    $ellipticInfo.Status = "Installed (service not running or not detected)"
                }
                break
            }
        }

        # Check for device presence in device manager
        $deviceInfo = Get-WmiObject -Class Win32_PnPEntity -ErrorAction SilentlyContinue |
                      Where-Object { $_.Name -like "*Elliptic*" -or $_.Name -like "*Sensor*" -and $_.Name -like "*Lock*" }
        if ($deviceInfo) {
            $ellipticInfo.Installed = $true
            $ellipticInfo.Status = "Device detected"
        }

        # Check installed applications
        if (-not $ellipticInfo.Installed) {
            $apps = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |
                    Where-Object { $_.DisplayName -like "*Elliptic*" }
            if ($apps) {
                $ellipticInfo.Installed = $true
                $ellipticInfo.Status = "Application installed"
                $ellipticInfo.Version = $apps[0].DisplayVersion
            }
        }
    }
    catch {
        Write-Debug "Error checking for Elliptic Virtual Lock Sensor: $_"
    }

    return $ellipticInfo
}

function Get-WindowsDynamicLock {
    Write-Debug "Checking for Windows Dynamic Lock..."
    $dynamicLockInfo = @{
        Enabled = $false
        Status = "Not enabled"
        RequiresBluetoothDevice = $true
    }

    try {
        # Check registry for Dynamic Lock setting
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
        $settingPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"

        # Check if Dynamic Lock is enabled
        $enabledSetting = Get-ItemProperty -Path $regPath -Name "EnableDynamicLock" -ErrorAction SilentlyContinue
        if ($null -ne $enabledSetting -and $enabledSetting.EnableDynamicLock -eq 1) {
            $dynamicLockInfo.Enabled = $true
            $dynamicLockInfo.Status = "Enabled"
        }

        # Check alternative path if not found in policies
        if (-not $dynamicLockInfo.Enabled) {
            $altSetting = Get-ItemProperty -Path $settingPath -Name "EnableDynamicLock" -ErrorAction SilentlyContinue
            if ($null -ne $altSetting -and $altSetting.EnableDynamicLock -eq 1) {
                $dynamicLockInfo.Enabled = $true
                $dynamicLockInfo.Status = "Enabled"
            }
        }

        # Check via Get-ItemProperty HKCU instead
        if (-not $dynamicLockInfo.Enabled) {
            $userSetting = Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "EnableDynamicLock" -ErrorAction SilentlyContinue
            if ($null -ne $userSetting -and $userSetting.EnableDynamicLock -eq 1) {
                $dynamicLockInfo.Enabled = $true
                $dynamicLockInfo.Status = "Enabled"
            }
        }

        # Check if Dynamic Lock is available via SigninOptions
        if (-not $dynamicLockInfo.Enabled) {
            $signInOptions = Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\SignInOptions" -Name "DynamicLockEnabled" -ErrorAction SilentlyContinue
            if ($null -ne $signInOptions -and $signInOptions.DynamicLockEnabled -eq 1) {
                $dynamicLockInfo.Enabled = $true
                $dynamicLockInfo.Status = "Enabled via Sign-in Options"
            }
        }

        # Check Bluetooth availability for Dynamic Lock
        $bluetoothAvailable = $false
        try {
            $bluetoothDevices = Get-PnpDevice -Class Bluetooth -ErrorAction SilentlyContinue
            if ($bluetoothDevices) {
                $connectedDevices = $bluetoothDevices | Where-Object { $_.Status -eq "OK" }
                if ($connectedDevices) {
                    $bluetoothAvailable = $true
                    if ($dynamicLockInfo.Enabled) {
                        $dynamicLockInfo.Status += " with Bluetooth devices connected"
                    }
                }
            }

            if (-not $bluetoothAvailable -and $dynamicLockInfo.Enabled) {
                $dynamicLockInfo.Status += " (Warning: No Bluetooth devices detected)"
            }
        }
        catch {
            Write-Debug "Error checking Bluetooth devices: $_"
            $dynamicLockInfo.RequiresBluetoothDevice = "Unknown (Error checking devices)"
        }
    }
    catch {
        Write-Debug "Error checking for Windows Dynamic Lock: $_"
    }

    return $dynamicLockInfo
}

function Get-AutoLockCapabilities {
    Write-Host "Checking for auto-lock capable hardware and software..." -ForegroundColor Cyan

    $intelContextSensing = Get-IntelContextSensing
    $dellOptimizer = Get-DellOptimizer
    $ellipticSensor = Get-EllipticSensor
    $dynamicLock = Get-WindowsDynamicLock

    $autoLockCapabilities = [PSCustomObject]@{
        IntelContextSensing = $intelContextSensing
        DellOptimizer = $dellOptimizer
        EllipticSensor = $ellipticSensor
        WindowsDynamicLock = $dynamicLock
        SupportedAutoLockMethods = @()
    }

    # Build list of supported methods
    if ($intelContextSensing.Installed) {
        $autoLockCapabilities.SupportedAutoLockMethods += "Intel Context Sensing"
    }

    if ($dellOptimizer.Installed -and $dellOptimizer.Features -match "Walk Away Lock|Presence Detection") {
        $autoLockCapabilities.SupportedAutoLockMethods += "Dell Optimizer"
    }

    if ($ellipticSensor.Installed) {
        $autoLockCapabilities.SupportedAutoLockMethods += "Elliptic Sensor"
    }

    if ($dynamicLock.Enabled) {
        $autoLockCapabilities.SupportedAutoLockMethods += "Windows Dynamic Lock"
    }

    return $autoLockCapabilities
}

function Set-PowerTimeout {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [int]$MonitorTimeoutAC,
        [Parameter()]
        [int]$MonitorTimeoutDC,
        [Parameter()]
        [int]$SleepTimeoutAC,
        [Parameter()]
        [int]$SleepTimeoutDC,
        [Parameter()]
        [int]$ScreenSaverTimeout
    )

    if ($PSCmdlet.ShouldProcess("Power Settings", "Update inactivity timeouts")) {
        try {
            # Get current power scheme GUID
            $schemeGuid = (powercfg /getactivescheme) -split " " | Where-Object { $_ -match '^[{]?[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}[}]?$' }
            if (-not $schemeGuid) {
                Write-Warning "Could not determine active power scheme GUID"
                return $false
            }
              # Set monitor timeout (AC and DC)
            if ($PSBoundParameters.ContainsKey('MonitorTimeoutAC')) {
                Write-Host "Setting AC monitor timeout..." -ForegroundColor Cyan
                powercfg /change monitor-timeout-ac $MonitorTimeoutAC
            }
            if ($PSBoundParameters.ContainsKey('MonitorTimeoutDC')) {
                Write-Host "Setting DC monitor timeout..." -ForegroundColor Cyan
                powercfg /change monitor-timeout-dc $MonitorTimeoutDC
            }

            # Set sleep timeout (AC and DC)
            if ($PSBoundParameters.ContainsKey('SleepTimeoutAC')) {
                Write-Host "Setting AC sleep timeout..." -ForegroundColor Cyan
                powercfg /change standby-timeout-ac $SleepTimeoutAC
            }
            if ($PSBoundParameters.ContainsKey('SleepTimeoutDC')) {
                Write-Host "Setting DC sleep timeout..." -ForegroundColor Cyan
                powercfg /change standby-timeout-dc $SleepTimeoutDC
            }

            # Set screen saver timeout
            if ($PSBoundParameters.ContainsKey('ScreenSaverTimeout')) {
                Write-Host "Setting screen saver timeout..." -ForegroundColor Cyan
                $timeoutSeconds = $ScreenSaverTimeout * 60
                Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "ScreenSaveTimeout" -Value $timeoutSeconds
            }

            Write-Host "All inactivity timers have been updated successfully!" -ForegroundColor Green
        }
        catch {
            Write-host "Error setting power settings: $_" -ForegroundColor Red
            return $false
        }
    }
    return $true
}

function Set-LockPolicySettings {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [bool]$ScreenSaverForced,
        [Parameter()]
        [bool]$AutoLockEnabled,
        [Parameter()]
        [int]$AutoLockTimeout
    )

    if ($PSCmdlet.ShouldProcess("Group Policy Settings", "Update lock policy settings")) {
        try {
            # Set screen saver security enforcement
            if ($PSBoundParameters.ContainsKey('ScreenSaverForced')) {
                Write-Host "Setting screen saver security enforcement..." -ForegroundColor Cyan
                $regPath = "HKCU:\Software\Policies\Microsoft\Windows\Control Panel\Desktop"
                if (!(Test-Path $regPath)) {
                    New-Item -Path $regPath -Force | Out-Null
                }
                Set-ItemProperty -Path $regPath -Name "ScreenSaverIsSecure" -Value ([int]$ScreenSaverForced) -Type DWord
            }

            # Set auto lock settings
            if ($PSBoundParameters.ContainsKey('AutoLockEnabled')) {
                Write-host "Setting auto lock enabled status..." -ForegroundColor Cyan
                $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\System"
                if (!(Test-Path $regPath)) {
                    New-Item -Path $regPath -Force | Out-Null
                }
                Set-ItemProperty -Path $regPath -Name "DisableLockWorkstation" -Value ([int](!$AutoLockEnabled)) -Type DWord
            }

            # Set auto lock timeout
            if ($PSBoundParameters.ContainsKey('AutoLockTimeout')) {
                Write-Host "Setting auto lock timeout..." -ForegroundColor Cyan
                $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
                if (!(Test-Path $regPath)) {
                    New-Item -Path $regPath -Force | Out-Null
                }
                Set-ItemProperty -Path $regPath -Name "InactivityTimeoutSecs" -Value ($AutoLockTimeout * 60) -Type DWord
            }

            Write-Host "Group Policy settings have been updated successfully!" -ForegroundColor Green
            return $true
        }
        catch {
            Write-host "Error setting Group Policy settings: $_" -ForegroundColor Red
            return $false
        }
    }
    return $true
}

# Script scope variable to track transcript status
$script:transcriptActive = $false
$script:logPath = $null

# Function to start transcript safely
function Start-TranscriptSafely {
    if ($DebugPreference -ne 'SilentlyContinue' -and -not $script:transcriptActive) {
        try {
            $script:logPath = Join-Path $PSScriptRoot "Get-SetInactivityTimers_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
            Start-Transcript -Path $script:logPath -Force
            $script:transcriptActive = $true
            Write-Debug "Debug logging started. Log file: $script:logPath"
        }
        catch {
            Write-Warning "Failed to start transcript: $_"
        }
    }
}

# Register cleanup for unexpected termination
$null = Register-EngineEvent -SourceIdentifier ([System.Management.Automation.PsEngineEvent]::Exiting) -Action {
    if ($script:transcriptActive) {
        Stop-TranscriptSafely
    }
}

# Main script execution
try {
    Start-TranscriptSafely

    # Display WhatIf mode disclaimer if applicable
    if ($WhatIfPreference) {
        Write-Host "`n[WhatIf Mode] This script is running in simulation mode. No actual changes will be made.`n" -ForegroundColor Yellow
    }

    # Get current settings
    $currentSettings = Get-PowerSettings
    $lockSettings = Get-LockPolicySettings
    $autoLockCapabilities = Get-AutoLockCapabilities

    # Display current settings
    Write-Host "`nPower Plan Information:" -ForegroundColor White
    Write-Host "---------------------" -ForegroundColor White
    # Check if Get-PowerSettings returned valid data
    if ($null -ne $currentSettings) {
        # Ensure PowerPlanName exists
                Write-Host ("Active Power Plan: {0}" -f $currentSettings.PowerPlanName)
        Write-Host ("Plan GUID: {0}" -f $currentSettings.PowerPlanGuid)

        Write-Host "`nCurrent Inactivity Settings:" -ForegroundColor White
        Write-Host "------------------------" -ForegroundColor White
        Write-Host "`nDisplay Settings:" -ForegroundColor Cyan
        Write-Host ("Monitor Timeout (AC Power): {0}" -f (Format-Minutes $currentSettings.MonitorTimeoutAC))
        Write-Host ("Monitor Timeout (Battery): {0}" -f (Format-Minutes $currentSettings.MonitorTimeoutDC))

        Write-Host "`nSleep Settings:" -ForegroundColor Cyan
        Write-Host ("Sleep Timer (AC Power): {0}" -f (Format-Minutes $currentSettings.SleepTimeoutAC))
        Write-Host ("Sleep Timer (Battery): {0}" -f (Format-Minutes $currentSettings.SleepTimeoutDC))
        Write-Host ("Hibernate Timer (AC Power): {0}" -f (Format-Minutes $currentSettings.HibernateTimeoutAC))
        Write-Host ("Hibernate Timer (Battery): {0}" -f (Format-Minutes $currentSettings.HibernateTimeoutDC))

        Write-Host "`nScreen Saver:" -ForegroundColor Cyan
        Write-Host ("Screen Saver Timeout: {0}" -f (Format-Minutes $currentSettings.ScreenSaverTimeout))
    } else {
         Write-Warning "Could not display power settings as they failed to load."
    }

    # Display Lock Settings
    if ($null -ne $lockSettings) {
        Write-Host "`nSecurity and Group Policy Settings:" -ForegroundColor White
        Write-Host "--------------------------------" -ForegroundColor White
        Write-Host ("Screen Saver Security Enforced (User cannot remove password requirement when returning from Screen Saver): {0}" -f $(if ($lockSettings.ScreenSaverForced) { "Yes" } else { "No" }))
        Write-Host ("Auto Lock Enabled: {0}" -f $(if ($lockSettings.AutoLockEnabled) { "Yes" } else { "No" }))
        if ($null -ne $lockSettings.AutoLockTimeout) {
            Write-Host ("Auto Lock Timeout: {0}" -f (Format-Minutes $lockSettings.AutoLockTimeout))
        }
    } else {
        Write-Warning "Could not display lock policy settings as they failed to load."
    }

    # Display Auto-Lock Capabilities
    if ($null -ne $autoLockCapabilities) {
        Write-Host "`nAuto-Lock Capable Hardware and Software:" -ForegroundColor White
        Write-Host "-------------------------------------" -ForegroundColor White

        # Intel Context Sensing
        Write-Host "`nIntel Context Sensing:" -ForegroundColor Cyan
        Write-Host ("Installed: {0}" -f $(if ($autoLockCapabilities.IntelContextSensing.Installed) { "Yes" } else { "No" }))
        if ($autoLockCapabilities.IntelContextSensing.Installed) {
            Write-Host ("Status: {0}" -f $autoLockCapabilities.IntelContextSensing.Status)
            Write-Host ("Version: {0}" -f $autoLockCapabilities.IntelContextSensing.Version)
        }

        # Dell Optimizer
        Write-Host "`nDell Optimizer:" -ForegroundColor Cyan
        Write-Host ("Installed: {0}" -f $(if ($autoLockCapabilities.DellOptimizer.Installed) { "Yes" } else { "No" }))
        if ($autoLockCapabilities.DellOptimizer.Installed) {
            Write-Host ("Status: {0}" -f $autoLockCapabilities.DellOptimizer.Status)
            Write-Host ("Version: {0}" -f $autoLockCapabilities.DellOptimizer.Version)
            Write-Host "Features:"
            foreach ($feature in $autoLockCapabilities.DellOptimizer.Features) {
                Write-Host ("  - {0}" -f $feature)
            }
        }

        # Elliptic Sensor
        Write-Host "`nElliptic Virtual Lock Sensor:" -ForegroundColor Cyan
        Write-Host ("Installed: {0}" -f $(if ($autoLockCapabilities.EllipticSensor.Installed) { "Yes" } else { "No" }))
        if ($autoLockCapabilities.EllipticSensor.Installed) {
            Write-Host ("Status: {0}" -f $autoLockCapabilities.EllipticSensor.Status)
            Write-Host ("Version: {0}" -f $autoLockCapabilities.EllipticSensor.Version)
        }

        # Windows Dynamic Lock
        Write-Host "`nWindows Dynamic Lock:" -ForegroundColor Cyan
        Write-Host ("Enabled: {0}" -f $(if ($autoLockCapabilities.WindowsDynamicLock.Enabled) { "Yes" } else { "No" }))
        if ($autoLockCapabilities.WindowsDynamicLock.Enabled) {
            Write-Host ("Status: {0}" -f $autoLockCapabilities.WindowsDynamicLock.Status)
        }

        # Summary of supported methods
        Write-Host "`nSupported Auto-Lock Methods:" -ForegroundColor White
        if ($autoLockCapabilities.SupportedAutoLockMethods.Count -gt 0) {
            foreach ($method in $autoLockCapabilities.SupportedAutoLockMethods) {
                Write-Host ("  - {0}" -f $method) -ForegroundColor Green
            }
        } else {
            Write-Host "  No auto-lock methods detected on this system" -ForegroundColor Yellow
        }
    }

    # Ask if user wants to change settings (only if settings loaded)
    if ($null -ne $currentSettings -and $null -ne $lockSettings) {
        $response = Read-Host "`nWould you like to change these settings? (Y/N)"

        if ($response -ne "Y") {
            Write-Debug "User chose not to make changes. Preparing to exit."
        }
        else {
            # If continuing, prepare power settings params
            $powerParams = @{}

            # Monitor timeout AC
            $userInput = Read-Host "Enter new Monitor Timeout for AC power (current: $(Format-Minutes $currentSettings.MonitorTimeoutAC)) [Enter to skip]"
            if ($userInput -match '^\d+$') { $powerParams['MonitorTimeoutAC'] = [int]$userInput }

            # Monitor timeout DC
            $userInput = Read-Host "Enter new Monitor Timeout for Battery (current: $(Format-Minutes $currentSettings.MonitorTimeoutDC)) [Enter to skip]"
            if ($userInput -match '^\d+$') { $powerParams['MonitorTimeoutDC'] = [int]$userInput }

            # Sleep timeout AC
            $userInput = Read-Host "Enter new Sleep Timeout for AC power (current: $(Format-Minutes $currentSettings.SleepTimeoutAC)) [Enter to skip]"
            if ($userInput -match '^\d+$') { $powerParams['SleepTimeoutAC'] = [int]$userInput }

            # Sleep timeout DC
            $userInput = Read-Host "Enter new Sleep Timeout for Battery (current: $(Format-Minutes $currentSettings.SleepTimeoutDC)) [Enter to skip]"
            if ($userInput -match '^\d+$') { $powerParams['SleepTimeoutDC'] = [int]$userInput }

            # Screen saver timeout
            $userInput = Read-Host "Enter new Screen Saver Timeout (current: $(Format-Minutes $currentSettings.ScreenSaverTimeout)) [Enter to skip]"
            if ($userInput -match '^\d+$') { $powerParams['ScreenSaverTimeout'] = [int]$userInput }

            # Group Policy settings params
            $gpoParams = @{}

            # Screen Saver Security Enforcement
            $userInput = Read-Host "Enforce Screen Saver Security? (current: $($lockSettings.ScreenSaverForced)) [Y/N/Enter to skip]"
            if ($userInput -match '^[YN]$') { $gpoParams['ScreenSaverForced'] = ($userInput -eq 'Y') }

            # Auto Lock Enabled
            $userInput = Read-Host "Enable Auto Lock? (current: $($lockSettings.AutoLockEnabled)) [Y/N/Enter to skip]"
            if ($userInput -match '^[YN]$') { $gpoParams['AutoLockEnabled'] = ($userInput -eq 'Y') }

            # Auto Lock Timeout
            # Determine if we should prompt for Auto Lock Timeout
            $promptForAutoLockTimeout = $false
            if ($gpoParams.ContainsKey('AutoLockEnabled')) {
                # If user explicitly set AutoLockEnabled, prompt if it's true
                if ($gpoParams['AutoLockEnabled']) {
                    $promptForAutoLockTimeout = $true
                }
            } elseif ($lockSettings.AutoLockEnabled) {
                # If user didn't explicitly set it, but it was already enabled, prompt
                $promptForAutoLockTimeout = $true
            }

            if ($promptForAutoLockTimeout) {
                $currentAutoLockFormatted = if ($null -ne $lockSettings.AutoLockTimeout) { Format-Minutes $lockSettings.AutoLockTimeout } else { "Not Set" }
                $userInput = Read-Host "Enter Auto Lock Timeout in minutes (current: $currentAutoLockFormatted) [Enter to skip]"
                if ($userInput -match '^\d+$') { $gpoParams['AutoLockTimeout'] = [int]$userInput }
            }

            # Apply power settings if any were changed
            if ($powerParams.Count -gt 0) {
                if (Set-PowerTimeout @powerParams) {
                    # Success message handled within Set-PowerTimeout
                }
            }

            # Apply GPO settings if any were changed
            if ($gpoParams.Count -gt 0) {
                if (Set-LockPolicySettings @gpoParams) {
                    # Success message handled within Set-LockPolicySettings
                }
            }

            # Check if no changes were actually made
            if ($powerParams.Count -eq 0 -and $gpoParams.Count -eq 0) {
                Write-Host "No changes were specified." -ForegroundColor Cyan
            }
        }
    # End of check if settings loaded
        }
}
catch {
    Write-Host "An error occurred: $_" -ForegroundColor Red
    # Log error more formally
        Write-Error "Script failed with error: $_"
}
finally {
    Write-Debug "Entering finally block."
    # Always attempt to stop transcript in finally block if it was started
    Stop-TranscriptSafely

    Write-Debug "Attempting to unregister engine event subscriber."
    # Clean up any remaining event subscribers
    Get-EventSubscriber -ErrorAction SilentlyContinue |
        Where-Object { $_.SourceIdentifier -eq [System.Management.Automation.PsEngineEvent]::Exiting } |
        ForEach-Object {
            Write-Debug "Unregistering event subscription ID: $($_.SubscriptionId)"
            Unregister-Event -SubscriptionId $_.SubscriptionId -ErrorAction SilentlyContinue
        }
    Write-Debug "Finished unregistering engine events."

    # Force final cleanup
    Write-Debug "Triggering final garbage collection in finally block."
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Write-Debug "Exiting finally block."
}
