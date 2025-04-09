# =============================================================================
# Script: Get-SetInactivityTimers.ps1
# Created: 2025-04-08 21:45:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-09 18:12:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2.2
# Additional Info: Fixed security policy check formatting causing positional parameter error
# =============================================================================

<#
.SYNOPSIS
Gets and optionally sets Windows system inactivity timers.

.DESCRIPTION
This script retrieves all available system inactivity settings including screen timeout,
sleep settings, power management configurations, and security policies related to machine locking.
It displays these settings to the user and provides the option to modify them. All changes support -WhatIf functionality for safety.

.EXAMPLE
.\Get-SetInactivityTimers.ps1
Displays current inactivity settings and prompts for changes

.EXAMPLE
.\Get-SetInactivityTimers.ps1 -WhatIf
Shows what changes would be made without actually making them
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param()

# Function to format minutes into a readable string
function Format-Minutes {
    param([int]$Minutes)
    if ($Minutes -eq 0) { return "Never" }
    if ($Minutes -ge 1440) { return "$($Minutes / 1440) hours" }
    if ($Minutes -ge 60) { return "$($Minutes / 60) hours" }
    return "$Minutes minutes"
}

function Get-PowerSettings {
    Write-Host "Retrieving current power settings..." -ForegroundColor Cyan
    # Get current power scheme GUID
    $schemeGuid = (powercfg /getactivescheme) -split " " | Where-Object { $_ -match '^[{]?[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}[}]?$' }
    if (-not $schemeGuid) {
        Write-Warning "Could not determine active power scheme GUID"
        return $null
    }
    # Get monitor timeout settings using active scheme
    $monitorTimeoutAC = powercfg /query $schemeGuid SUB_VIDEO VIDEOIDLE | Select-String "AC Power Setting Index: ([0-9a-fx]+)" | ForEach-Object { $_.Matches.Groups[1].Value }
    $monitorTimeoutDC = powercfg /query $schemeGuid SUB_VIDEO VIDEOIDLE | Select-String "DC Power Setting Index: ([0-9a-fx]+)" | ForEach-Object { $_.Matches.Groups[1].Value }
    
    # Get sleep settings using active scheme
    $sleepTimeoutAC = powercfg /query $schemeGuid SUB_SLEEP STANDBYIDLE | Select-String "AC Power Setting Index: ([0-9a-fx]+)" | ForEach-Object { $_.Matches.Groups[1].Value }
    $sleepTimeoutDC = powercfg /query $schemeGuid SUB_SLEEP STANDBYIDLE | Select-String "DC Power Setting Index: ([0-9a-fx]+)" | ForEach-Object { $_.Matches.Groups[1].Value }
    
    # Screen saver settings from registry
    $screenSaverTimeout = Get-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "ScreenSaveTimeout" -ErrorAction SilentlyContinue
    
    return @{
        MonitorAC = $monitorTimeoutAC
        MonitorDC = $monitorTimeoutDC
        SleepAC = $sleepTimeoutAC
        SleepDC = $sleepTimeoutDC
        ScreenSaver = $screenSaverTimeout
    }
}

function Get-LockPolicySettings {
    Write-Host "Checking Group Policy and security settings..." -ForegroundColor Cyan
      $settings = @{
        ScreenSaverForced = $false
        ScreenSaverSecure = $false
        LockoutDuration = $null
        AutoLockEnabled = $false
        AutoLockTimeout = $null
    }

    try {
        # Check screen saver policy settings
        $screenSaverPolicy = Get-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Windows\Control Panel\Desktop" -ErrorAction SilentlyContinue        if ($screenSaverPolicy) {
            $settings.ScreenSaverForced = $screenSaverPolicy.ScreenSaverIsSecure -eq 1
        }
        
        # Check security policy settings
        secedit /export /cfg "$env:TEMP\secpol.cfg" 2>$null
        if (Test-Path "$env:TEMP\secpol.cfg") {
            $secPolContent = Get-Content "$env:TEMP\secpol.cfg" -Raw
            if ($secPolContent -match "LockoutDuration\s*=\s*(\d+)") {
                $settings.LockoutDuration = $matches[1]
            }
            Remove-Item "$env:TEMP\secpol.cfg" -Force
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
            Write-Host "Error setting power settings: $_" -ForegroundColor Red
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
                Write-Host "Setting auto lock enabled status..." -ForegroundColor Cyan
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
            Write-Host "Error setting Group Policy settings: $_" -ForegroundColor Red
            return $false
        }
    }
    return $true
}

# Main script execution
try {
    # Display WhatIf mode disclaimer if applicable
    if ($WhatIfPreference) {
        Write-Host "`n[WhatIf Mode] This script is running in simulation mode. No actual changes will be made.`n" -ForegroundColor Yellow
    }

    # Get current settings
    $currentSettings = Get-PowerSettings
    $lockSettings = Get-LockPolicySettings
    
    # Display current settings
    Write-Host "`nCurrent Inactivity Settings:" -ForegroundColor White
    Write-Host "------------------------" -ForegroundColor White
    Write-Host "Monitor Timeout (AC): $(Format-Minutes $currentSettings.MonitorAC)"
    Write-Host "Monitor Timeout (Battery): $(Format-Minutes $currentSettings.MonitorDC)"
    Write-Host "Sleep Timeout (AC): $(Format-Minutes $currentSettings.SleepAC)"
    Write-Host "Sleep Timeout (Battery): $(Format-Minutes $currentSettings.SleepDC)"
    Write-Host "Screen Saver Timeout: $(Format-Minutes $currentSettings.ScreenSaver)"
      Write-Host "`nSecurity and Group Policy Settings:" -ForegroundColor White
    Write-Host "--------------------------------" -ForegroundColor White
    Write-Host "Screen Saver Security Enforced (User cannot remove password requirement when returning from Screen Saver): $($lockSettings.ScreenSaverForced)"
    Write-Host "Auto Lock Enabled: $($lockSettings.AutoLockEnabled)"
    if ($lockSettings.AutoLockTimeout) {
        Write-Host "Auto Lock Timeout: $(Format-Minutes $lockSettings.AutoLockTimeout)"
    }
    if ($lockSettings.LockoutDuration) {
        Write-Host "Account Lockout Duration: $($lockSettings.LockoutDuration) minutes"
    }
      # Ask if user wants to change settings
    $response = Read-Host "`nWould you like to change these settings? (Y/N)"
      if ($response -eq "Y") {
        # Power settings params
        $powerParams = @{}
        
        # Monitor timeout AC
        $userInput = Read-Host "Enter new Monitor Timeout for AC power (current: $(Format-Minutes $currentSettings.MonitorAC)) [Enter to skip]"
        if ($userInput -match '^\d+$') { $powerParams['MonitorTimeoutAC'] = [int]$userInput }
        
        # Monitor timeout DC
        $userInput = Read-Host "Enter new Monitor Timeout for Battery (current: $(Format-Minutes $currentSettings.MonitorDC)) [Enter to skip]"
        if ($userInput -match '^\d+$') { $powerParams['MonitorTimeoutDC'] = [int]$userInput }
        
        # Sleep timeout AC
        $userInput = Read-Host "Enter new Sleep Timeout for AC power (current: $(Format-Minutes $currentSettings.SleepAC)) [Enter to skip]"
        if ($userInput -match '^\d+$') { $powerParams['SleepTimeoutAC'] = [int]$userInput }
        
        # Sleep timeout DC
        $userInput = Read-Host "Enter new Sleep Timeout for Battery (current: $(Format-Minutes $currentSettings.SleepDC)) [Enter to skip]"
        if ($userInput -match '^\d+$') { $powerParams['SleepTimeoutDC'] = [int]$userInput }
        
        # Screen saver timeout
        $userInput = Read-Host "Enter new Screen Saver Timeout (current: $(Format-Minutes $currentSettings.ScreenSaver)) [Enter to skip]"
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
        if ($gpoParams['AutoLockEnabled'] -or ($lockSettings.AutoLockEnabled -and !$PSBoundParameters.ContainsKey('AutoLockEnabled'))) {
            $userInput = Read-Host "Enter Auto Lock Timeout in minutes (current: $(Format-Minutes $lockSettings.AutoLockTimeout)) [Enter to skip]"
            if ($userInput -match '^\d+$') { $gpoParams['AutoLockTimeout'] = [int]$userInput }
        }
        
        # Apply power settings if any were changed
        if ($powerParams.Count -gt 0) {
            if (Set-PowerTimeout @powerParams) {
                Write-Host "Power settings updated successfully!" -ForegroundColor Green
            }
        }

        # Apply GPO settings if any were changed
        if ($gpoParams.Count -gt 0) {
            if (Set-LockPolicySettings @gpoParams) {
                Write-Host "Group Policy settings updated successfully!" -ForegroundColor Green
            }
        }

        if ($powerParams.Count -eq 0 -and $gpoParams.Count -eq 0) {
            Write-Host "No changes were made." -ForegroundColor Cyan
        }
    }
}
catch {
    Write-Host "An error occurred: $_" -ForegroundColor Red
    exit 1
}
