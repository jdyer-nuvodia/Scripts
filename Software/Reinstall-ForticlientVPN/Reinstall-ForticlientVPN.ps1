# =============================================================================
# Script: Reinstall-ForticlientVPN.ps1
# Created: 2024-02-13 18:30:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2024-02-13 18:30:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1
# Additional Info: Script to remove and reinstall Forticlient VPN
# =============================================================================

<#
.SYNOPSIS
    Removes existing Forticlient installations and installs latest VPN client.
.DESCRIPTION
    This script performs the following actions:
    - Stops Forticlient services
    - Uninstalls existing Forticlient applications
    - Downloads latest Forticlient VPN installer
    - Installs new Forticlient VPN client
    - Cleans up temporary files

    By default, runs in silent mode with no UI. Use parameters to show installation window or run interactively.
.PARAMETER Verbose
    When specified, provides detailed information about script execution.
.PARAMETER ShowWindow
    Shows the installation window instead of running hidden.
.PARAMETER Interactive
    Runs the installer in interactive mode instead of silent mode.
.EXAMPLE
    .\Reinstall-ForticlientVPN.ps1
    Runs completely silently with no UI
.EXAMPLE
    .\Reinstall-ForticlientVPN.ps1 -ShowWindow
    Shows the installation window but still runs silently
.EXAMPLE
    .\Reinstall-ForticlientVPN.ps1 -Interactive
    Runs the installer in interactive mode with UI
#>

[CmdletBinding()]
param(
    [switch]$ShowWindow,
    [switch]$Interactive
)

# Ensure running as administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Verbose "Restarting script with elevated privileges..."
    Start-Process PowerShell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

function Stop-ForticlientServices {
    Write-Verbose "Searching for Forticlient processes..."
    Get-Process | Where-Object { $_.Name -like "*Forti*" } | ForEach-Object {
        Write-Verbose "Attempting to stop process: $($_.Name)"
        $_ | Stop-Process -Force -ErrorAction SilentlyContinue
    }
    
    Write-Verbose "Searching for Forticlient services..."
    $services = Get-Service -Name "Forticlient*" -ErrorAction SilentlyContinue
    foreach ($service in $services) {
        Write-Verbose "Attempting to stop service: $($service.Name)"
        Stop-Service -Name $service.Name -Force -ErrorAction SilentlyContinue
        Write-Host "Stopped service: $($service.Name)"
    }
}

function Uninstall-ExistingForticlient {
    Write-Verbose "Searching for installed Forticlient applications..."
    $uninstallKeys = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    foreach ($key in $uninstallKeys) {
        Write-Verbose "Checking registry key: $key"
        $apps = Get-ItemProperty $key | Where-Object { $_.DisplayName -like "*Forticlient*" }
        foreach ($app in $apps) {
            if ($app.UninstallString) {
                $uninstallCmd = $app.UninstallString
                if ($uninstallCmd -match "msiexec") {
                    $productCode = $uninstallCmd -replace ".*({.*})", '$1'
                    Write-Verbose "Found product code: $productCode"
                    Write-Host "Uninstalling: $($app.DisplayName)"
                    Write-Verbose "Executing: msiexec.exe /x $productCode /qn /norestart"
                    Start-Process "msiexec.exe" -ArgumentList "/x $productCode /qn /norestart" -Wait
                }
            }
        }
    }
}

function Verify-Installation {
    Write-Verbose "Verifying FortiClient installation..."
    $maxAttempts = 3
    $retryDelay = 10 # seconds
    
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        Write-Verbose "Verification attempt $attempt of $maxAttempts"
        
        # Check registry
        $installed = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | 
                    Where-Object { $_.DisplayName -like "*FortiClient*" }
        
        if (-not $installed) {
            $installed = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | 
                        Where-Object { $_.DisplayName -like "*FortiClient*" }
        }
        
        # Check file system
        $programFiles = @(
            "${env:ProgramFiles}\Fortinet\FortiClient",
            "${env:ProgramFiles(x86)}\Fortinet\FortiClient"
        )
        $filesExist = $programFiles | Where-Object { Test-Path $_ } | Select-Object -First 1
        
        # Check services
        $serviceExists = Get-Service -Name "FortiClient*" -ErrorAction SilentlyContinue

        if ($installed -and $filesExist -and $serviceExists) {
            Write-Host "FortiClient installation verified successfully"
            return $true
        }
        
        if ($attempt -lt $maxAttempts) {
            Write-Verbose "Waiting $retryDelay seconds before next verification attempt..."
            Start-Sleep -Seconds $retryDelay
        }
    }
    
    Write-Warning "FortiClient installation verification failed after $maxAttempts attempts"
    Write-Verbose "Registry check: $($installed -ne $null)"
    Write-Verbose "Files check: $($filesExist -ne $null)"
    Write-Verbose "Service check: $($serviceExists -ne $null)"
    return $false
}

function Install-ForticlientVPN {
    $tempPath = Join-Path $env:TEMP "ForticlientVPN"
    Write-Verbose "Creating temporary directory: $tempPath"
    New-Item -ItemType Directory -Force -Path $tempPath | Out-Null
    $installerPath = Join-Path $tempPath "ForticlientVPN.exe"

    Write-Host "Downloading Forticlient VPN installer..."
    $downloadUrl = "https://links.fortinet.com/forticlient/win/vpnagent"
    
    try {
        Write-Verbose "Downloading from: $downloadUrl"
        Write-Verbose "Saving to: $installerPath"
        
        # Add timeout and retry for download
        $downloadAttempts = 3
        $success = $false
        
        for ($i = 1; $i -le $downloadAttempts; $i++) {
            try {
                Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath -TimeoutSec 60
                $success = $true
                break
            }
            catch {
                Write-Warning "Download attempt $i failed: $_"
                if ($i -lt $downloadAttempts) {
                    Start-Sleep -Seconds 10
                }
            }
        }
        
        if (-not $success) {
            throw "Failed to download installer after $downloadAttempts attempts"
        }
        
        if (-not (Test-Path $installerPath)) {
            throw "Installer file not found at $installerPath"
        }
        
        Write-Host "Installing Forticlient VPN..."

        # Build installation arguments based on parameters
        $installArgs = @()
        if (-not $Interactive) {
            $installArgs += "/quiet", "/accepteula"
        }
        $installArgs += "/norestart"
        
        Write-Verbose "Install arguments: $($installArgs -join ' ')"
        
        $startProcessParams = @{
            FilePath = $installerPath
            ArgumentList = $installArgs
            Wait = $true
        }
        
        if (-not $ShowWindow) {
            $startProcessParams.WindowStyle = 'Hidden'
        }
        
        Write-Verbose "Starting installation process with parameters: $($startProcessParams | ConvertTo-Json)"
        Start-Process @startProcessParams
        
        # Wait for installation to complete and verify
        Start-Sleep -Seconds 10  # Initial wait for installer to start
        $timeout = 300  # 5 minutes timeout
        $timer = [Diagnostics.Stopwatch]::StartNew()
        
        while ($timer.Elapsed.TotalSeconds -lt $timeout) {
            if (-not (Get-Process | Where-Object { $_.Name -like "*FortiClient*Installer*" })) {
                Write-Verbose "Installation process completed"
                break
            }
            Start-Sleep -Seconds 5
        }
        
        if ($timer.Elapsed.TotalSeconds -ge $timeout) {
            throw "Installation timed out after 5 minutes"
        }
        
        # Extended wait and verification
        $verificationAttempts = 6
        $verificationDelay = 30
        
        for ($i = 1; $i -le $verificationAttempts; $i++) {
            Write-Verbose "Verification attempt $i of $verificationAttempts"
            if (Verify-Installation) {
                Write-Host "Installation completed and verified successfully"
                return
            }
            if ($i -lt $verificationAttempts) {
                Write-Verbose "Waiting $verificationDelay seconds before next verification..."
                Start-Sleep -Seconds $verificationDelay
            }
        }
        
        throw "Installation could not be verified after $verificationAttempts attempts"
    }
    catch {
        Write-Error "Error during installation: $_"
        Write-Verbose "Stack trace: $($_.ScriptStackTrace)"
    }
    finally {
        Write-Verbose "Cleaning up temporary directory: $tempPath"
        Remove-Item -Path $tempPath -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# Main execution
Write-Verbose "=== Starting Forticlient VPN reinstallation process ==="
Write-Host "Starting Forticlient VPN reinstallation..."
Stop-ForticlientServices
Write-Verbose "Services stopped successfully"
Uninstall-ExistingForticlient
Write-Verbose "Uninstallation completed"
Install-ForticlientVPN
Write-Verbose "=== Forticlient VPN reinstallation process completed ==="
if (Verify-Installation) {
    Write-Host "FortiClient VPN reinstallation completed and verified"
} else {
    Write-Error "FortiClient VPN reinstallation could not be verified"
}
