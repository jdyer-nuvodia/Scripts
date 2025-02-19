# =============================================================================
# Script: Reinstall-ForticlientVPN.ps1
# Created: 2024-02-13 18:30:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2024-02-13 18:30:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0
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
.PARAMETER Verbose
    When specified, provides detailed information about script execution.
.EXAMPLE
    .\Reinstall-ForticlientVPN.ps1
.EXAMPLE
    .\Reinstall-ForticlientVPN.ps1 -Verbose
#>

[CmdletBinding()]
param()

# Ensure running as administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Verbose "Restarting script with elevated privileges..."
    Start-Process PowerShell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

function Stop-ForticlientServices {
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
                    Write-Verbose "Executing: msiexec.exe /x $productCode /qn"
                    Start-Process "msiexec.exe" -ArgumentList "/x $productCode /qn" -Wait
                }
            }
        }
    }
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
        Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath
        Write-Host "Installing Forticlient VPN..."
        Write-Verbose "Executing installer with quiet parameters"
        Start-Process -FilePath $installerPath -ArgumentList "/quiet /norestart" -Wait
        Write-Host "Installation completed successfully"
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
Write-Host "Forticlient VPN reinstallation completed"
