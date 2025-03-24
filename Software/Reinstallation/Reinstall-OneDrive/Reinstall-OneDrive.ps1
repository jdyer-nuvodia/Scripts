# =============================================================================
# Script: Reinstall-OneDrive.ps1
# Created: 2025-02-27 18:50:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-27 18:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0
# Additional Info: Script to uninstall and reinstall OneDrive
# =============================================================================

<#
.SYNOPSIS
    Uninstalls existing OneDrive installations and installs the latest version.
.DESCRIPTION
    This script performs the following actions:
     - Uninstalls all existing OneDrive installations
     - Downloads and installs the latest version of OneDrive
     - Includes timeout protection for long-running operations
     - Requires administrator privileges
.EXAMPLE
    .\Reinstall-OneDrive.ps1
    Completely reinstalls OneDrive with the latest version
.NOTES
    Security Level: Medium
    Required Permissions: Administrator
    Validation Requirements: Verify OneDrive is running after installation
#>

# Requires administrator privileges
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Script is not running as Administrator. Restarting with elevated privileges..." -ForegroundColor Yellow
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-File `"$($MyInvocation.MyCommand.Path)`""
    Exit
}

# Main script execution with timeout
$job = Start-Job -ScriptBlock {
    # Function to uninstall OneDrive
    function Uninstall-OneDrive {
        Write-Host "Stopping OneDrive processes..." -ForegroundColor Cyan
        Stop-Process -Name OneDrive* -Force -ErrorAction SilentlyContinue

        $oneDrivePaths = @(
            "$env:SystemRoot\System32\OneDriveSetup.exe",
            "$env:SystemRoot\SysWOW64\OneDriveSetup.exe",
            "${env:ProgramFiles}\Microsoft OneDrive\OneDriveSetup.exe",
            "${env:ProgramFiles(x86)}\Microsoft OneDrive\OneDriveSetup.exe"
        )

        foreach ($StartPath in $oneDrivePaths) {
            if (Test-Path $StartPath) {
                Write-Host "Uninstalling OneDrive from $StartPath" -ForegroundColor Cyan
                Start-Process $StartPath -ArgumentList "/uninstall" -Wait
            }
        }

        Write-Host "Uninstalling OneDrive using WinGet..." -ForegroundColor Cyan
        winget uninstall Microsoft.OneDrive
    }

    # Function to download and install the latest OneDrive
    function Install-LatestOneDrive {
        $url = "https://go.microsoft.com/fwlink/p/?LinkID=2182910"
        $outPath = "$env:TEMP\OneDriveSetup.exe"

        Write-Host "Downloading the latest OneDrive installer..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $url -OutFile $outPath

        Write-Host "Installing the latest version of OneDrive..." -ForegroundColor Cyan
        Start-Process $outPath -ArgumentList "/allusers" -Wait
    }

    # Execute the functions
    Write-Host "Starting OneDrive reinstallation process..." -ForegroundColor Cyan
    Uninstall-OneDrive

    Write-Host "Installing the latest version of OneDrive..." -ForegroundColor Cyan
    Install-LatestOneDrive

    Write-Host "OneDrive update process completed successfully." -ForegroundColor Green
}

$timeout = 300 # 5 minutes in seconds
$completed = Wait-Job $job -Timeout $timeout

if ($completed -eq $null) {
    Write-Host "The script did not complete within 5 minutes. Stopping the process..." -ForegroundColor Red
    Stop-Job $job
    Remove-Job $job
} else {
    Receive-Job $job
    Remove-Job $job
}
