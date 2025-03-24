# =============================================================================
# Script: Reinstall-DotNET.ps1
# Created: 2024-02-13 19:30:00 UTC
# Author: nunya-nunya
# Last Updated: 2024-02-13 19:30:00 UTC
# Updated By: nunya-nunya
# Version: 1.0
# Additional Info: Script to update all installed .NET Framework versions
# =============================================================================

<#
.SYNOPSIS
    Updates all installed versions of .NET Framework.
.DESCRIPTION
    This script performs the following actions:
    - Detects installed .NET Framework versions
    - Downloads required updates from Microsoft
    - Installs updates silently
    - Verifies successful installation
.PARAMETER Verbose
    When specified, provides detailed information about script execution.
.EXAMPLE
    .\Reinstall-DotNET.ps1
    Updates all installed .NET Framework versions silently
.NOTES
    Security Level: Medium
    Required Permissions: Administrator
    Validation Requirements: Verify .NET versions after update
#>

[CmdletBinding()]
param()

# Ensure running as administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Verbose "Restarting script with elevated privileges..."
    Start-Process PowerShell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

function Get-InstalledDotNetVersions {
    Write-Verbose "Detecting installed .NET Framework versions..."
    $versions = @()
    
    # Check .NET Framework 4.x
    $net4 = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction SilentlyContinue
    if ($net4) {
        $version = $net4.GetValue("Version")
        $release = $net4.GetValue("Release")
        if ($version) {
            $versions += @{
                Version = $version
                Release = $release
                Family = "4"
            }
        }
    }

    # Check .NET Framework 3.5
    $net35 = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5' -ErrorAction SilentlyContinue
    if ($net35) {
        $version = $net35.GetValue("Version")
        if ($version) {
            $versions += @{
                Version = $version
                Family = "3.5"
            }
        }
    }

    return $versions
}

function Update-DotNetFramework {
    $tempPath = Join-Path $env:TEMP "DotNetUpdates"
    New-Item -ItemType Directory -Force -Path $tempPath | Out-Null
    
    try {
        # Update .NET Framework 4.x
        Write-Host "Updating .NET Framework 4.x..."
        $url = "https://go.microsoft.com/fwlink/?LinkId=2085155"
        $installer = Join-Path $tempPath "ndp48-x86-x64-allos-enu.exe"
        
        Write-Verbose "Downloading .NET Framework 4.8 update..."
        Invoke-WebRequest -Uri $url -OutFile $installer
        
        Write-Verbose "Installing .NET Framework 4.8 update..."
        $process = Start-Process -FilePath $installer -ArgumentList "/quiet /norestart" -Wait -PassThru
        
        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
            Write-Host ".NET Framework 4.x update completed successfully"
        } else {
            Write-Warning "Failed to update .NET Framework 4.x. Exit code: $($process.ExitCode)"
        }

        # Enable .NET Framework 3.5 if needed
        Write-Host "Checking .NET Framework 3.5..."
        $state = (Get-WindowsOptionalFeature -Online -FeatureName "NetFx3").State
        if ($state -ne "Enabled") {
            Write-Verbose "Enabling .NET Framework 3.5..."
            Enable-WindowsOptionalFeature -Online -FeatureName "NetFx3" -NoRestart
        }

        # Run Windows Update to get latest .NET updates
        Write-Host "Checking for additional .NET updates through Windows Update..."
        Install-Module PSWindowsUpdate -Force -AllowClobber
        Get-WindowsUpdate -MicrosoftUpdate -Category "Microsoft .NET Framework" -AcceptAll -Install
    }
    catch {
        Write-Error "Error during .NET Framework update: $_"
        Write-Verbose "Stack trace: $($_.ScriptStackTrace)"
    }
    finally {
        Write-Verbose "Cleaning up temporary files..."
        Remove-Item -Path $tempPath -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Test-DotNetUpdates {
    Write-Verbose "Verifying .NET Framework updates..."
    $versions = Get-InstalledDotNetVersions
    
    foreach ($version in $versions) {
        Write-Host "Detected .NET Framework $($version.Family) - Version: $($version.Version)"
    }
}

# Main execution
Write-Verbose "=== Starting .NET Framework update process ==="
Write-Host "Starting .NET Framework updates..."
$originalVersions = Get-InstalledDotNetVersions
Update-DotNetFramework
Test-DotNetUpdates

Write-Host "`nUpdate process completed. A system restart may be required."
$restartPrompt = Read-Host "Would you like to restart now? (Y/N)"
if ($restartPrompt -eq 'Y' -or $restartPrompt -eq 'y') {
    Restart-Computer -Force
}
