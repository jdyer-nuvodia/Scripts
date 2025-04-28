# =============================================================================
# Script: Install-MicrosoftTeams.ps1
# Created: 2025-04-28 15:00:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-28 15:00:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.0
# Additional Info: Script to uninstall all versions of Microsoft Teams and install latest silently
# =============================================================================

<#[
.SYNOPSIS
Uninstall all existing Microsoft Teams installations and install the latest version silently.

.DESCRIPTION
This script uninstalls all existing installations of Microsoft Teams (both machine-wide and per-user instances) by querying the registry uninstall keys and invoking msiexec. It then downloads the latest Teams MSI installer from Microsoft and installs it silently.
It supports PowerShell 5.1 and later versions and includes -WhatIf support for all actions.

.PARAMETER None
This script does not accept parameters. Use -WhatIf to simulate actions.

.EXAMPLE
.\Install-MicrosoftTeams.ps1 -WhatIf
Simulates uninstallation and installation actions without making any changes.

.EXAMPLE
.\Install-MicrosoftTeams.ps1
Uninstalls existing Teams installations and installs the latest version silently.
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param()

function Uninstall-Teams {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
    param()
    Write-Host 'Scanning for existing Microsoft Teams installations...' -ForegroundColor Cyan

    $uninstallPaths = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall'
    )

    $teamsApps = foreach ($path in $uninstallPaths) {
        Get-ChildItem -Path $path -ErrorAction SilentlyContinue | ForEach-Object {
            $app = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            if ($app.DisplayName -like '*Teams*') {
                [PSCustomObject]@{
                    DisplayName = $app.DisplayName
                    GUID        = $_.PSChildName
                }
            }
        }
    }

    if (-not $teamsApps) {
        Write-Host 'No Microsoft Teams installations found.' -ForegroundColor Yellow
        return
    }

    foreach ($app in $teamsApps) {
        if ($PSCmdlet.ShouldProcess($app.DisplayName, 'Uninstall')) {
            Write-Host "Uninstalling $($app.DisplayName) GUID $($app.GUID)..." -ForegroundColor Cyan
            Start-Process -FilePath 'msiexec.exe' -ArgumentList "/x $($app.GUID) /qn /norestart" -Wait
            Write-Host "Successfully uninstalled $($app.DisplayName)." -ForegroundColor Green
        }
    }
}

function Install-Teams {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
    param()
    $downloadUrl    = 'https://aka.ms/teamsmsi'
    $tempDir        = [System.IO.Path]::GetTempPath()
    $installerPath  = Join-Path -Path $tempDir -ChildPath 'Teams_latest.msi'

    if ($PSCmdlet.ShouldProcess('Download Microsoft Teams', "Download from $downloadUrl")) {
        Write-Host "Downloading Microsoft Teams installer from $downloadUrl..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath -UseBasicParsing
        Write-Host "Downloaded installer to $installerPath." -ForegroundColor Green
    }

    if ($PSCmdlet.ShouldProcess('Install Microsoft Teams', "Install using $installerPath")) {
        Write-Host "Installing Microsoft Teams silently..." -ForegroundColor Cyan
        Start-Process -FilePath 'msiexec.exe' -ArgumentList "/i `"$installerPath`" /qn /norestart" -Wait
        Write-Host "Microsoft Teams installed successfully." -ForegroundColor Green
    }
}

# Main script execution
Write-Host 'Starting Microsoft Teams uninstall and install process.' -ForegroundColor Cyan
Uninstall-Teams
Install-Teams
Write-Host 'Microsoft Teams uninstall and install process completed.' -ForegroundColor Green
