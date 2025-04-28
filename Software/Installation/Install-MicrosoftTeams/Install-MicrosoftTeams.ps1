# =============================================================================
# Script: Install-MicrosoftTeams.ps1
# Created: 2025-04-28 15:00:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-28 22:24:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1.0
# Additional Info: Enhanced Teams detection and uninstallation process
# =============================================================================

<#
.SYNOPSIS
Uninstall all existing Microsoft Teams installations and install the latest version silently.

.DESCRIPTION
This script uninstalls all existing installations of Microsoft Teams (both machine-wide and per-user instances) using multiple detection methods:
- Registry uninstall keys
- User profile directories
- Common installation locations
- Running processes
- Installed AppX packages

After uninstalling all Teams instances, it downloads the latest Teams MSI installer from Microsoft and installs it silently.
The script supports PowerShell 5.1 and later versions and includes -WhatIf support for all actions.

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
    
    # Track successful uninstalls
    $uninstallCount = 0
    
    #region Registry-based uninstallation
    Write-Host 'Checking registry uninstall keys...' -ForegroundColor DarkGray
    $uninstallPaths = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall'
    )

    $teamsApps = foreach ($path in $uninstallPaths) {
        Get-ChildItem -Path $path -ErrorAction SilentlyContinue | ForEach-Object {
            $app = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            if ($app.DisplayName -like '*Teams*' -or $app.Publisher -like '*Microsoft*' -and $app.DisplayName -like '*Teams*') {
                [PSCustomObject]@{
                    DisplayName = $app.DisplayName
                    GUID        = $_.PSChildName
                    Source      = 'Registry'
                }
            }
        }
    }

    foreach ($app in $teamsApps) {
        if ($PSCmdlet.ShouldProcess($app.DisplayName, 'Uninstall (Registry GUID)')) {
            try {
                Write-Host "Uninstalling $($app.DisplayName) GUID $($app.GUID)..." -ForegroundColor Cyan
                Start-Process -FilePath 'msiexec.exe' -ArgumentList "/x $($app.GUID) /qn /norestart" -Wait
                Write-Host "Successfully uninstalled $($app.DisplayName)." -ForegroundColor Green
                $uninstallCount++
            }
            catch {
                Write-Host "Error uninstalling $($app.DisplayName): $_" -ForegroundColor Red
            }
        }
    }
    #endregion
    
    #region Per-user Teams installations
    Write-Host 'Checking for per-user Teams installations...' -ForegroundColor DarkGray
    $teamsUserFolders = @(
        "$env:LOCALAPPDATA\Microsoft\Teams",
        "$env:APPDATA\Microsoft\Teams"
    )
    
    foreach ($folder in $teamsUserFolders) {
        if (Test-Path -Path $folder) {
            if ($PSCmdlet.ShouldProcess("Teams folder: $folder", "Remove")) {
                try {
                    # First, kill any running Teams processes
                    $processes = Get-Process -Name "*teams*" -ErrorAction SilentlyContinue
                    if ($processes) {
                        Write-Host "Stopping Teams processes..." -ForegroundColor Yellow
                        $processes | Stop-Process -Force -ErrorAction SilentlyContinue
                        Start-Sleep -Seconds 2
                    }
                    
                    # Try to uninstall using Update.exe
                    $updateExe = Join-Path -Path $folder -ChildPath "Update.exe"
                    if (Test-Path -Path $updateExe) {
                        Write-Host "Running Teams uninstaller: $updateExe --uninstall" -ForegroundColor Cyan
                        Start-Process -FilePath $updateExe -ArgumentList "--uninstall" -Wait -ErrorAction SilentlyContinue
                    }
                    
                    # Remove the folder
                    Write-Host "Removing Teams folder: $folder" -ForegroundColor Cyan
                    Remove-Item -Path $folder -Recurse -Force -ErrorAction Stop
                    Write-Host "Successfully removed per-user Teams installation at $folder" -ForegroundColor Green
                    $uninstallCount++
                }
                catch {
                    Write-Host "Error removing Teams folder $folder`: $_" -ForegroundColor Red
                }
            }
        }
    }
    #endregion
    
    #region AppX/Store version of Teams
    Write-Host 'Checking for Microsoft Store version of Teams...' -ForegroundColor DarkGray
    try {
        $teamsAppx = Get-AppxPackage -Name "*MicrosoftTeams*" -ErrorAction SilentlyContinue
        if ($teamsAppx) {
            foreach ($app in $teamsAppx) {
                if ($PSCmdlet.ShouldProcess("Microsoft Store Teams: $($app.Name) v$($app.Version)", "Remove")) {
                    Write-Host "Removing Microsoft Store Teams app: $($app.Name) v$($app.Version)" -ForegroundColor Cyan
                    Remove-AppxPackage -Package $app.PackageFullName -ErrorAction Stop
                    Write-Host "Successfully removed Microsoft Store Teams app." -ForegroundColor Green
                    $uninstallCount++
                }
            }
        }
    }
    catch {
        Write-Host "Error removing Microsoft Store Teams app: $_" -ForegroundColor Red
    }
    #endregion
    
    #region Check common installation directories
    Write-Host 'Checking common Teams installation directories...' -ForegroundColor DarkGray
    $teamsCommonLocations = @(
        "${env:ProgramFiles}\Microsoft\Teams",
        "${env:ProgramFiles(x86)}\Microsoft\Teams",
        "$env:ProgramData\Microsoft\Teams"
    )
    
    foreach ($location in $teamsCommonLocations) {
        if (Test-Path -Path $location) {
            if ($PSCmdlet.ShouldProcess("Teams directory: $location", "Remove")) {
                try {
                    # Look for uninstaller exe or MSI
                    $uninstallers = Get-ChildItem -Path $location -Filter "*uninst*.exe" -Recurse -ErrorAction SilentlyContinue
                    foreach ($uninstaller in $uninstallers) {
                        Write-Host "Running Teams uninstaller: $($uninstaller.FullName)" -ForegroundColor Cyan
                        Start-Process -FilePath $uninstaller.FullName -ArgumentList "/S", "/Silent", "/Q", "/quiet", "/qn", "/norestart" -Wait -ErrorAction SilentlyContinue
                    }
                    
                    # Remove the directory
                    Write-Host "Removing Teams directory: $location" -ForegroundColor Cyan
                    Remove-Item -Path $location -Recurse -Force -ErrorAction Stop
                    Write-Host "Successfully removed Teams installation at $location" -ForegroundColor Green
                    $uninstallCount++
                }
                catch {
                    Write-Host "Error removing Teams directory $location`: $_" -ForegroundColor Red
                }
            }
        }
    }
    #endregion
    
    # Final status
    if ($uninstallCount -eq 0) {
        Write-Host 'No Microsoft Teams installations found.' -ForegroundColor Yellow
    }
    else {
        Write-Host "Successfully uninstalled/removed $uninstallCount Teams components." -ForegroundColor Green
    }
    
    # Give the system a moment to finalize uninstallation
    Start-Sleep -Seconds 3
}

function Install-Teams {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
    param()
    $downloadUrl    = 'https://aka.ms/teamsmsi'
    $tempDir        = [System.IO.Path]::GetTempPath()
    $installerPath  = Join-Path -Path $tempDir -ChildPath 'Teams_latest.msi'

    if ($PSCmdlet.ShouldProcess('Download Microsoft Teams', "Download from $downloadUrl")) {
        Write-Host "Downloading Microsoft Teams installer from $downloadUrl..." -ForegroundColor Cyan
        try {
            Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath -UseBasicParsing
            if (Test-Path -Path $installerPath) {
                $fileInfo = Get-Item -Path $installerPath
                Write-Host "Downloaded installer to $installerPath (Size: $([math]::Round($fileInfo.Length / 1MB, 2)) MB)" -ForegroundColor Green
            }
            else {
                throw "Failed to download installer to $installerPath"
            }
        }
        catch {
            Write-Host "Error downloading Microsoft Teams installer: $_" -ForegroundColor Red
            return
        }
    }

    if ($PSCmdlet.ShouldProcess('Install Microsoft Teams', "Install using $installerPath")) {
        Write-Host "Installing Microsoft Teams silently..." -ForegroundColor Cyan
        try {
            $process = Start-Process -FilePath 'msiexec.exe' -ArgumentList "/i `"$installerPath`" /qn /norestart ALLUSERS=1" -Wait -PassThru
            
            if ($process.ExitCode -eq 0) {
                Write-Host "Microsoft Teams installed successfully." -ForegroundColor Green
            }
            else {
                Write-Host "Microsoft Teams installation exited with code: $($process.ExitCode). There might be an issue with the installation." -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host "Error installing Microsoft Teams: $_" -ForegroundColor Red
        }
    }
}

# Main script execution
Write-Host 'Starting Microsoft Teams uninstall and install process.' -ForegroundColor Cyan
Write-Host '----------------------------------------------------------------' -ForegroundColor DarkGray

# Uninstall any existing Teams instances
Uninstall-Teams

# Install latest version
Install-Teams

# Verify installation
Write-Host '----------------------------------------------------------------' -ForegroundColor DarkGray
Write-Host 'Verifying Microsoft Teams installation...' -ForegroundColor Cyan

$teamsInstalled = $false

# Check registry for Teams
$uninstallPaths = @(
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)

foreach ($path in $uninstallPaths) {
    Get-ChildItem -Path $path -ErrorAction SilentlyContinue | ForEach-Object {
        $app = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
        if ($app.DisplayName -like '*Teams*') {
            Write-Host "Microsoft Teams found in registry: $($app.DisplayName) v$($app.DisplayVersion)" -ForegroundColor Green
            $teamsInstalled = $true
        }
    }
}

# Check program files
$teamsLocations = @(
    "${env:ProgramFiles}\Microsoft\Teams",
    "${env:ProgramFiles(x86)}\Microsoft\Teams"
)

foreach ($location in $teamsLocations) {
    if (Test-Path -Path $location) {
        $exeFiles = Get-ChildItem -Path $location -Filter "Teams*.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($exeFiles) {
            $fileInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($exeFiles[0].FullName)
            Write-Host "Microsoft Teams executable found: $($exeFiles[0].FullName) v$($fileInfo.FileVersion)" -ForegroundColor Green
            $teamsInstalled = $true
        }
    }
}

if (-not $teamsInstalled) {
    Write-Host "Warning: Microsoft Teams installation could not be verified. It may not have installed correctly." -ForegroundColor Yellow
}

Write-Host 'Microsoft Teams uninstall and install process completed.' -ForegroundColor Green
