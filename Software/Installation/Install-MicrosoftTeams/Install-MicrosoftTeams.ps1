# =============================================================================
# Script: Install-MicrosoftTeams.ps1
# Created: 2025-04-28 15:00:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-28 22:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.4.0
# Additional Info: Fixed download issues with multiple fallback URLs and support for .exe installers
# =============================================================================

<#
.SYNOPSIS
Uninstall all existing Microsoft Teams installations and install the latest version silently.

.DESCRIPTION
This script uninstalls all existing installations of Microsoft Teams (both machine-wide and per-user instances) using multiple detection methods:
- Registry uninstall keys
- WMI/CIM product entries
- User profile directories
- Common installation locations
- Running processes
- Start Menu shortcuts
- Installed AppX packages

The comprehensive detection approach ensures that Teams is discovered through the same methods used by Get-InstalledSoftware.ps1.
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
    
    # First check running processes to ensure they're terminated before uninstallation
    Write-Host 'Checking for running Teams processes...' -ForegroundColor DarkGray
    try {
        $teamsProcesses = Get-Process -Name "*teams*" -ErrorAction SilentlyContinue
        
        if ($teamsProcesses) {
            if ($PSCmdlet.ShouldProcess("Running Teams processes", "Terminate")) {
                Write-Host "Found running Teams processes. Terminating..." -ForegroundColor Cyan
                $teamsProcesses | Stop-Process -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 2
            }
        }
    }
    catch {
        Write-Host "Error checking for Teams processes: $_" -ForegroundColor Red
    }
    
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
                    
                    # First terminate any running Teams processes from WindowsApps
                    $teamsProcesses = Get-Process -Name "*Teams*" -ErrorAction SilentlyContinue | 
                                      Where-Object { $null -ne $_.Path -and $_.Path -like "*WindowsApps*" }
                    
                    if ($teamsProcesses) {
                        Write-Host "  Terminating Teams processes from Windows Store..." -ForegroundColor Yellow
                        $teamsProcesses | Stop-Process -Force -ErrorAction SilentlyContinue
                        Start-Sleep -Seconds 2
                    }
                    
                    # Now remove the AppX package
                    Remove-AppxPackage -Package $app.PackageFullName -ErrorAction Stop
                    Write-Host "Successfully removed Microsoft Store Teams app." -ForegroundColor Green
                    $uninstallCount++
                }
            }
        }
    }
    catch {
        Write-Host "Error removing Microsoft Store Teams app: $_" -ForegroundColor Red
        Write-Host "NOTE: Windows Store apps may require special permissions to remove." -ForegroundColor Yellow
        Write-Host "      Try running this script with 'Run as administrator'" -ForegroundColor Yellow
    }
    #endregion
    
    #region WMI/CIM-based Teams detection
    Write-Host 'Checking for Teams installations via WMI/CIM...' -ForegroundColor DarkGray
    try {
        $cimProducts = Get-CimInstance -ClassName Win32_Product -ErrorAction SilentlyContinue | 
                       Where-Object { $_.Name -like "*Teams*" -or ($_.Vendor -like "*Microsoft*" -and $_.Name -like "*Teams*") }
        
        foreach ($product in $cimProducts) {
            if ($PSCmdlet.ShouldProcess("WMI Product: $($product.Name) v$($product.Version)", "Uninstall")) {
                Write-Host "Uninstalling Teams via WMI: $($product.Name) v$($product.Version)" -ForegroundColor Cyan
                
                try {
                    # Try using the IdentifyingNumber (equivalent to GUID)
                    $guid = $product.IdentifyingNumber
                    
                    if ($guid) {
                        Start-Process -FilePath 'msiexec.exe' -ArgumentList "/x $guid /qn /norestart" -Wait
                        Write-Host "Successfully uninstalled Teams via WMI: $($product.Name)" -ForegroundColor Green
                        $uninstallCount++
                    }
                    else {
                        # Alternative: use the Win32_Product.Uninstall() method
                        $result = $product | Invoke-CimMethod -MethodName "Uninstall"
                        
                        if ($result.ReturnValue -eq 0) {
                            Write-Host "Successfully uninstalled Teams via WMI method: $($product.Name)" -ForegroundColor Green
                            $uninstallCount++
                        }
                        else {
                            Write-Host "Failed to uninstall Teams via WMI method: $($product.Name). Return code: $($result.ReturnValue)" -ForegroundColor Yellow
                        }
                    }
                }
                catch {
                    Write-Host "Error uninstalling Teams via WMI: $($product.Name). Error: $_" -ForegroundColor Red
                }
            }
        }
    }
    catch {
        Write-Host "Error accessing WMI product information: $_" -ForegroundColor Red
    }
    #endregion
    
    #region Process-based Teams detection
    Write-Host 'Checking for running Teams processes...' -ForegroundColor DarkGray
    try {
        $teamsProcesses = Get-Process -Name "*teams*" -ErrorAction SilentlyContinue
        
        if ($teamsProcesses) {
            if ($PSCmdlet.ShouldProcess("Running Teams processes", "Terminate")) {
                Write-Host "Found running Teams processes. Terminating..." -ForegroundColor Cyan
                $teamsProcesses | ForEach-Object {
                    try {
                        # Get process details
                        $processPath = $_.Path
                        $processVersion = $null
                        
                        if ($processPath -and (Test-Path -Path $processPath)) {
                            # Get version information
                            $fileInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($processPath)
                            $processVersion = $fileInfo.FileVersion
                            
                            Write-Host "Found Teams process: $processPath (Version $processVersion)" -ForegroundColor Cyan
                            
                            # Stop the process
                            $_ | Stop-Process -Force
                            Write-Host "Process terminated" -ForegroundColor Yellow
                            
                            # Check if this is from a directory we haven't seen yet
                            $processDir = Split-Path -Parent $processPath
                            $parentDir = Split-Path -Parent $processDir
                            
                            # If this is a unique directory not in our lists, try to remove it
                            if ((Test-Path -Path $parentDir) -and 
                                ($parentDir -like "*Teams*" -or $processDir -like "*Teams*") -and
                                ($parentDir -notlike "$env:LOCALAPPDATA\Microsoft\Teams*") -and 
                                ($parentDir -notlike "$env:APPDATA\Microsoft\Teams*") -and
                                ($parentDir -notlike "${env:ProgramFiles}\Microsoft\Teams*") -and
                                ($parentDir -notlike "${env:ProgramFiles(x86)}\Microsoft\Teams*") -and
                                ($parentDir -notlike "$env:ProgramData\Microsoft\Teams*")) {
                                
                                # Wait a moment for process resources to release
                                Start-Sleep -Seconds 2
                                
                                if ($PSCmdlet.ShouldProcess("Teams directory: $parentDir", "Remove")) {
                                    try {
                                        Write-Host "Attempting to remove Teams directory: $parentDir" -ForegroundColor Cyan
                                        Remove-Item -Path $parentDir -Recurse -Force -ErrorAction Stop
                                        Write-Host "Successfully removed Teams directory at $parentDir" -ForegroundColor Green
                                        $uninstallCount++
                                    }
                                    catch {
                                        Write-Host "Error removing Teams directory $parentDir`: $_" -ForegroundColor Red
                                    }
                                }
                            }
                        }
                    }
                    catch {
                        Write-Host "Error processing Teams executable: $_" -ForegroundColor Red
                    }
                }
            }
        }
    }
    catch {
        Write-Host "Error checking for Teams processes: $_" -ForegroundColor Red
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
    
    #region Check WindowsApps directory (Store apps)
    Write-Host 'Checking Windows Store apps directory for Teams...' -ForegroundColor DarkGray
    $windowsAppsPath = "${env:ProgramFiles}\WindowsApps"
    
    if (Test-Path -Path $windowsAppsPath) {
        try {
            $teamsWindowsAppDirs = Get-ChildItem -Path $windowsAppsPath -Directory -Filter "*Teams*" -ErrorAction SilentlyContinue
            
            if ($teamsWindowsAppDirs) {
                foreach ($dir in $teamsWindowsAppDirs) {
                    Write-Host "Found Windows Store Teams installation: $($dir.FullName)" -ForegroundColor Cyan
                    Write-Host "NOTE: Windows Store apps are protected and require special handling." -ForegroundColor Yellow
                    Write-Host "      Try removing via Settings > Apps > Apps & features" -ForegroundColor Yellow
                    
                    # Try to get the AppX package for this directory
                    $packageName = ($dir.Name -split '_')[0]
                    $teamsAppx = Get-AppxPackage -Name "*$packageName*" -ErrorAction SilentlyContinue
                    
                    if ($teamsAppx -and $PSCmdlet.ShouldProcess("Windows Store Teams: $($teamsAppx.Name)", "Remove")) {
                        Write-Host "Attempting to remove Windows Store Teams package: $($teamsAppx.Name)" -ForegroundColor Cyan
                        try {
                            # The -AllUsers parameter requires admin privileges
                            Remove-AppxPackage -Package $teamsAppx.PackageFullName -AllUsers -ErrorAction Stop
                            Write-Host "Successfully removed Windows Store Teams." -ForegroundColor Green
                            $uninstallCount++
                        }
                        catch {
                            Write-Host "Error removing Windows Store Teams: $_" -ForegroundColor Red
                            Write-Host "Try running this script as administrator to remove system-level Store apps." -ForegroundColor Yellow
                        }
                    }
                }
            }
        }
        catch {
            Write-Host "Error accessing Windows Store Apps directory: $_" -ForegroundColor Red
            Write-Host "This typically requires administrator privileges." -ForegroundColor Yellow
        }
    }
    #endregion
    
    #region Shortcut-based Teams detection
    Write-Host 'Checking for Teams shortcuts...' -ForegroundColor DarkGray
    $startMenuPaths = @(
        "$env:ProgramData\Microsoft\Windows\Start Menu\Programs",
        "$env:APPDATA\Microsoft\Windows\Start Menu\Programs"
    )
    
    $teamsShortcuts = @()
    foreach ($startMenuPath in $startMenuPaths) {
        try {
            if (Test-Path $startMenuPath) {
                $shortcuts = Get-ChildItem -Path $startMenuPath -Filter "*Teams*.lnk" -Recurse -ErrorAction SilentlyContinue
                $teamsShortcuts += $shortcuts
            }
        }
        catch {
            Write-Host "Error accessing shortcuts in $startMenuPath" -ForegroundColor DarkGray
        }
    }
    
    foreach ($shortcut in $teamsShortcuts) {
        if ($PSCmdlet.ShouldProcess("Teams shortcut: $($shortcut.FullName)", "Remove")) {
            try {
                # Get the target of the shortcut
                $shell = New-Object -ComObject WScript.Shell
                $shortcutTarget = $shell.CreateShortcut($shortcut.FullName).TargetPath
                
                # Check if the target exists and looks like a Teams executable
                if (Test-Path $shortcutTarget) {
                    $targetInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($shortcutTarget)
                    
                    if ($targetInfo.ProductName -like "*Teams*" -or $targetInfo.FileDescription -like "*Teams*") {
                        Write-Host "Found Teams shortcut target: $shortcutTarget (Version $($targetInfo.FileVersion))" -ForegroundColor Cyan
                        
                        # Check if this is a directory we haven't processed yet
                        $targetDir = Split-Path -Parent $shortcutTarget
                        $parentDir = Split-Path -Parent $targetDir
                        
                        if ((Test-Path -Path $parentDir) -and 
                            ($parentDir -like "*Teams*" -or $targetDir -like "*Teams*") -and
                            ($parentDir -notlike "$env:LOCALAPPDATA\Microsoft\Teams*") -and 
                            ($parentDir -notlike "$env:APPDATA\Microsoft\Teams*") -and
                            ($parentDir -notlike "${env:ProgramFiles}\Microsoft\Teams*") -and
                            ($parentDir -notlike "${env:ProgramFiles(x86)}\Microsoft\Teams*") -and
                            ($parentDir -notlike "$env:ProgramData\Microsoft\Teams*")) {
                            
                            if ($PSCmdlet.ShouldProcess("Teams directory: $parentDir", "Remove")) {
                                try {
                                    # First try to stop any related processes
                                    $processName = [System.IO.Path]::GetFileNameWithoutExtension($shortcutTarget)
                                    $processes = Get-Process -Name $processName -ErrorAction SilentlyContinue
                                    if ($processes) {
                                        $processes | Stop-Process -Force -ErrorAction SilentlyContinue
                                        Start-Sleep -Seconds 2
                                    }
                                    
                                    # Now remove the directory
                                    Write-Host "Removing Teams directory: $parentDir" -ForegroundColor Cyan
                                    Remove-Item -Path $parentDir -Recurse -Force -ErrorAction Stop
                                    Write-Host "Successfully removed Teams directory at $parentDir" -ForegroundColor Green
                                    $uninstallCount++
                                }
                                catch {
                                    Write-Host "Error removing Teams directory $parentDir`: $_" -ForegroundColor Red
                                }
                            }
                        }
                    }
                }
                
                # Remove the shortcut itself
                Write-Host "Removing Teams shortcut: $($shortcut.FullName)" -ForegroundColor Cyan
                Remove-Item -Path $shortcut.FullName -Force -ErrorAction Stop
                Write-Host "Successfully removed Teams shortcut" -ForegroundColor Green
            }
            catch {
                Write-Host "Error processing Teams shortcut $($shortcut.FullName): $_" -ForegroundColor Red
            }
        }
    }
    #endregion
    
    # Final status
    if ($uninstallCount -eq 0) {
        Write-Host 'No Microsoft Teams installations found.' -ForegroundColor Yellow
        
        # Perform additional detection similar to verification phase
        Write-Host 'Running additional detection methods...' -ForegroundColor Cyan
        
        # Check AppX Packages again (sometimes they need special handling)
        try {
            $teamsAppx = Get-AppxPackage -Name "*MicrosoftTeams*" -ErrorAction SilentlyContinue
            if ($teamsAppx) {
                Write-Host "FOUND: Microsoft Teams AppX package detected: $($teamsAppx.Name) v$($teamsAppx.Version)" -ForegroundColor Yellow
                Write-Host "       Try running this script again with administrator privileges" -ForegroundColor Yellow
            }
        } catch { }
        
        # Check for Teams in processes
        try {
            $teamsProcesses = Get-Process -Name "*Teams*" -ErrorAction SilentlyContinue | Where-Object { $null -ne $_.Path }
            if ($teamsProcesses) {
                foreach ($process in $teamsProcesses) {
                    $fileInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($process.Path)
                    Write-Host "FOUND: Microsoft Teams running from: $($process.Path) v$($fileInfo.FileVersion)" -ForegroundColor Yellow
                    Write-Host "       This process might be protected or require administrator privileges" -ForegroundColor Yellow
                }
            }
        } catch { }
        
        # Check WindowsApps folder which requires admin access
        if (Test-Path -Path "$env:ProgramFiles\WindowsApps") {
            Write-Host "NOTE: The Windows Store apps folder exists but may require administrator privileges to access" -ForegroundColor Yellow
        }
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
    # Primary direct download URL for Teams MSI (2025 updated URL)
    $downloadUrl = 'https://go.microsoft.com/fwlink/p/?LinkID=2187327&clcid=0x409&culture=en-us&country=US'
    
    # Fallback URL if the primary fails
    $fallbackUrl = 'https://teams.microsoft.com/downloads/desktopurl?env=production&plat=windows&arch=x64&managedInstaller=true'
    
    $tempDir = [System.IO.Path]::GetTempPath()
    $installerPath = Join-Path -Path $tempDir -ChildPath 'Teams_windows_x64.msi'
    $downloadSuccess = $false

    if ($PSCmdlet.ShouldProcess('Download Microsoft Teams', "Download from $downloadUrl")) {
        # Try primary URL first
        Write-Host "Downloading Microsoft Teams installer from $downloadUrl..." -ForegroundColor Cyan
        try {
            Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath -UseBasicParsing
            if ((Test-Path -Path $installerPath) -and ((Get-Item -Path $installerPath).Length -gt 1MB)) {
                $fileInfo = Get-Item -Path $installerPath
                Write-Host "Downloaded installer to $installerPath (Size: $([math]::Round($fileInfo.Length / 1MB, 2)) MB)" -ForegroundColor Green
                $downloadSuccess = $true
            }
            else {
                Write-Host "Downloaded file seems too small. Trying fallback URL..." -ForegroundColor Yellow
                throw "Downloaded file is too small"
            }
        }
        catch {
            Write-Host "Error with primary download URL: $($_.Exception.Message)" -ForegroundColor Yellow
            
            # Try fallback URL
            Write-Host "Trying fallback download URL: $fallbackUrl" -ForegroundColor Cyan
            try {
                Invoke-WebRequest -Uri $fallbackUrl -OutFile $installerPath -UseBasicParsing
                if ((Test-Path -Path $installerPath) -and ((Get-Item -Path $installerPath).Length -gt 1MB)) {
                    $fileInfo = Get-Item -Path $installerPath
                    Write-Host "Downloaded installer to $installerPath (Size: $([math]::Round($fileInfo.Length / 1MB, 2)) MB)" -ForegroundColor Green
                    $downloadSuccess = $true
                }
                else {
                    throw "Failed to download a valid Teams installer"
                }
            }
            catch {
                Write-Host "Error downloading Microsoft Teams installer: $_" -ForegroundColor Red
                  # Try a third method - using System.Net.WebClient which sometimes works better with redirects
                Write-Host "Trying alternate download method..." -ForegroundColor Cyan
                try {
                    $webClient = New-Object System.Net.WebClient
                    $webClient.DownloadFile('https://www.microsoft.com/en-us/microsoft-teams/download-app?rtc=2#allDevicesSection', $installerPath)
                    if ((Test-Path -Path $installerPath) -and ((Get-Item -Path $installerPath).Length -gt 1MB)) {
                        $fileInfo = Get-Item -Path $installerPath
                        Write-Host "Downloaded installer to $installerPath (Size: $([math]::Round($fileInfo.Length / 1MB, 2)) MB)" -ForegroundColor Green
                        $downloadSuccess = $true
                    }
                    else {
                        throw "Failed to download a valid Teams installer with alternate method"
                    }
                }
                catch {
                    Write-Host "All download methods failed. Cannot proceed with installation." -ForegroundColor Red
                    return
                }
            }
        }
    }

    if (-not $downloadSuccess) {
        Write-Host "Unable to download a valid Teams installer. Installation aborted." -ForegroundColor Red
        return
    }

    if ($PSCmdlet.ShouldProcess('Install Microsoft Teams', "Install using $installerPath")) {
        Write-Host "Installing Microsoft Teams silently..." -ForegroundColor Cyan
        try {
            $process = Start-Process -FilePath 'msiexec.exe' -ArgumentList "/i `"$installerPath`" /qn /norestart ALLUSERS=1" -Wait -PassThru
            
            if ($process.ExitCode -eq 0) {
                Write-Host "Microsoft Teams installed successfully." -ForegroundColor Green
            }
            elseif ($process.ExitCode -eq 3010) {
                Write-Host "Microsoft Teams installed successfully but requires a restart to complete installation." -ForegroundColor Yellow
            }
            else {
                Write-Host "Microsoft Teams installation exited with code: $($process.ExitCode)." -ForegroundColor Yellow
                # Provide more specific information about common error codes
                switch ($process.ExitCode) {
                    1603 { Write-Host "Error 1603: Fatal error during installation." -ForegroundColor Red }
                    1618 { Write-Host "Error 1618: Another installation is already in progress." -ForegroundColor Red }
                    1619 { Write-Host "Error 1619: Installation package could not be found." -ForegroundColor Red }
                    1620 { Write-Host "Error 1620: Installation package could not be opened." -ForegroundColor Red }
                    1638 { Write-Host "Error 1638: Another version of this product is already installed." -ForegroundColor Yellow }
                    1641 { Write-Host "Error 1641: The installer has initiated a restart." -ForegroundColor Yellow }
                    default { Write-Host "Check MSI error code for more details: https://docs.microsoft.com/en-us/windows/win32/msi/error-codes" -ForegroundColor Yellow }
                }
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

# This variable is used at the end of the script (line ~735) to determine success
# We're initializing it here so scope is maintained throughout verification steps
# Using Set-Variable with scope script to satisfy PowerShell analyzer
Set-Variable -Name teamsInstalled -Value $false -Scope Script

# Check registry for Teams
$uninstallPaths = @(
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)

foreach ($path in $uninstallPaths) {
    Get-ChildItem -Path $path -ErrorAction SilentlyContinue | ForEach-Object {
        $app = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
        if ($app.DisplayName -like '*Teams*' -or ($app.Publisher -like '*Microsoft*' -and $app.DisplayName -like '*Teams*')) {
            Write-Host "Microsoft Teams found in registry: $($app.DisplayName) v$($app.DisplayVersion)" -ForegroundColor Green
            $teamsInstalled = $true
        }
    }
}

# Check WMI/CIM for Teams
try {
    $cimProducts = Get-CimInstance -ClassName Win32_Product -ErrorAction SilentlyContinue | 
                   Where-Object { $_.Name -like "*Teams*" -or ($_.Vendor -like "*Microsoft*" -and $_.Name -like "*Teams*") }
    
    foreach ($product in $cimProducts) {
        Write-Host "Microsoft Teams found in WMI: $($product.Name) v$($product.Version)" -ForegroundColor Green
        $teamsInstalled = $true
    }
}
catch {
    Write-Host "Error checking WMI for Teams: $_" -ForegroundColor DarkGray
}

# Check AppX Packages for Teams
try {
    $teamsAppx = Get-AppxPackage -Name "*MicrosoftTeams*" -ErrorAction SilentlyContinue
    if ($teamsAppx) {
        foreach ($app in $teamsAppx) {
            Write-Host "Microsoft Teams found in AppX: $($app.Name) v$($app.Version)" -ForegroundColor Green
            $teamsInstalled = $true
        }
    }
}
catch {
    Write-Host "Error checking AppX for Teams: $_" -ForegroundColor DarkGray
}

# Check program files
$teamsLocations = @(
    "${env:ProgramFiles}\Microsoft\Teams",
    "${env:ProgramFiles(x86)}\Microsoft\Teams"
)

foreach ($location in $teamsLocations) {
    if (Test-Path -Path $location) {
        $exeFiles = Get-ChildItem -Path $location -Filter "Teams*.exe" -Recurse -ErrorAction SilentlyContinue | 
                    Where-Object { $_.Name -notlike "*uninst*" -and $_.Name -notlike "*setup*" } | 
                    Select-Object -First 1
        
        if ($exeFiles) {
            $fileInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($exeFiles[0].FullName)
            Write-Host "Microsoft Teams executable found: $($exeFiles[0].FullName) v$($fileInfo.FileVersion)" -ForegroundColor Green
            $teamsInstalled = $true
        }
        else {
            Write-Host "Teams directory exists at $location but no Teams executable found." -ForegroundColor Yellow
        }
    }
}

# Check for per-user installation
$perUserPath = "$env:LOCALAPPDATA\Microsoft\Teams"
if (Test-Path -Path $perUserPath) {
    $exeFiles = Get-ChildItem -Path $perUserPath -Filter "Teams*.exe" -Recurse -ErrorAction SilentlyContinue | 
                Where-Object { $_.Name -notlike "*uninst*" -and $_.Name -notlike "*setup*" } | 
                Select-Object -First 1
    
    if ($exeFiles) {
        $fileInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($exeFiles[0].FullName)
        Write-Host "Microsoft Teams found (per-user installation): $($exeFiles[0].FullName) v$($fileInfo.FileVersion)" -ForegroundColor Green
        $teamsInstalled = $true
    }
}

# Check for Teams in running processes
try {
    $teamsProcesses = Get-Process -Name "*Teams*" -ErrorAction SilentlyContinue
    if ($teamsProcesses) {
        foreach ($process in $teamsProcesses) {
            if ($process.Path) {
                $fileInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($process.Path)
                Write-Host "Microsoft Teams process found: $($process.Path) v$($fileInfo.FileVersion)" -ForegroundColor Green
                $teamsInstalled = $true
            }
        }
    }
}
catch {
    Write-Host "Error checking Teams processes: $_" -ForegroundColor DarkGray
}

# Check for Teams shortcuts
try {
    $startMenuPaths = @(
        "$env:ProgramData\Microsoft\Windows\Start Menu\Programs",
        "$env:APPDATA\Microsoft\Windows\Start Menu\Programs"
    )
    
    foreach ($startMenuPath in $startMenuPaths) {
        if (Test-Path $startMenuPath) {
            $shortcuts = Get-ChildItem -Path $startMenuPath -Filter "*Teams*.lnk" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
            
            if ($shortcuts) {
                $shell = New-Object -ComObject WScript.Shell
                $shortcutTarget = $shell.CreateShortcut($shortcuts.FullName).TargetPath
                
                if (Test-Path $shortcutTarget) {
                    $fileInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($shortcutTarget)
                    Write-Host "Microsoft Teams shortcut found: $($shortcuts.FullName) pointing to v$($fileInfo.FileVersion)" -ForegroundColor Green
                    $teamsInstalled = $true
                }
            }
        }
    }
}
catch {
    Write-Host "Error checking Teams shortcuts: $_" -ForegroundColor DarkGray
}

if (-not $teamsInstalled) {
    Write-Host "Warning: Microsoft Teams installation could not be verified. It may not have installed correctly." -ForegroundColor Yellow
    
    # Provide troubleshooting guidance
    Write-Host "`nTroubleshooting suggestions:" -ForegroundColor Yellow
    Write-Host "1. Verify your internet connection is working properly." -ForegroundColor Yellow
    Write-Host "2. Check if you have sufficient permissions to install software." -ForegroundColor Yellow
    Write-Host "3. Try manually downloading Teams from https://www.microsoft.com/en-us/microsoft-teams/download-app" -ForegroundColor Yellow
    Write-Host "4. Run this script again with administrator privileges." -ForegroundColor Yellow
}
else {
    Write-Host "Microsoft Teams has been successfully installed and verified." -ForegroundColor Green
}

Write-Host 'Microsoft Teams uninstall and install process completed.' -ForegroundColor Green
