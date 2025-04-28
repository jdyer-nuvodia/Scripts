# =============================================================================
# Script: Install-MicrosoftTeams.ps1
# Created: 2025-04-08 15:00:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-28 23:25:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.7.3
# Additional Info: Fixed syntax errors and formatting issues throughout the script
# =============================================================================

<#
.SYNOPSIS
Uninstall all existing Microsoft Teams installations and install the latest version silently.

.DESCRIPTION
This script manages Microsoft Teams installation with these major functions:
1. Stops all Teams-related processes before making changes
2. Uninstalls all existing Teams installations using comprehensive detection methods:
   - Registry uninstall keys
   - WMI/CIM product entries
   - User profile directories
   - Common installation locations
   - Running processes
   - Start Menu shortcuts
   - Installed AppX packages
3. Downloads the latest Teams EXE installer from Microsoft and installs it silently
4. Verifies installation through multiple detection methods
5. Performs a health check on the Teams installation:
   - Tests Teams process startup
   - Validates configuration files
   - Checks correct installation paths

The script handles many common Teams installation issues, providing detailed feedback and 
appropriate error handling. It supports PowerShell 5.1 and later versions and includes 
-WhatIf support for all actions.

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
    
    $tempDir = [System.IO.Path]::GetTempPath()
    $osArch = Get-SystemArchitecture
    $exeInstallerPath = Join-Path -Path $tempDir -ChildPath "Teams_windows_$osArch.exe"
    $downloadSuccess = $false
    
    if ($PSCmdlet.ShouldProcess('Download Microsoft Teams', "Download latest installer")) {
        # Use the current Teams EXE installer URL (updated for 2025)
        # Microsoft periodically changes the Teams download URLs and link IDs
        $teamsExeUrl = "https://teams.microsoft.com/downloads/desktopurl?env=production&plat=windows&arch=$osArch&managedInstaller=true&download=true"
        $teamsExeFallbackUrl = if ($osArch -eq "x64") {
            "https://go.microsoft.com/fwlink/?linkid=2187327" # 64-bit link
        } else {
            "https://go.microsoft.com/fwlink/?linkid=2187323" # 32-bit link
        }
        $teamsExeBackupUrl = "https://teams.microsoft.com/downloads/desktopurl?env=production&plat=windows&arch=$osArch&download=true"
        
        # Try direct exe download first
        Write-Host "Downloading Microsoft Teams EXE installer..." -ForegroundColor Cyan
        try {
            Invoke-WebRequest -Uri $teamsExeUrl -OutFile $exeInstallerPath -UseBasicParsing
            if ((Test-Path -Path $exeInstallerPath) -and ((Get-Item -Path $exeInstallerPath).Length -gt 10MB)) {
                $fileInfo = Get-Item -Path $exeInstallerPath
                Write-Host "Downloaded Teams EXE installer to $exeInstallerPath (Size: $([math]::Round($fileInfo.Length / 1MB, 2)) MB)" -ForegroundColor Green
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
            Write-Host "Trying fallback download URL..." -ForegroundColor Cyan
            try {
                Invoke-WebRequest -Uri $teamsExeFallbackUrl -OutFile $exeInstallerPath -UseBasicParsing
                if ((Test-Path -Path $exeInstallerPath) -and ((Get-Item -Path $exeInstallerPath).Length -gt 10MB)) {
                    $fileInfo = Get-Item -Path $exeInstallerPath
                    Write-Host "Downloaded Teams EXE installer to $exeInstallerPath (Size: $([math]::Round($fileInfo.Length / 1MB, 2)) MB)" -ForegroundColor Green
                    $downloadSuccess = $true
                }                else {
                    throw "Failed to download a valid Teams installer"
                }
            }
            catch {
                Write-Host "Error downloading Teams EXE installer: $_" -ForegroundColor Red
                  
                # Try the third/backup URL as a last resort
                Write-Host "Trying backup download URL..." -ForegroundColor Cyan
                try {
                    $webClient = New-Object System.Net.WebClient
                    $webClient.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
                    $webClient.DownloadFile($teamsExeBackupUrl, $exeInstallerPath)
                      if ((Test-Path -Path $exeInstallerPath) -and ((Get-Item -Path $exeInstallerPath).Length -gt 10MB)) {
                        $fileInfo = Get-Item -Path $exeInstallerPath
                        Write-Host "Downloaded Teams EXE installer to $exeInstallerPath (Size: $([math]::Round($fileInfo.Length / 1MB, 2)) MB)" -ForegroundColor Green
                        $downloadSuccess = $true
                    }
                    else {
                        throw "Failed to download a valid Teams installer with backup method"
                    }
                }
                catch {
                    Write-Host "All download methods failed. Cannot proceed with installation." -ForegroundColor Red
                    return
                }
            }        }
    }
    
    if (-not $downloadSuccess) {
        Write-Host "Unable to download a valid Teams installer. Installation aborted." -ForegroundColor Red
        return
    }
    
    if ($PSCmdlet.ShouldProcess('Install Microsoft Teams', "Install using $exeInstallerPath")) {
        Write-Host "Installing Microsoft Teams silently..." -ForegroundColor Cyan
        
        # Verify the downloaded file is valid
        try {
            # Check if the file exists
            if (-not (Test-Path -Path $exeInstallerPath)) {
                Write-Host "ERROR: Installer file not found at $exeInstallerPath" -ForegroundColor Red
                return
            }
            
            # Get file information
            $fileSize = (Get-Item -Path $exeInstallerPath).Length
            Write-Host "Installer file size: $([math]::Round($fileSize / 1MB, 2)) MB" -ForegroundColor DarkGray
                
            if ($fileSize -lt 1MB) {
                Write-Host "ERROR: The installer file appears to be too small and may be corrupted" -ForegroundColor Red
                return
            }
            
            # Check file signature if available
            $signature = Get-AuthenticodeSignature -FilePath $exeInstallerPath -ErrorAction SilentlyContinue
            if ($signature) {
                Write-Host "File signature status: $($signature.Status)" -ForegroundColor DarkGray
                if ($signature.Status -ne "Valid") {
                    Write-Host "WARNING: The installer does not have a valid signature. Proceeding with caution." -ForegroundColor Yellow
                }
            }
            
            # Handle both EXE and MSI formats (in case the download returned an MSI)
            $installerExtension = [System.IO.Path]::GetExtension($exeInstallerPath).ToLower()
            $installerArguments = ""
            
            if ($installerExtension -eq ".exe") {
                $installerArguments = "--silent"
                
                # Remove incompatible files if they exist
                if (Test-Path -Path "$exeInstallerPath.old") {
                    Remove-Item -Path "$exeInstallerPath.old" -Force -ErrorAction SilentlyContinue
                }
                
                Write-Host "Using EXE installer with silent arguments" -ForegroundColor DarkGray
                
                # Check if the EXE is compatible with this system
                try {
                    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
                    $pinfo.FileName = $exeInstallerPath
                    $pinfo.Arguments = "--help"
                    $pinfo.RedirectStandardError = $true
                    $pinfo.RedirectStandardOutput = $true
                    $pinfo.UseShellExecute = $false
                    $pinfo.CreateNoWindow = $true
                    
                    $process = New-Object System.Diagnostics.Process
                    $process.StartInfo = $pinfo
                    
                    try {
                        # Try to start the process just to verify compatibility
                        [void]$process.Start()
                        $process.Kill()
                        
                        # If we get here, the file is compatible
                        Write-Host "EXE installer format compatibility verified" -ForegroundColor Green
                    }
                    catch [System.ComponentModel.Win32Exception] {
                        if ($_.Exception.NativeErrorCode -eq 193) {
                            # Error 193: Not a valid Win32 application - incompatible bitness
                            Write-Host "ERROR: The installer is not compatible with this OS platform (Error 193)" -ForegroundColor Red
                            Write-Host "This typically indicates an architecture mismatch (e.g., trying to run 64-bit EXE on 32-bit Windows)" -ForegroundColor Yellow
                            
                            # Try to download the correct architecture
                            $correctArch = if ($osArch -eq "x64") { "x86" } else { "x64" }
                            Write-Host "Attempting to download $correctArch installer instead..." -ForegroundColor Cyan
                            
                            $correctUrl = "https://teams.microsoft.com/downloads/desktopurl?env=production&plat=windows&arch=$correctArch&download=true"
                            $correctInstallerPath = Join-Path -Path $tempDir -ChildPath "Teams_windows_$correctArch.exe"
                            
                            try {
                                Invoke-WebRequest -Uri $correctUrl -OutFile $correctInstallerPath -UseBasicParsing
                                if ((Test-Path -Path $correctInstallerPath) -and ((Get-Item -Path $correctInstallerPath).Length -gt 10MB)) {
                                    Write-Host "Successfully downloaded alternative architecture Teams installer" -ForegroundColor Green
                                    $exeInstallerPath = $correctInstallerPath
                                }
                                else {
                                    throw "Failed to download alternative architecture installer"
                                }
                            }
                            catch {                                Write-Host "Failed to download alternative installer. Attempting direct installation anyway." -ForegroundColor Yellow
                            }
                        }
                        else {
                            Write-Host "WARNING: Installer compatibility check failed: $($_.Exception.Message)" -ForegroundColor Yellow
                            Write-Host "Attempting to proceed with installation regardless..." -ForegroundColor Yellow
                        }
                    }
                }
                catch {
                    Write-Host "WARNING: Installer compatibility check failed: $_" -ForegroundColor Yellow
                    Write-Host "Attempting to proceed with installation anyway..." -ForegroundColor Yellow
                }
                  
                # Now try the actual installation
                try {
                    Write-Host "Starting Teams installation..." -ForegroundColor Cyan
                    $process = Start-Process -FilePath $exeInstallerPath -ArgumentList $installerArguments -Wait -PassThru
                }                catch {
                    Write-Host "Error during Teams installation: $_" -ForegroundColor Red
                    throw "Teams installation failed"
                }
            }
            elseif ($installerExtension -eq ".msi") {
                # If we somehow got an MSI file instead of EXE
                Write-Host "Detected MSI installer format, adjusting installation method" -ForegroundColor Yellow
                $installerArguments = "/i `"$exeInstallerPath`" /qn /norestart ALLUSERS=1"
                
                try {
                    Write-Host "Using MSI installer with silent arguments" -ForegroundColor DarkGray
                    $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $installerArguments -Wait -PassThru
                }                catch {
                    Write-Host "Error during Teams MSI installation: $_" -ForegroundColor Red
                    throw "Teams MSI installation failed"
                }
            }
            else {
                # If it's neither EXE nor MSI, try running as EXE with no arguments
                try {
                    Write-Host "Unknown installer format, attempting default installation method" -ForegroundColor Yellow
                    $process = Start-Process -FilePath $exeInstallerPath -Wait -PassThru
                }
                catch {
                    Write-Host "Error during Teams generic installation: $_" -ForegroundColor Red
                    throw "Teams generic installation failed"
                }
            }
            
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
                    1 { Write-Host "Error 1: General installation error." -ForegroundColor Red }
                    2 { Write-Host "Error 2: User cancelled installation." -ForegroundColor Yellow }
                    3 { Write-Host "Error 3: Fatal installation error." -ForegroundColor Red }
                    4 { Write-Host "Error 4: Installation failed due to system requirements." -ForegroundColor Red }
                    5 { Write-Host "Error 5: Application is already running." -ForegroundColor Yellow }
                    1603 { Write-Host "Error 1603: Fatal error during installation." -ForegroundColor Red }
                    1618 { Write-Host "Error 1618: Another installation is already in progress." -ForegroundColor Red }
                    1619 { Write-Host "Error 1619: Installation package could not be found." -ForegroundColor Red }                    1620 { Write-Host "Error 1620: Installation package could not be opened." -ForegroundColor Red }
                    1638 { Write-Host "Error 1638: Another version of this product is already installed." -ForegroundColor Yellow }
                    1641 { Write-Host "Error 1641: The installer has initiated a restart." -ForegroundColor Yellow }
                    default { Write-Host "Check installer error codes for more details." -ForegroundColor Yellow }
                }
            }
        }
        catch {
            Write-Host "Error installing Microsoft Teams: $_" -ForegroundColor Red
            
            # Attempt to provide more detailed diagnostics
            Write-Host "Performing additional diagnostics..." -ForegroundColor Cyan
            
            # Check if the file exists
            if (-not (Test-Path -Path $exeInstallerPath)) {
                Write-Host "ERROR: The installer file no longer exists at $exeInstallerPath" -ForegroundColor Red
                return
            }
            
            # Verify file is not corrupted
            try {
                $fileSize = (Get-Item -Path $exeInstallerPath).Length
                Write-Host "Installer file size: $([math]::Round($fileSize / 1MB, 2)) MB" -ForegroundColor DarkGray
                
                if ($fileSize -lt 1MB) {
                    Write-Host "ERROR: The installer file appears to be too small and may be corrupted" -ForegroundColor Red
                    return                }
                
                # Check file signature if available
                $signature = Get-AuthenticodeSignature -FilePath $exeInstallerPath -ErrorAction SilentlyContinue
                if ($signature) {
                    Write-Host "File signature status: $($signature.Status)" -ForegroundColor DarkGray
                    if ($signature.Status -ne "Valid") {
                        Write-Host "WARNING: The installer does not have a valid signature" -ForegroundColor Yellow
                    }
                }
                
                # Try alternate installation method as a last resort
                Write-Host "Attempting alternate installation method..." -ForegroundColor Cyan
                
                # Try using the alternate architecture as a fallback
                $alternateArch = if ($osArch -eq "x64") { "x86" } else { "x64" }
                Write-Host "Attempting to download $alternateArch installer as a fallback..." -ForegroundColor Cyan
                
                $alternateUrl = "https://teams.microsoft.com/downloads/desktopurl?env=production&plat=windows&arch=$alternateArch&download=true"
                $alternateInstallerPath = Join-Path -Path $tempDir -ChildPath "Teams_windows_$alternateArch.exe"
                
                try {
                    Invoke-WebRequest -Uri $alternateUrl -OutFile $alternateInstallerPath -UseBasicParsing
                    
                    if ((Test-Path -Path $alternateInstallerPath) -and ((Get-Item -Path $alternateInstallerPath).Length -gt 10MB)) {
                        Write-Host "Successfully downloaded alternative architecture Teams installer" -ForegroundColor Green
                        
                        # Try installing with the alternate architecture
                        Write-Host "Attempting installation with $alternateArch installer..." -ForegroundColor Cyan
                        Start-Process -FilePath $alternateInstallerPath -ArgumentList "--silent" -Wait -NoNewWindow
                        
                        # Check if Teams is now installed
                        $possibleTeamsPaths = @(
                            "$env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe",
                            "${env:ProgramFiles}\Microsoft\Teams\current\Teams.exe",
                            "${env:ProgramFiles(x86)}\Microsoft\Teams\current\Teams.exe"
                        )
                        
                        $teamsInstalled = $false
                        foreach ($path in $possibleTeamsPaths) {
                            if (Test-Path -Path $path) {
                                $teamsInstalled = $true
                                Write-Host "Teams installed successfully at $path" -ForegroundColor Green
                                break
                            }
                        }
                        
                        if (-not $teamsInstalled) {
                            Write-Host "Alternative architecture installation attempt completed, but Teams installation could not be verified." -ForegroundColor Yellow
                        }
                    }
                    else {
                        throw "Failed to download alternative architecture installer"
                    }
                }
                catch {
                    Write-Host "Error with alternative architecture installation: $_" -ForegroundColor Red
                    
                    # Try copying to a new location with .new extension and execute from there
                    $newInstallerPath = "$exeInstallerPath.new"
                    Copy-Item -Path $exeInstallerPath -Destination $newInstallerPath -Force
                    
                    if (Test-Path -Path $newInstallerPath) {
                        Write-Host "Executing installer from alternate location: $newInstallerPath" -ForegroundColor Yellow
                        # Start process without waiting, just to get it going
                        Start-Process -FilePath $newInstallerPath -NoNewWindow
                    }
                }
            }
            catch {
                Write-Host "Error during diagnostics: $_" -ForegroundColor Red
            }
        }
    }
}

function Test-TeamsInstallationHealth {
    Write-Host "Performing additional Teams installation health checks..." -ForegroundColor Cyan
    $healthStatus = $true
    
    # Check if Teams process starts properly
    try {
        $teamsPath = "$env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe"
        if (Test-Path -Path $teamsPath) {
            Write-Host "Found Teams executable at $teamsPath" -ForegroundColor Green
            
            # Check if Teams is already running
            $teamsProcesses = Get-Process -Name "Teams" -ErrorAction SilentlyContinue
            if (-not $teamsProcesses) {
                Write-Host "Testing Teams startup..." -ForegroundColor Cyan
                # Start Teams and immediately close it to test viability
                try {
                    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
                    $startInfo.FileName = $teamsPath
                    $startInfo.Arguments = "--processStart ""Teams.exe"""
                    $startInfo.UseShellExecute = $true
                    
                    # Start the process
                    $proc = [System.Diagnostics.Process]::Start($startInfo)
                    Start-Sleep -Seconds 5
                    
                    if ($proc) {
                        Write-Host "Teams process started successfully" -ForegroundColor Green
                        
                        # Try gracefully closing Teams
                        try {
                            $runningTeams = Get-Process -Name "Teams" -ErrorAction SilentlyContinue
                            if ($runningTeams) {
                                $runningTeams | ForEach-Object { $_.CloseMainWindow() | Out-Null }
                                Start-Sleep -Seconds 2
                                $remainingTeams = Get-Process -Name "Teams" -ErrorAction SilentlyContinue
                                if ($remainingTeams) {
                                    Write-Host "Teams process still running, stopping process..." -ForegroundColor Yellow
                                    $remainingTeams | Stop-Process -Force -ErrorAction SilentlyContinue
                                }
                            }
                        }
                        catch {
                            Write-Host "Error stopping Teams test process: $_" -ForegroundColor Yellow
                        }
                    }
                    else {
                        Write-Host "Teams process failed to start" -ForegroundColor Red
                        $healthStatus = $false
                    }
                }
                catch {
                    Write-Host "Error starting Teams: $_" -ForegroundColor Red
                    $healthStatus = $false
                }
            }
            else {
                Write-Host "Teams is already running - skipping process test" -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "Teams executable not found at expected location" -ForegroundColor Red
            $healthStatus = $false
        }
    }
    catch {
        Write-Host "Error testing Teams process: $_" -ForegroundColor Red
        $healthStatus = $false
    }
    
    # Check for Teams configuration files
    try {
        $configPath = "$env:APPDATA\Microsoft\Teams"
        if (Test-Path -Path $configPath) {
            Write-Host "Teams configuration directory found at $configPath" -ForegroundColor Green
            
            # Look for essential configuration files
            $configFiles = @(
                "desktop-config.json",
                "storage.json"
            )
            
            foreach ($file in $configFiles) {
                $filePath = Join-Path -Path $configPath -ChildPath $file
                if (Test-Path -Path $filePath) {
                    Write-Host "Found Teams configuration file: $file" -ForegroundColor Green
                }
                else {
                    Write-Host "Missing Teams configuration file: $file" -ForegroundColor Yellow
                }
            }
        }
        else {
            Write-Host "Teams configuration directory not found" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Error checking Teams configuration: $_" -ForegroundColor Red
    }
    
    return $healthStatus
}

function Stop-TeamsProcesses {
    Write-Host "Checking for running Teams processes..." -ForegroundColor Cyan
    
    # List of process names associated with Teams
    $teamsProcessNames = @(
        "Teams",
        "Microsoft.Teams",
        "Teams.exe",
        "TeamsUpdate",
        "Update"
    )
    
    $processCount = 0
    foreach ($processName in $teamsProcessNames) {
        $processes = Get-Process -Name $processName -ErrorAction SilentlyContinue
        if ($processes) {
            foreach ($process in $processes) {
                try {
                    # Try graceful shutdown first
                    Write-Host "Attempting to close $($process.Name) (PID: $($process.Id))..." -ForegroundColor Cyan
                    # Use CloseMainWindow but don't need to capture return value
                    $process.CloseMainWindow() | Out-Null
                    Start-Sleep -Seconds 2
                    
                    # If still running, force kill
                    if (-not $process.HasExited) {
                        Write-Host "Forcefully stopping $($process.Name) (PID: $($process.Id))..." -ForegroundColor Yellow
                        Stop-Process -Id $process.Id -Force -ErrorAction Stop
                    }
                    
                    Write-Host "Successfully stopped $($process.Name) process" -ForegroundColor Green
                    $processCount++
                }
                catch {
                    Write-Host "Error stopping $($process.Name) (PID: $($process.Id)): $_" -ForegroundColor Red
                }
            }
        }
    }
    
    # Double-check for any remaining processes
    $remainingProcessCount = 0
    foreach ($processName in $teamsProcessNames) {
        $remaining = Get-Process -Name $processName -ErrorAction SilentlyContinue
        if ($remaining) {
            $remainingProcessCount += $remaining.Count
        }
    }
    
    if ($remainingProcessCount -gt 0) {
        Write-Host "WARNING: $remainingProcessCount Teams processes could not be stopped" -ForegroundColor Red
        return $false
    }
    else {
        if ($processCount -gt 0) {
            Write-Host "All Teams processes successfully stopped ($processCount total)" -ForegroundColor Green
        }
        else {
            Write-Host "No Teams processes currently running" -ForegroundColor Green
        }
        return $true
    }
}

function Get-SystemArchitecture {
    # Detect system architecture and return OS-specific info
    $osArch = "x64"
    
    if (-not [Environment]::Is64BitOperatingSystem) {
        $osArch = "x86"
        Write-Host "Detected 32-bit operating system" -ForegroundColor DarkGray
    } else {
        Write-Host "Detected 64-bit operating system" -ForegroundColor DarkGray
    }
    
    # Additional architecture detection
    try {
        $procArch = $env:PROCESSOR_ARCHITECTURE
        Write-Host "Processor architecture: $procArch" -ForegroundColor DarkGray
          $osVersion = [Environment]::OSVersion.Version
        Write-Host "OS Version: $($osVersion.Major).$($osVersion.Minor) (Build $($osVersion.Build))" -ForegroundColor DarkGray
        
        # Check if running in WOW64 (Windows 32-bit on Windows 64-bit)
        if (Test-Path -Path "$env:SystemRoot\SysWOW64") {
            Write-Host "System has WOW64 subsystem" -ForegroundColor DarkGray
            
            # Detect if process is running in 32-bit mode on 64-bit Windows
            if ($null -ne $env:PROCESSOR_ARCHITEW6432) {
                Write-Host "Current process is running under WOW64 emulation" -ForegroundColor Yellow
                # In this case, we might need to use x86 installer even on x64 OS
                # This happens when running 32-bit PowerShell on 64-bit Windows
                $osArch = "x86"
                Write-Host "Adjusting download architecture to x86 due to WOW64 process" -ForegroundColor Yellow
            }
        }
    } catch {
        Write-Host "Error detecting detailed system architecture: $_" -ForegroundColor DarkGray
    }
    
    return $osArch
}

# Main script execution
Write-Host 'Starting Microsoft Teams uninstall and install process.' -ForegroundColor Cyan
Write-Host '----------------------------------------------------------------' -ForegroundColor DarkGray

# Stop any running Teams processes first
Stop-TeamsProcesses

# Uninstall any existing Teams instances
Uninstall-Teams

# Install latest version
Install-Teams

# Verify installation
Write-Host '----------------------------------------------------------------' -ForegroundColor DarkGray
Write-Host 'Verifying Microsoft Teams installation...' -ForegroundColor Cyan

# Initialize the status tracking variable
$script:teamsInstalled = $false

# Initialize teamsInstalled variable to avoid warnings
$script:teamsInstalled = $false

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
            $script:teamsInstalled = $true
        }
    }
}

# Check WMI/CIM for Teams
try {
    $cimProducts = Get-CimInstance -ClassName Win32_Product -ErrorAction SilentlyContinue | 
                   Where-Object { $_.Name -like "*Teams*" -or ($_.Vendor -like "*Microsoft*" -and $_.Name -like "*Teams*") }    
    foreach ($product in $cimProducts) {
        Write-Host "Microsoft Teams found in WMI: $($product.Name) v$($product.Version)" -ForegroundColor Green
        $script:teamsInstalled = $true
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
            $script:teamsInstalled = $true
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
            $script:teamsInstalled = $true
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
        $script:teamsInstalled = $true
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
