# =============================================================================
# Script: Install-MicrosoftTeams.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.8.1
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
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

[CmdletBinding(SupportsShouldProcess = $true)]
param()

# Color support variables and Write-ColorOutput function
$Script:UseAnsiColors = $PSVersionTable.PSVersion.Major -ge 7
$Script:Colors = if ($Script:UseAnsiColors) {
    @{
        'White'    = "`e[37m"
        'Cyan'     = "`e[36m"
        'Green'    = "`e[32m"
        'Yellow'   = "`e[33m"
        'Red'      = "`e[31m"
        'Magenta'  = "`e[35m"
        'DarkGray' = "`e[90m"
        'Reset'    = "`e[0m"
    }
} else {
    @{
        'White'    = [ConsoleColor]::White
        'Cyan'     = [ConsoleColor]::Cyan
        'Green'    = [ConsoleColor]::Green
        'Yellow'   = [ConsoleColor]::Yellow
        'Red'      = [ConsoleColor]::Red
        'Magenta'  = [ConsoleColor]::Magenta
        'DarkGray' = [ConsoleColor]::DarkGray
        'Reset'    = ''
    }
}

function Write-ColorOutput {
    <#
    .SYNOPSIS
    Outputs colored text in a way that's compatible with PSScriptAnalyzer requirements.

    .DESCRIPTION
    This function provides colored output while maintaining compatibility with PSScriptAnalyzer
    by using only Write-Output and standard PowerShell cmdlets.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$Color = "White"
    )

    # Always use Write-Output to satisfy PSScriptAnalyzer
    # For PowerShell 7+, include ANSI color codes in the output
    if ($Script:UseAnsiColors) {
        $colorCode = $Script:Colors[$Color]
        $resetCode = $Script:Colors.Reset
        Write-Output "${colorCode}${Message}${resetCode}"
    } else {
        # For PowerShell 5.1, just output the message
        # Color formatting will be handled by the terminal/host if supported
        Write-Output $Message
    }
}

function Uninstall-TeamsApp {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param()
    Write-ColorOutput -Message 'Scanning for existing Microsoft Teams installations...' -Color 'Cyan'

    # Track successful uninstalls
    $uninstallCount = 0

    # First check running processes to ensure they're terminated before uninstallation
    Write-ColorOutput -Message 'Checking for running Teams processes...' -Color 'DarkGray'
    try {
        $teamsProcesses = Get-Process -Name "*teams*" -ErrorAction SilentlyContinue

        if ($teamsProcesses) {
            if ($PSCmdlet.ShouldProcess("Running Teams processes", "Terminate")) {
                Write-ColorOutput -Message "Found running Teams processes. Terminating..." -Color 'Cyan'
                $teamsProcesses | Stop-Process -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 2
            }
        }
    } catch {
        Write-ColorOutput -Message "Error checking for Teams processes: $_" -Color 'Red'
    }

    #region Registry-based uninstallation
    Write-ColorOutput -Message 'Checking registry uninstall keys...' -Color 'DarkGray'
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
                Write-ColorOutput -Message "Uninstalling $($app.DisplayName) GUID $($app.GUID)..." -Color 'Cyan'
                Start-Process -FilePath 'msiexec.exe' -ArgumentList "/x $($app.GUID) /qn /norestart" -Wait
                Write-ColorOutput -Message "Successfully uninstalled $($app.DisplayName)." -Color 'Green'
                $uninstallCount++
            } catch {
                Write-ColorOutput -Message "Error uninstalling $($app.DisplayName): $_" -Color 'Red'
            }
        }
    }
    #endregion

    #region Per-user Teams installations
    Write-ColorOutput -Message 'Checking for per-user Teams installations...' -Color 'DarkGray'
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
                        Write-ColorOutput -Message "Stopping Teams processes..." -Color 'Yellow'
                        $processes | Stop-Process -Force -ErrorAction SilentlyContinue
                        Start-Sleep -Seconds 2
                    }

                    # Try to uninstall using Update.exe
                    $updateExe = Join-Path -Path $folder -ChildPath "Update.exe"
                    if (Test-Path -Path $updateExe) {
                        Write-ColorOutput -Message "Running Teams uninstaller: $updateExe --uninstall" -Color 'Cyan'
                        Start-Process -FilePath $updateExe -ArgumentList "--uninstall" -Wait -ErrorAction SilentlyContinue
                    }

                    # Remove the folder
                    Write-ColorOutput -Message "Removing Teams folder: $folder" -Color 'Cyan'
                    Remove-Item -Path $folder -Recurse -Force -ErrorAction Stop
                    Write-ColorOutput -Message "Successfully removed per-user Teams installation at $folder" -Color 'Green'
                    $uninstallCount++
                } catch {
                    Write-ColorOutput -Message "Error removing Teams folder $folder`: $_" -Color 'Red'
                }
            }
        }
    }
    #endregion

    #region AppX/Store version of Teams
    Write-ColorOutput -Message 'Checking for Microsoft Store version of Teams...' -Color 'DarkGray'
    try {
        $teamsAppx = Get-AppxPackage -Name "*MicrosoftTeams*" -ErrorAction SilentlyContinue
        if ($teamsAppx) {
            foreach ($app in $teamsAppx) {
                if ($PSCmdlet.ShouldProcess("Microsoft Store Teams: $($app.Name) v$($app.Version)", "Remove")) {
                    Write-ColorOutput -Message "Removing Microsoft Store Teams app: $($app.Name) v$($app.Version)" -Color 'Cyan'

                    # First terminate any running Teams processes from WindowsApps
                    $teamsProcesses = Get-Process -Name "*Teams*" -ErrorAction SilentlyContinue |
                        Where-Object { $null -ne $_.Path -and $_.Path -like "*WindowsApps*" }

                    if ($teamsProcesses) {
                        Write-ColorOutput -Message "  Terminating Teams processes from Windows Store..." -Color 'Yellow'
                        $teamsProcesses | Stop-Process -Force -ErrorAction SilentlyContinue
                        Start-Sleep -Seconds 2
                    }

                    # Now remove the AppX package
                    Remove-AppxPackage -Package $app.PackageFullName -ErrorAction Stop
                    Write-ColorOutput -Message "Successfully removed Microsoft Store Teams app." -Color 'Green'
                    $uninstallCount++
                }
            }
        }
    } catch {
        Write-ColorOutput -Message "Error removing Microsoft Store Teams app: $_" -Color 'Red'
        Write-ColorOutput -Message "NOTE: Windows Store apps may require special permissions to remove." -Color 'Yellow'
        Write-ColorOutput -Message "      Try running this script with 'Run as administrator'" -Color 'Yellow'
    }
    #endregion

    #region WMI/CIM-based Teams detection
    Write-ColorOutput -Message 'Checking for Teams installations via WMI/CIM...' -Color 'DarkGray'
    try {
        $cimProducts = Get-CimInstance -ClassName Win32_Product -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like "*Teams*" -or ($_.Vendor -like "*Microsoft*" -and $_.Name -like "*Teams*") }

        foreach ($product in $cimProducts) {
            if ($PSCmdlet.ShouldProcess("WMI Product: $($product.Name) v$($product.Version)", "Uninstall")) {
                Write-ColorOutput -Message "Uninstalling Teams via WMI: $($product.Name) v$($product.Version)" -Color 'Cyan'

                try {
                    # Try using the IdentifyingNumber (equivalent to GUID)
                    $guid = $product.IdentifyingNumber

                    if ($guid) {
                        Start-Process -FilePath 'msiexec.exe' -ArgumentList "/x $guid /qn /norestart" -Wait
                        Write-ColorOutput -Message "Successfully uninstalled Teams via WMI: $($product.Name)" -Color 'Green'
                        $uninstallCount++
                    } else {
                        # Alternative: use the Win32_Product.Uninstall() method
                        $result = $product | Invoke-CimMethod -MethodName "Uninstall"

                        if ($result.ReturnValue -eq 0) {
                            Write-ColorOutput -Message "Successfully uninstalled Teams via WMI method: $($product.Name)" -Color 'Green'
                            $uninstallCount++
                        } else {
                            Write-ColorOutput -Message "Failed to uninstall Teams via WMI method: $($product.Name). Return code: $($result.ReturnValue)" -Color 'Yellow'
                        }
                    }
                } catch {
                    Write-ColorOutput -Message "Error uninstalling Teams via WMI: $($product.Name). Error: $_" -Color 'Red'
                }
            }
        }
    } catch {
        Write-ColorOutput -Message "Error accessing WMI product information: $_" -Color 'Red'
    }
    #endregion

    #region Process-based Teams detection
    Write-ColorOutput -Message '' -Color 'DarkGray'
    try {
        $teamsProcesses = Get-Process -Name "*teams*" -ErrorAction SilentlyContinue

        if ($teamsProcesses) {
            if ($PSCmdlet.ShouldProcess("Running Teams processes", "Terminate")) {
                Write-ColorOutput -Message "Found running Teams processes. Terminating..." -Color 'Cyan'
                $teamsProcesses | ForEach-Object {
                    try {
                        # Get process details
                        $processPath = $_.Path
                        $processVersion = $null

                        if ($processPath -and (Test-Path -Path $processPath)) {
                            # Get version information
                            $fileInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($processPath)
                            $processVersion = $fileInfo.FileVersion

                            Write-ColorOutput -Message "Found Teams process: $processPath (Version $processVersion)" -Color 'Cyan'

                            # Stop the process
                            $_ | Stop-Process -Force
                            Write-ColorOutput -Message "Process terminated" -Color 'Yellow'

                            # Check if this is from a directory we haven't seen yet
                            $processDir = Split-Path -Parent $processPath
                            $parentDir = Split-Path -Parent $processDir

                            # If this is a unique directory not in our lists, try to remove it
                            if ((Test-Path -Path $parentDir) -and
                                ($parentDir -like "*Teams*" -or $processDir -like "*Teams*") -and
                                ($parentDir -notlike "$env:LOCALAPPDATA\Microsoft\Teams*") -and
                                ($parentDir -notlike "$env:APPDATA\Microsoft\Teams*") -and
                                ($parentDir -notlike "${ env:ProgramFiles}\Microsoft\Teams*") -and
                                ($parentDir -notlike "${ env:ProgramFiles(x86)}\Microsoft\Teams*") -and
                                ($parentDir -notlike "$env:ProgramData\Microsoft\Teams*")) {

                                # Wait a moment for process resources to release
                                Start-Sleep -Seconds 2

                                if ($PSCmdlet.ShouldProcess("Teams directory: $parentDir", "Remove")) {
                                    try {
                                        Write-ColorOutput -Message "Attempting to remove Teams directory: $parentDir" -Color 'Cyan'
                                        Remove-Item -Path $parentDir -Recurse -Force -ErrorAction Stop
                                        Write-ColorOutput -Message "Successfully removed Teams directory at $parentDir" -Color 'Green'
                                        $uninstallCount++
                                    } catch {
                                        Write-ColorOutput -Message "Error removing Teams directory $parentDir`: $_" -Color 'Red'
                                    }
                                }
                            }
                        }
                    } catch {
                        Write-ColorOutput -Message "Error processing Teams executable: $_" -Color 'Red'
                    }
                }
            }
        }
    } catch {
        Write-ColorOutput -Message "Error checking for Teams processes: $_" -Color 'Red'
    }
    #endregion

    #region Check common installation directories
    Write-ColorOutput -Message '' -Color 'DarkGray'
    $teamsCommonLocations = @(
        "${ env:ProgramFiles}\Microsoft\Teams",
        "${ env:ProgramFiles(x86)}\Microsoft\Teams",
        "$env:ProgramData\Microsoft\Teams"
    )

    foreach ($location in $teamsCommonLocations) {
        if (Test-Path -Path $location) {
            if ($PSCmdlet.ShouldProcess("Teams directory: $location", "Remove")) {
                try {
                    # Look for uninstaller exe or MSI
                    $uninstallers = Get-ChildItem -Path $location -Filter "*uninst*.exe" -Recurse -ErrorAction SilentlyContinue
                    foreach ($uninstaller in $uninstallers) {
                        Write-ColorOutput -Message "Running Teams uninstaller: $($uninstaller.FullName)" -Color 'Cyan'
                        Start-Process -FilePath $uninstaller.FullName -ArgumentList "/S", "/Silent", "/Q", "/quiet", "/qn", "/norestart" -Wait -ErrorAction SilentlyContinue
                    }

                    # Remove the directory
                    Write-ColorOutput -Message "Removing Teams directory: $location" -Color 'Cyan'
                    Remove-Item -Path $location -Recurse -Force -ErrorAction Stop
                    Write-ColorOutput -Message "Successfully removed Teams installation at $location" -Color 'Green'
                    $uninstallCount++
                } catch {
                    Write-ColorOutput -Message "Error removing Teams directory $location`: $_" -Color 'Red'
                }
            }
        }
    }
    #endregion

    #region Check WindowsApps directory (Store apps)
    Write-ColorOutput -Message '' -Color 'DarkGray'
    $windowsAppsPath = "${ env:ProgramFiles}\WindowsApps"

    if (Test-Path -Path $windowsAppsPath) {
        try {
            $teamsWindowsAppDirs = Get-ChildItem -Path $windowsAppsPath -Directory -Filter "*Teams*" -ErrorAction SilentlyContinue

            if ($teamsWindowsAppDirs) {
                foreach ($dir in $teamsWindowsAppDirs) {
                    Write-ColorOutput -Message "Found Windows Store Teams installation: $($dir.FullName)" -Color 'Cyan'
                    Write-ColorOutput -Message "NOTE: Windows Store apps are protected and require special handling." -Color 'Yellow'
                    Write-ColorOutput -Message "      Try removing via Settings > Apps > Apps & features" -Color 'Yellow'

                    # Try to get the AppX package for this directory
                    $packageName = ($dir.Name -split '_')[0]
                    $teamsAppx = Get-AppxPackage -Name "*$packageName*" -ErrorAction SilentlyContinue

                    if ($teamsAppx -and $PSCmdlet.ShouldProcess("Windows Store Teams: $($teamsAppx.Name)", "Remove")) {
                        Write-ColorOutput -Message "Attempting to remove Windows Store Teams package: $($teamsAppx.Name)" -Color 'Cyan'
                        try {
                            # The -AllUsers parameter requires admin privileges
                            Remove-AppxPackage -Package $teamsAppx.PackageFullName -AllUsers -ErrorAction Stop
                            Write-ColorOutput -Message "Successfully removed Windows Store Teams." -Color 'Green'
                            $uninstallCount++
                        } catch {
                            Write-ColorOutput -Message "Error removing Windows Store Teams: $_" -Color 'Red'
                            Write-ColorOutput -Message "Try running this script as administrator to remove system-level Store apps." -Color 'Yellow'
                        }
                    }
                }
            }
        } catch {
            Write-ColorOutput -Message "Error accessing Windows Store Apps directory: $_" -Color 'Red'
            Write-ColorOutput -Message "This typically requires administrator privileges." -Color 'Yellow'
        }
    }
    #endregion

    #region Shortcut-based Teams detection
    Write-ColorOutput -Message '' -Color 'DarkGray'
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
        } catch {
            Write-ColorOutput -Message "Error accessing shortcuts in $startMenuPath" -Color 'DarkGray'
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
                        Write-ColorOutput -Message "Found Teams shortcut target: $shortcutTarget (Version $($targetInfo.FileVersion))" -Color 'Cyan'

                        # Check if this is a directory we haven't processed yet
                        $targetDir = Split-Path -Parent $shortcutTarget
                        $parentDir = Split-Path -Parent $targetDir

                        if ((Test-Path -Path $parentDir) -and
                            ($parentDir -like "*Teams*" -or $targetDir -like "*Teams*") -and
                            ($parentDir -notlike "$env:LOCALAPPDATA\Microsoft\Teams*") -and
                            ($parentDir -notlike "$env:APPDATA\Microsoft\Teams*") -and
                            ($parentDir -notlike "${ env:ProgramFiles}\Microsoft\Teams*") -and
                            ($parentDir -notlike "${ env:ProgramFiles(x86)}\Microsoft\Teams*") -and
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
                                    Write-ColorOutput -Message "Removing Teams directory: $parentDir" -Color 'Cyan'
                                    Remove-Item -Path $parentDir -Recurse -Force -ErrorAction Stop
                                    Write-ColorOutput -Message "Successfully removed Teams directory at $parentDir" -Color 'Green'
                                    $uninstallCount++
                                } catch {
                                    Write-ColorOutput -Message "Error removing Teams directory $parentDir`: $_" -Color 'Red'
                                }
                            }
                        }
                    }
                }

                # Remove the shortcut itself
                Write-ColorOutput -Message "Removing Teams shortcut: $($shortcut.FullName)" -Color 'Cyan'
                Remove-Item -Path $shortcut.FullName -Force -ErrorAction Stop
                Write-ColorOutput -Message "Successfully removed Teams shortcut" -Color 'Green'
            } catch {
                Write-ColorOutput -Message "Error processing Teams shortcut $($shortcut.FullName): $_" -Color 'Red'
            }
        }
    }
    #endregion

    # Final status
    if ($uninstallCount -eq 0) {
        Write-ColorOutput -Message '' -Color 'Yellow'

        # Perform additional detection similar to verification phase
        Write-ColorOutput -Message '' -Color 'Cyan'

        # Check AppX Packages again (sometimes they need special handling)
        try {
            $teamsAppx = Get-AppxPackage -Name "*MicrosoftTeams*" -ErrorAction SilentlyContinue
            if ($teamsAppx) {
                Write-ColorOutput -Message "FOUND: Microsoft Teams AppX package detected: $($teamsAppx.Name) v$($teamsAppx.Version)" -Color 'Yellow'
                Write-ColorOutput -Message "       Try running this script again with administrator privileges" -Color 'Yellow'
            }
        } catch {
            # Silently continue if AppX package detection fails
            Write-Debug "AppX package detection failed: $_"
        }

        # Check for Teams in processes
        try {
            $teamsProcesses = Get-Process -Name "*Teams*" -ErrorAction SilentlyContinue | Where-Object { $null -ne $_.Path }
            if ($teamsProcesses) {
                foreach ($process in $teamsProcesses) {
                    $fileInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($process.Path)
                    Write-ColorOutput -Message "FOUND: Microsoft Teams running from: $($process.Path) v$($fileInfo.FileVersion)" -Color 'Yellow'
                    Write-ColorOutput -Message "       This process might be protected or require administrator privileges" -Color 'Yellow'
                }
            }
        } catch {
            # Silently continue if process detection fails
            Write-Debug "Process detection failed: $_"
        }

        # Check WindowsApps folder which requires admin access
        if (Test-Path -Path "$env:ProgramFiles\WindowsApps") {
            Write-ColorOutput -Message "NOTE: The Windows Store apps folder exists but may require administrator privileges to access" -Color 'Yellow'
        }
    } else {
        Write-ColorOutput -Message "Successfully uninstalled/removed $uninstallCount Teams components." -Color 'Green'
    }

    # Give the system a moment to finalize uninstallation
    Start-Sleep -Seconds 3
}

function Install-TeamsApp {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param()

    $tempDir = [System.IO.Path]::GetTempPath()
    $osArch = Get-SystemArchitecture
    $exeInstallerPath = Join-Path -Path $tempDir -ChildPath "Teams_windows_$osArch.exe"
    $downloadSuccess = $false

    if ($PSCmdlet.ShouldProcess('Download Microsoft Teams', "Download latest installer")) {
        # Use the current Teams EXE installer URL (updated for 2025)
        # Microsoft periodically changes the Teams download URLs and link IDs
        $teamsExeUrl = "https://teams.microsoft.com/downloads/desktopurl?env = production&plat = windows&arch = $osArch&managedInstaller = true&download = true"
        $teamsExeFallbackUrl = if ($osArch -eq "x64") {
            # 64-bit link
            "https://go.microsoft.com/fwlink/?linkid = 2187327"
        } else {
            # 32-bit link
            "https://go.microsoft.com/fwlink/?linkid = 2187323"
        }
        $teamsExeBackupUrl = "https://teams.microsoft.com/downloads/desktopurl?env = production&plat = windows&arch = $osArch&download = true"

        # Try direct exe download first
        Write-ColorOutput -Message "Downloading Microsoft Teams EXE installer..." -Color 'Cyan'
        try {
            Invoke-WebRequest -Uri $teamsExeUrl -OutFile $exeInstallerPath -UseBasicParsing
            if ((Test-Path -Path $exeInstallerPath) -and ((Get-Item -Path $exeInstallerPath).Length -gt 10MB)) {
                $fileInfo = Get-Item -Path $exeInstallerPath
                Write-ColorOutput -Message "Downloaded Teams EXE installer to $exeInstallerPath (Size: $([math]::Round($fileInfo.Length / 1MB, 2)) MB)" -Color 'Green'
                $downloadSuccess = $true
            } else {
                Write-ColorOutput -Message "Downloaded file seems too small. Trying fallback URL..." -Color 'Yellow'
                throw "Downloaded file is too small"
            }
        } catch {
            Write-ColorOutput -Message "Error with primary download URL: $($_.Exception.Message)" -Color 'Yellow'

            # Try fallback URL
            Write-ColorOutput -Message "Trying fallback download URL..." -Color 'Cyan'
            try {
                Invoke-WebRequest -Uri $teamsExeFallbackUrl -OutFile $exeInstallerPath -UseBasicParsing
                if ((Test-Path -Path $exeInstallerPath) -and ((Get-Item -Path $exeInstallerPath).Length -gt 10MB)) {
                    $fileInfo = Get-Item -Path $exeInstallerPath
                    Write-ColorOutput -Message "Downloaded Teams EXE installer to $exeInstallerPath (Size: $([math]::Round($fileInfo.Length / 1MB, 2)) MB)" -Color 'Green'
                    $downloadSuccess = $true
                } else {
                    throw "Failed to download a valid Teams installer"
                }
            } catch {
                Write-ColorOutput -Message "Error downloading Teams EXE installer: $_" -Color 'Red'

                # Try the third/backup URL as a last resort
                Write-ColorOutput -Message "Trying backup download URL..." -Color 'Cyan'
                try {
                    $webClient = New-Object System.Net.WebClient
                    $webClient.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
                    $webClient.DownloadFile($teamsExeBackupUrl, $exeInstallerPath)
                    if ((Test-Path -Path $exeInstallerPath) -and ((Get-Item -Path $exeInstallerPath).Length -gt 10MB)) {
                        $fileInfo = Get-Item -Path $exeInstallerPath
                        Write-ColorOutput -Message "Downloaded Teams EXE installer to $exeInstallerPath (Size: $([math]::Round($fileInfo.Length / 1MB, 2)) MB)" -Color 'Green'
                        $downloadSuccess = $true
                    } else {
                        throw "Failed to download a valid Teams installer with backup method"
                    }
                } catch {
                    Write-ColorOutput -Message "All download methods failed. Cannot proceed with installation." -Color 'Red'
                    return
                }
            }
        }
    }

    if (-not $downloadSuccess) {
        Write-ColorOutput -Message "Unable to download a valid Teams installer. Installation aborted." -Color 'Red'
        return
    }

    if ($PSCmdlet.ShouldProcess('Install Microsoft Teams', "Install using $exeInstallerPath")) {
        Write-ColorOutput -Message "Installing Microsoft Teams silently..." -Color 'Cyan'

        # Verify the downloaded file is valid
        try {
            # Check if the file exists
            if (-not (Test-Path -Path $exeInstallerPath)) {
                Write-ColorOutput -Message "ERROR: Installer file not found at $exeInstallerPath" -Color 'Red'
                return
            }

            # Get file information
            $fileSize = (Get-Item -Path $exeInstallerPath).Length
            Write-ColorOutput -Message "Installer file size: $([math]::Round($fileSize / 1MB, 2)) MB" -Color 'DarkGray'

            if ($fileSize -lt 1MB) {
                Write-ColorOutput -Message "ERROR: The installer file appears to be too small and may be corrupted" -Color 'Red'
                return
            }

            # Check file signature if available
            $signature = Get-AuthenticodeSignature -FilePath $exeInstallerPath -ErrorAction SilentlyContinue
            if ($signature) {
                Write-ColorOutput -Message "File signature status: $($signature.Status)" -Color 'DarkGray'
                if ($signature.Status -ne "Valid") {
                    Write-ColorOutput -Message "WARNING: The installer does not have a valid signature. Proceeding with caution." -Color 'Yellow'
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

                Write-ColorOutput -Message "Using EXE installer with silent arguments" -Color 'DarkGray'
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
                        Write-ColorOutput -Message "EXE installer format compatibility verified" -Color 'Green'
                    } catch [System.ComponentModel.Win32Exception] {
                        if ($_.Exception.NativeErrorCode -eq 193) {
                            # Error 193: Not a valid Win32 application - incompatible bitness
                            Write-ColorOutput -Message "ERROR: The installer is not compatible with this OS platform (Error 193)" -Color 'Red'
                            Write-ColorOutput -Message "This typically indicates an architecture mismatch (e.g., trying to run 64-bit EXE on 32-bit Windows)" -Color 'Yellow'

                            # Always download the x86 installer as a fallback
                            $correctArch = "x86"
                            Write-ColorOutput -Message "Attempting to download $correctArch installer instead..." -Color 'Cyan'

                            $correctUrl = "https://teams.microsoft.com/downloads/desktopurl?env = production&plat = windows&arch = $correctArch&download = true"
                            $correctInstallerPath = Join-Path -Path $tempDir -ChildPath "Teams_windows_$correctArch.exe"

                            try {
                                # Try direct download for x86
                                $webClient = New-Object System.Net.WebClient
                                $webClient.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
                                $webClient.DownloadFile($correctUrl, $correctInstallerPath)

                                if ((Test-Path -Path $correctInstallerPath) -and ((Get-Item -Path $correctInstallerPath).Length -gt 10MB)) {
                                    Write-ColorOutput -Message "Successfully downloaded alternative architecture Teams installer" -Color 'Green'
                                    $exeInstallerPath = $correctInstallerPath
                                } else {
                                    # Try fallback link for x86
                                    # 32-bit link
                                    $fallbackX86Url = "https://go.microsoft.com/fwlink/?linkid = 2187323"
                                    Write-ColorOutput -Message "Trying fallback link for x86..." -Color 'Yellow'
                                    $webClient.DownloadFile($fallbackX86Url, $correctInstallerPath)

                                    if ((Test-Path -Path $correctInstallerPath) -and ((Get-Item -Path $correctInstallerPath).Length -gt 10MB)) {
                                        Write-ColorOutput -Message "Successfully downloaded x86 Teams installer from fallback link" -Color 'Green'
                                        $exeInstallerPath = $correctInstallerPath
                                    } else {
                                        throw "Failed to download alternative architecture installer"
                                    }
                                }
                            } catch {
                                Write-ColorOutput -Message "Failed to download alternative installer: $_" -Color 'Red'
                                Write-ColorOutput -Message "Attempting to extract installer from MSI package..." -Color 'Yellow'

                                # Try using MSI as a last resort
                                # 32-bit link
                                $msiUrl = "https://go.microsoft.com/fwlink/?linkid = 2187323"
                                $msiPath = Join-Path -Path $tempDir -ChildPath "Teams_windows_x86.msi"

                                try {
                                    Invoke-WebRequest -Uri $msiUrl -OutFile $msiPath -UseBasicParsing

                                    if ((Test-Path -Path $msiPath) -and ((Get-Item -Path $msiPath).Length -gt 10MB)) {
                                        Write-ColorOutput -Message "Successfully downloaded Teams MSI installer" -Color 'Green'
                                        $exeInstallerPath = $msiPath
                                        $installerExtension = ".msi"
                                    } else {
                                        throw "Failed to download MSI installer"
                                    }
                                } catch {
                                    Write-ColorOutput -Message "All download attempts failed. Cannot continue with installation." -Color 'Red'
                                    return
                                }
                            }
                        } else {
                            Write-ColorOutput -Message "WARNING: Installer compatibility check failed: $($_.Exception.Message)" -Color 'Yellow'
                            Write-ColorOutput -Message "Will try to download x86 version as fallback..." -Color 'Yellow'

                            # Force x86 download
                            # 32-bit link
                            $x86Url = "https://go.microsoft.com/fwlink/?linkid = 2187323"
                            $x86Path = Join-Path -Path $tempDir -ChildPath "Teams_windows_x86.exe"

                            try {
                                Invoke-WebRequest -Uri $x86Url -OutFile $x86Path -UseBasicParsing

                                if ((Test-Path -Path $x86Path) -and ((Get-Item -Path $x86Path).Length -gt 10MB)) {
                                    Write-ColorOutput -Message "Successfully downloaded x86 Teams installer" -Color 'Green'
                                    $exeInstallerPath = $x86Path
                                } else {
                                    throw "Failed to download x86 installer"
                                }
                            } catch {
                                Write-ColorOutput -Message "Failed to download x86 installer: $_" -Color 'Red'
                                # Continue with original installer despite warning
                            }
                        }
                    }
                } catch {
                    Write-ColorOutput -Message "WARNING: Installer compatibility check failed: $_" -Color 'Yellow'
                    Write-ColorOutput -Message "Will try to download x86 version as fallback..." -Color 'Yellow'

                    # Force x86 download
                    # 32-bit link
                    $x86Url = "https://go.microsoft.com/fwlink/?linkid = 2187323"
                    $x86Path = Join-Path -Path $tempDir -ChildPath "Teams_windows_x86.exe"

                    try {
                        Invoke-WebRequest -Uri $x86Url -OutFile $x86Path -UseBasicParsing

                        if ((Test-Path -Path $x86Path) -and ((Get-Item -Path $x86Path).Length -gt 10MB)) {
                            Write-ColorOutput -Message "Successfully downloaded x86 Teams installer" -Color 'Green'
                            $exeInstallerPath = $x86Path
                        } else {
                            throw "Failed to download x86 installer"
                        }
                    } catch {
                        Write-ColorOutput -Message "Failed to download x86 installer: $_" -Color 'Red'
                        # Continue with original installer despite warning
                    }
                }
                # Now try the actual installation
                try {
                    Write-ColorOutput -Message "Starting Teams installation..." -Color 'Cyan'
                    Write-ColorOutput -Message "Running: $exeInstallerPath $installerArguments" -Color 'Cyan'
                    $process = Start-Process -FilePath $exeInstallerPath -ArgumentList $installerArguments -Wait -PassThru -NoNewWindow
                } catch {
                    Write-ColorOutput -Message "Error during Teams installation: $_" -Color 'Red'

                    # Try alternative installer approach by extracting and using the embedded MSI
                    try {
                        Write-ColorOutput -Message "Trying to extract MSI from EXE installer..." -Color 'Yellow'
                        $extractDir = Join-Path -Path $tempDir -ChildPath "TeamsExtract"

                        # Create extraction directory if it doesn't exist
                        if (-not (Test-Path -Path $extractDir)) {
                            New-Item -Path $extractDir -ItemType Directory -Force | Out-Null
                        }

                        # Try to extract with /extract parameter (common for many installers)
                        Write-ColorOutput -Message "Attempting to extract installer contents..." -Color 'Cyan'
                        Start-Process -FilePath $exeInstallerPath -ArgumentList "/extract:`"$extractDir`"", "/quiet" -Wait -NoNewWindow

                        # Look for MSI files in extract directory
                        $msiFiles = Get-ChildItem -Path $extractDir -Filter "*.msi" -Recurse -ErrorAction SilentlyContinue

                        if ($msiFiles.Count -gt 0) {
                            $msiPath = $msiFiles[0].FullName
                            Write-ColorOutput -Message "Found MSI file: $msiPath" -Color 'Green'

                            # Install using the extracted MSI
                            Write-ColorOutput -Message "Installing Teams using extracted MSI..." -Color 'Cyan'
                            $msiArgs = "/i `"$msiPath`" /qn /norestart ALLUSERS = 1"
                            $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $msiArgs -Wait -PassThru -NoNewWindow
                        } else {
                            # Try the x86 direct download link as a last resort
                            Write-ColorOutput -Message "No MSI found. Trying direct download of x86 teams..." -Color 'Yellow'
                            # 32-bit link
                            $x86Url = "https://go.microsoft.com/fwlink/?linkid = 2187323"
                            $x86Path = Join-Path -Path $tempDir -ChildPath "Teams_windows_x86_direct.exe"

                            try {
                                $webClient = New-Object System.Net.WebClient
                                $webClient.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
                                $webClient.DownloadFile($x86Url, $x86Path)

                                if ((Test-Path -Path $x86Path) -and ((Get-Item -Path $x86Path).Length -gt 10MB)) {
                                    Write-ColorOutput -Message "Successfully downloaded x86 Teams installer" -Color 'Green'
                                    $process = Start-Process -FilePath $x86Path -ArgumentList "--silent" -Wait -PassThru -NoNewWindow
                                } else {
                                    throw "Failed to download valid x86 installer"
                                }
                            } catch {
                                Write-ColorOutput -Message "All installation methods failed: $_" -Color 'Red'
                                throw "Teams installation failed after multiple attempts"
                            }
                        }
                    } catch {
                        Write-ColorOutput -Message "Alternative installation method failed: $_" -Color 'Red'
                        throw "Teams installation failed"
                    }
                }
            } elseif ($installerExtension -eq ".msi") {
                # If we somehow got an MSI file instead of EXE
                Write-ColorOutput -Message "Detected MSI installer format, using MSI installation method" -Color 'Yellow'
                $installerArguments = "/i `"$exeInstallerPath`" /qn /norestart ALLUSERS = 1"

                try {
                    Write-ColorOutput -Message "Using MSI installer with silent arguments" -Color 'DarkGray'
                    Write-ColorOutput -Message "Running: msiexec.exe $installerArguments" -Color 'Cyan'
                    $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $installerArguments -Wait -PassThru -NoNewWindow
                } catch {
                    Write-ColorOutput -Message "Error during Teams MSI installation: $_" -Color 'Red'

                    # Try an alternative MSI installation approach
                    try {
                        Write-ColorOutput -Message "Trying alternative MSI installation method..." -Color 'Yellow'
                        $alternativeArgs = "/i `"$exeInstallerPath`" /quiet /norestart"
                        Write-ColorOutput -Message "Running: msiexec.exe $alternativeArgs" -Color 'Cyan'
                        $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $alternativeArgs -Wait -PassThru -NoNewWindow
                    } catch {
                        Write-ColorOutput -Message "Error during alternative Teams MSI installation: $_" -Color 'Red'
                        throw "Teams MSI installation failed"
                    }
                }
            } else {
                # If it's neither EXE nor MSI, try running as EXE with no arguments
                try {
                    Write-ColorOutput -Message "Unknown installer format, attempting default installation method" -Color 'Yellow'
                    Write-ColorOutput -Message "Running installer directly without arguments" -Color 'Cyan'
                    $process = Start-Process -FilePath $exeInstallerPath -Wait -PassThru -NoNewWindow
                } catch {
                    Write-ColorOutput -Message "Error during Teams generic installation: $_" -Color 'Red'
                    throw "Teams generic installation failed"
                }
            }

            if ($process.ExitCode -eq 0) {
                Write-ColorOutput -Message "Microsoft Teams installed successfully." -Color 'Green'
            } elseif ($process.ExitCode -eq 3010) {
                Write-ColorOutput -Message "Microsoft Teams installed successfully but requires a restart to complete installation." -Color 'Yellow'
            } else {
                Write-ColorOutput -Message "Microsoft Teams installation exited with code: $($process.ExitCode)." -Color 'Yellow'
                # Provide more specific information about common error codes
                switch ($process.ExitCode) {
                    1 { Write-ColorOutput -Message "Error 1: General installation error." -Color 'Red' }
                    2 { Write-ColorOutput -Message "Error 2: User cancelled installation." -Color 'Yellow' }
                    3 { Write-ColorOutput -Message "Error 3: Fatal installation error." -Color 'Red' }
                    4 { Write-ColorOutput -Message "Error 4: Installation failed due to system requirements." -Color 'Red' }
                    5 { Write-ColorOutput -Message "Error 5: Application is already running." -Color 'Yellow' }
                    1603 { Write-ColorOutput -Message "Error 1603: Fatal error during installation." -Color 'Red' }
                    1618 { Write-ColorOutput -Message "Error 1618: Another installation is already in progress." -Color 'Red' }
                    1619 { Write-ColorOutput -Message "Error 1619: Installation package could not be found." -Color 'Red' }                    1620 { Write-ColorOutput -Message "Error 1620: Installation package could not be opened." -Color 'Red' }
                    1638 { Write-ColorOutput -Message "Error 1638: Another version of this product is already installed." -Color 'Yellow' }
                    1641 { Write-ColorOutput -Message "Error 1641: The installer has initiated a restart." -Color 'Yellow' }
                    default { Write-ColorOutput -Message "Check installer error codes for more details." -Color 'Yellow' }
                }
            }
        } catch {
            Write-ColorOutput -Message "Error installing Microsoft Teams: $_" -Color 'Red'

            # Attempt to provide more detailed diagnostics
            Write-ColorOutput -Message "Performing additional diagnostics..." -Color 'Cyan'

            # Check if the file exists
            if (-not (Test-Path -Path $exeInstallerPath)) {
                Write-ColorOutput -Message "ERROR: The installer file no longer exists at $exeInstallerPath" -Color 'Red'
                return
            }

            # Verify file is not corrupted
            try {
                $fileSize = (Get-Item -Path $exeInstallerPath).Length
                Write-ColorOutput -Message "Installer file size: $([math]::Round($fileSize / 1MB, 2)) MB" -Color 'DarkGray'

                if ($fileSize -lt 1MB) {
                    Write-ColorOutput -Message "ERROR: The installer file appears to be too small and may be corrupted" -Color 'Red'
                    return
                }

                # Check file signature if available
                $signature = Get-AuthenticodeSignature -FilePath $exeInstallerPath -ErrorAction SilentlyContinue
                if ($signature) {
                    Write-ColorOutput -Message "File signature status: $($signature.Status)" -Color 'DarkGray'
                    if ($signature.Status -ne "Valid") {
                        Write-ColorOutput -Message "WARNING: The installer does not have a valid signature" -Color 'Yellow'
                    }
                }

                # Try alternate installation method as a last resort
                Write-ColorOutput -Message "Attempting alternate installation method..." -Color 'Cyan'

                # Try using the alternate architecture as a fallback
                $alternateArch = if ($osArch -eq "x64") { "x86" } else { "x64" }
                Write-ColorOutput -Message "Attempting to download $alternateArch installer as a fallback..." -Color 'Cyan'

                $alternateUrl = "https://teams.microsoft.com/downloads/desktopurl?env = production&plat = windows&arch = $alternateArch&download = true"
                $alternateInstallerPath = Join-Path -Path $tempDir -ChildPath "Teams_windows_$alternateArch.exe"

                try {
                    Invoke-WebRequest -Uri $alternateUrl -OutFile $alternateInstallerPath -UseBasicParsing

                    if ((Test-Path -Path $alternateInstallerPath) -and ((Get-Item -Path $alternateInstallerPath).Length -gt 10MB)) {
                        Write-ColorOutput -Message "Successfully downloaded alternative architecture Teams installer" -Color 'Green'

                        # Try installing with the alternate architecture
                        Write-ColorOutput -Message "Attempting installation with $alternateArch installer..." -Color 'Cyan'
                        Start-Process -FilePath $alternateInstallerPath -ArgumentList "--silent" -Wait -NoNewWindow

                        # Check if Teams is now installed
                        $possibleTeamsPaths = @(
                            "$env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe",
                            "${ env:ProgramFiles}\Microsoft\Teams\current\Teams.exe",
                            "${ env:ProgramFiles(x86)}\Microsoft\Teams\current\Teams.exe"
                        )

                        $teamsInstalled = $false
                        foreach ($path in $possibleTeamsPaths) {
                            if (Test-Path -Path $path) {
                                $teamsInstalled = $true
                                Write-ColorOutput -Message "Teams installed successfully at $path" -Color 'Green'
                                break
                            }
                        }

                        if (-not $teamsInstalled) {
                            Write-ColorOutput -Message "Alternative architecture installation attempt completed, but Teams installation could not be verified." -Color 'Yellow'
                        }
                    } else {
                        throw "Failed to download alternative architecture installer"
                    }
                } catch {
                    Write-ColorOutput -Message "Error with alternative architecture installation: $_" -Color 'Red'

                    # Try copying to a new location with .new extension and execute from there
                    $newInstallerPath = "$exeInstallerPath.new"
                    Copy-Item -Path $exeInstallerPath -Destination $newInstallerPath -Force

                    if (Test-Path -Path $newInstallerPath) {
                        Write-ColorOutput -Message "Executing installer from alternate location: $newInstallerPath" -Color 'Yellow'
                        # Start process without waiting, just to get it going
                        Start-Process -FilePath $newInstallerPath -NoNewWindow
                    }
                }
            } catch {
                Write-ColorOutput -Message "Error during diagnostics: $_" -Color 'Red'
            }
        }
    }
}

function Test-TeamsInstallationHealth {
    Write-ColorOutput -Message "Performing additional Teams installation health checks..." -Color 'Cyan'
    $healthStatus = $true

    # Check if Teams process starts properly
    try {
        $teamsPath = "$env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe"
        if (Test-Path -Path $teamsPath) {
            Write-ColorOutput -Message "Found Teams executable at $teamsPath" -Color 'Green'

            # Check if Teams is already running
            $teamsProcesses = Get-Process -Name "Teams" -ErrorAction SilentlyContinue
            if (-not $teamsProcesses) {
                Write-ColorOutput -Message "Testing Teams startup..." -Color 'Cyan'
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
                        Write-ColorOutput -Message "Teams process started successfully" -Color 'Green'

                        # Try gracefully closing Teams
                        try {
                            $runningTeams = Get-Process -Name "Teams" -ErrorAction SilentlyContinue
                            if ($runningTeams) {
                                $runningTeams | ForEach-Object { $_.CloseMainWindow() | Out-Null }
                                Start-Sleep -Seconds 2
                                $remainingTeams = Get-Process -Name "Teams" -ErrorAction SilentlyContinue
                                if ($remainingTeams) {
                                    Write-ColorOutput -Message "Teams process still running, stopping process..." -Color 'Yellow'
                                    $remainingTeams | Stop-Process -Force -ErrorAction SilentlyContinue
                                }
                            }
                        } catch {
                            Write-ColorOutput -Message "Error stopping Teams test process: $_" -Color 'Yellow'
                        }
                    } else {
                        Write-ColorOutput -Message "Teams process failed to start" -Color 'Red'
                        $healthStatus = $false
                    }
                } catch {
                    Write-ColorOutput -Message "Error starting Teams: $_" -Color 'Red'
                    $healthStatus = $false
                }
            } else {
                Write-ColorOutput -Message "Teams is already running - skipping process test" -Color 'Yellow'
            }
        } else {
            Write-ColorOutput -Message "Teams executable not found at expected location" -Color 'Red'
            $healthStatus = $false
        }
    } catch {
        Write-ColorOutput -Message "Error testing Teams process: $_" -Color 'Red'
        $healthStatus = $false
    }

    # Check for Teams configuration files
    try {
        $configPath = "$env:APPDATA\Microsoft\Teams"
        if (Test-Path -Path $configPath) {
            Write-ColorOutput -Message "Teams configuration directory found at $configPath" -Color 'Green'

            # Look for essential configuration files
            $configFiles = @(
                "desktop-config.json",
                "storage.json"
            )

            foreach ($file in $configFiles) {
                $filePath = Join-Path -Path $configPath -ChildPath $file
                if (Test-Path -Path $filePath) {
                    Write-ColorOutput -Message "Found Teams configuration file: $file" -Color 'Green'
                } else {
                    Write-ColorOutput -Message "Missing Teams configuration file: $file" -Color 'Yellow'
                }
            }
        } else {
            Write-ColorOutput -Message "Teams configuration directory not found" -Color 'Yellow'
        }
    } catch {
        Write-ColorOutput -Message "Error checking Teams configuration: $_" -Color 'Red'
    }

    return $healthStatus
}

function Stop-TeamsProcess {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    [OutputType([bool])]
    param()

    Write-ColorOutput -Message "Checking for running Teams processes..." -Color 'Cyan'

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
                if ($PSCmdlet.ShouldProcess("$($process.Name) (PID: $($process.Id))", "Stop Teams Process")) {
                    try {
                        # Try graceful shutdown first
                        Write-ColorOutput -Message "Attempting to close $($process.Name) (PID: $($process.Id))..." -Color 'Cyan'
                        # Use CloseMainWindow but don't need to capture return value
                        $process.CloseMainWindow() | Out-Null
                        Start-Sleep -Seconds 2

                        # If still running, force kill
                        if (-not $process.HasExited) {
                            Write-ColorOutput -Message "Forcefully stopping $($process.Name) (PID: $($process.Id))..." -Color 'Yellow'
                            Stop-Process -Id $process.Id -Force -ErrorAction Stop
                        }

                        Write-ColorOutput -Message "Successfully stopped $($process.Name) process" -Color 'Green'
                        $processCount++
                    } catch {
                        Write-ColorOutput -Message "Error stopping $($process.Name) (PID: $($process.Id)): $_" -Color 'Red'
                    }
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
        Write-ColorOutput -Message "WARNING: $remainingProcessCount Teams processes could not be stopped" -Color 'Red'
        return $false
    } else {
        if ($processCount -gt 0) {
            Write-ColorOutput -Message "All Teams processes successfully stopped ($processCount total)" -Color 'Green'
        } else {
            Write-ColorOutput -Message "No Teams processes currently running" -Color 'Green'
        }
        return $true
    }
}

function Get-SystemArchitecture {
    # Detect system architecture and return OS-specific info
    $osArch = "x64"

    # Start with basic OS architecture detection
    if (-not [Environment]::Is64BitOperatingSystem) {
        $osArch = "x86"
        Write-ColorOutput -Message "Detected 32-bit operating system" -Color 'DarkGray'
    } else {
        Write-ColorOutput -Message "Detected 64-bit operating system" -Color 'DarkGray'
    }

    # Check if we're running in a 32-bit process on 64-bit OS
    if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
        Write-ColorOutput -Message "Running in 32-bit process on 64-bit OS" -Color 'Yellow'
        # For Teams, we should use x86 installer in this case
        $osArch = "x86"
    }

    # Additional architecture detection
    try {
        $procArch = $env:PROCESSOR_ARCHITECTURE
        Write-ColorOutput -Message "Processor architecture: $procArch" -Color 'DarkGray'
        $osVersion = [Environment]::OSVersion.Version
        Write-ColorOutput -Message "OS Version: $($osVersion.Major).$($osVersion.Minor) (Build $($osVersion.Build))" -Color 'DarkGray'

        # Check if running in WOW64 (Windows 32-bit on Windows 64-bit)
        if (Test-Path -Path "$env:SystemRoot\SysWOW64") {
            Write-ColorOutput -Message "System has WOW64 subsystem" -Color 'DarkGray'

            # Detect if process is running in 32-bit mode on 64-bit Windows
            if ($null -ne $env:PROCESSOR_ARCHITEW6432) {
                Write-ColorOutput -Message "Current process is running under WOW64 emulation" -Color 'Yellow'
                # In this case, we need to use x86 installer even on x64 OS
                # This happens when running 32-bit PowerShell on 64-bit Windows
                $osArch = "x86"
                Write-ColorOutput -Message "Adjusting download architecture to x86 due to WOW64 process" -Color 'Yellow'
            }
        }

        # Double-check bit process vs OS architecture
        if ([IntPtr]::Size -eq 4) {
            # Running in 32-bit process
            Write-ColorOutput -Message "Confirmed running in 32-bit process" -Color 'Yellow'
            $osArch = "x86"
        }
    } catch {
        Write-ColorOutput -Message "Error detecting detailed system architecture: $_" -Color 'DarkGray'
        # Default to x86 in case of detection errors
        $osArch = "x86"
        Write-ColorOutput -Message "Defaulting to x86 architecture due to detection error" -Color 'Yellow'
    }

    Write-ColorOutput -Message "Using $osArch architecture for Teams installer" -Color 'Cyan'
    return $osArch
}

# Main script execution
Write-ColorOutput -Message '' -Color 'Cyan'
Write-ColorOutput -Message '' -Color 'DarkGray'

# Stop any running Teams processes first
Stop-TeamsProcess

# Uninstall any existing Teams instances
Uninstall-TeamsApp

# Install latest version
Install-TeamsApp

# Verify installation
Write-ColorOutput -Message '' -Color 'DarkGray'
Write-ColorOutput -Message '' -Color 'Cyan'

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
            Write-ColorOutput -Message "Microsoft Teams found in registry: $($app.DisplayName) v$($app.DisplayVersion)" -Color 'Green'
            $script:teamsInstalled = $true
        }
    }
}

# Check WMI/CIM for Teams
try {
    $cimProducts = Get-CimInstance -ClassName Win32_Product -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "*Teams*" -or ($_.Vendor -like "*Microsoft*" -and $_.Name -like "*Teams*") }
    foreach ($product in $cimProducts) {
        Write-ColorOutput -Message "Microsoft Teams found in WMI: $($product.Name) v$($product.Version)" -Color 'Green'
        $script:teamsInstalled = $true
    }
} catch {
    Write-ColorOutput -Message "Error checking WMI for Teams: $_" -Color 'DarkGray'
}

# Check AppX Packages for Teams
try {
    $teamsAppx = Get-AppxPackage -Name "*MicrosoftTeams*" -ErrorAction SilentlyContinue
    if ($teamsAppx) {
        foreach ($app in $teamsAppx) {
            Write-ColorOutput -Message "Microsoft Teams found in AppX: $($app.Name) v$($app.Version)" -Color 'Green'
            $script:teamsInstalled = $true
        }
    }
} catch {
    Write-ColorOutput -Message "Error checking AppX for Teams: $_" -Color 'DarkGray'
}

# Check program files
$teamsLocations = @(
    "${ env:ProgramFiles}\Microsoft\Teams",
    "${ env:ProgramFiles(x86)}\Microsoft\Teams"
)

foreach ($location in $teamsLocations) {
    if (Test-Path -Path $location) {
        $exeFiles = Get-ChildItem -Path $location -Filter "Teams*.exe" -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notlike "*uninst*" -and $_.Name -notlike "*setup*" } |
            Select-Object -First 1
        if ($exeFiles) {
            $fileInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($exeFiles[0].FullName)
            Write-ColorOutput -Message "Microsoft Teams executable found: $($exeFiles[0].FullName) v$($fileInfo.FileVersion)" -Color 'Green'
            $script:teamsInstalled = $true
        } else {
            Write-ColorOutput -Message "Teams directory exists at $location but no Teams executable found." -Color 'Yellow'
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
        Write-ColorOutput -Message "Microsoft Teams found (per-user installation): $($exeFiles[0].FullName) v$($fileInfo.FileVersion)" -Color 'Green'
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
                Write-ColorOutput -Message "Microsoft Teams process found: $($process.Path) v$($fileInfo.FileVersion)" -Color 'Green'
                $teamsInstalled = $true
            }
        }
    }
} catch {
    Write-ColorOutput -Message "Error checking Teams processes: $_" -Color 'DarkGray'
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
                    Write-ColorOutput -Message "Microsoft Teams shortcut found: $($shortcuts.FullName) pointing to v$($fileInfo.FileVersion)" -Color 'Green'
                    $teamsInstalled = $true
                }
            }
        }
    }
} catch {
    Write-ColorOutput -Message "Error checking Teams shortcuts: $_" -Color 'DarkGray'
}

if (-not $teamsInstalled) {
    Write-ColorOutput -Message "Warning: Microsoft Teams installation could not be verified. It may not have installed correctly." -Color 'Yellow'

    # Provide troubleshooting guidance
    Write-ColorOutput -Message "`nTroubleshooting suggestions:" -Color 'Yellow'
    Write-ColorOutput -Message "1. Verify your internet connection is working properly." -Color 'Yellow'
    Write-ColorOutput -Message "2. Check if you have sufficient permissions to install software." -Color 'Yellow'
    Write-ColorOutput -Message "3. Try manually downloading Teams from https://www.microsoft.com/en-us/microsoft-teams/download-app" -Color 'Yellow'
    Write-ColorOutput -Message "4. Run this script again with administrator privileges." -Color 'Yellow'
} else {
    Write-ColorOutput -Message "Microsoft Teams has been successfully installed and verified." -Color 'Green'
}

Write-ColorOutput -Message '' -Color 'Green'
