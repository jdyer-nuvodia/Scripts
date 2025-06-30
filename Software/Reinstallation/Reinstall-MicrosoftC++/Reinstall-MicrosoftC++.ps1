# =============================================================================
# Script: Reinstall-MicrosoftC++.ps1
# Created: 2025-02-27 18:51:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-14 17:30:00 UTC
# Updated By: jdyer-nuvodia
# Version: 2.5.0
# Additional Info: Added cleanup functionality, restart parameter, and fixed duplicate log messages
# =============================================================================

<#
.SYNOPSIS
    Removes and reinstalls Microsoft Visual C++ Redistributables and Runtimes (x86 and x64) from 2008 to latest.
.DESCRIPTION
    This script automates the process of removing existing Microsoft Visual C++ Redistributables and installing
    all versions from 2008 to the latest, including both redistributable packages and runtime components.
    Key actions:
     - Removes all existing Visual C++ Redistributables and Runtimes
     - Creates a temporary directory for downloads
     - Downloads all versions (2008, 2010, 2012, 2013, 2015-2022) of Visual C++ Redistributables and Runtimes
     - Installs the components silently
     - No system restart is forced after installation unless specified
     - Includes -WhatIf parameter to preview changes without executing them
     - Cleans up downloaded files after installation unless specified
    Dependencies:
     - Requires internet connection
     - Requires administrative privileges
.PARAMETER WhatIf
    Simulates the removal and installation process without making actual changes.

.PARAMETER NoCleanup
    Skip cleanup of downloaded installation files.

.PARAMETER Restart
    Automatically restart the computer if needed after installation.

.EXAMPLE
    .\Reinstall-MicrosoftC++.ps1
    Removes all existing Visual C++ Redistributables/Runtimes, installs all versions from 2008 to latest, and cleans up downloaded files

.EXAMPLE
    .\Reinstall-MicrosoftC++.ps1 -WhatIf
    Shows what would happen if the script is run without making any changes

.EXAMPLE
    .\Reinstall-MicrosoftC++.ps1 -NoCleanup
    Performs installation but keeps the downloaded installation files

.EXAMPLE
    .\Reinstall-MicrosoftC++.ps1 -Restart
    Performs installation and restarts the computer if needed

.NOTES
    Security Level: Medium
    Required Permissions: Administrative privileges
    Validation Requirements: Verify successful installation in Programs and Features
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [switch]$NoCleanup,
    [switch]$Restart
)

# Check for administrative privileges
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "This script requires administrative privileges. Please run as Administrator." -ForegroundColor Red
    Exit
}

# Get the directory where the script is located
$scriptDirectory = $PSScriptRoot
# Define the path where the redistributable installers will be saved
$downloadPath = "$scriptDirectory\Redistributables"
# Get computer name for log files
$computerName = $env:COMPUTERNAME
# Add a global variable to track if reboot is recommended
$global:rebootRecommended = $false

# Create a single log file name with timestamp and computer name
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$global:logFile = "$scriptDirectory\reinstall_vcredist_${computerName}_${timestamp}.log"
$global:logFileCreated = $false

# Function to create a log file
function Write-Log {
    param (
        [string]$Message,
        [switch]$NoConsole
    )

    # Create log file if it doesn't exist
    if (-not $global:logFileCreated) {
        "$([DateTime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ss UTC')) - Starting Microsoft Visual C++ Redistributable Reinstallation" |
            Out-File -FilePath $global:logFile -Force
        $global:logFileCreated = $true
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] $Message"

    Add-Content -Path $global:logFile -Value $logMessage

    if (-not $NoConsole) {
        Write-Host $Message
    }
}

# Create the download directory if it doesn't exist
if (!(Test-Path -Path $downloadPath)) {
    if ($PSCmdlet.ShouldProcess("Directory $downloadPath", "Create")) {
        New-Item -ItemType Directory -Path $downloadPath -Force | Out-Null
        Write-Host "Created temporary directory: $downloadPath" -ForegroundColor Cyan
    }
}

# Define the URLs and filenames for all Visual C++ Redistributables and Runtimes
$redistributables = @(
    # 2008 SP1
    @{
        Name = "Microsoft Visual C++ 2008 SP1 Redistributable (x86)";
        URL = "https://download.microsoft.com/download/5/D/8/5D8C65CB-C849-4025-8E95-C3966CAFD8AE/vcredist_x86.exe";
        Filename = "vcredist_2008_x86.exe";
        ProductCode = "{FF66E9F6-83E7-3A3E-AF14-8DE9A809A6A4}";
        # Add specific arguments for 2008 SP1 x86
                Args = "/q";
    },
    @{
        Name = "Microsoft Visual C++ 2008 SP1 Redistributable (x64)";
        URL = "https://download.microsoft.com/download/5/D/8/5D8C65CB-C849-4025-8E95-C3966CAFD8AE/vcredist_x64.exe";
        Filename = "vcredist_2008_x64.exe";
        ProductCode = "{350AA351-21FA-3270-8B7A-835434E766AD}";
        # Add specific arguments for 2008 SP1 x64
                Args = "/q";
    },
    # 2010 SP1
    @{
        Name = "Microsoft Visual C++ 2010 SP1 Redistributable (x86)";
        URL = "https://download.microsoft.com/download/C/6/D/C6D0FD4E-9E53-4897-9B91-836EBA2AACD3/vcredist_x86.exe";
        Filename = "vcredist_2010_x86.exe";
        ProductCode = "{F0C3E5D1-1ADE-321E-8167-68EF0DE699A5}";
    },
    @{
        Name = "Microsoft Visual C++ 2010 SP1 Redistributable (x64)";
        URL = "https://download.microsoft.com/download/A/8/0/A80747C3-41BD-45DF-B505-E9710D2744E0/vcredist_x64.exe";
        Filename = "vcredist_2010_x64.exe";
        ProductCode = "{1D8E6291-B0D5-35EC-8441-6616F567A0F7}";
    },
    # 2012 Update 4
    @{
        Name = "Microsoft Visual C++ 2012 Redistributable (x86)";
        URL = "https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x86.exe";
        Filename = "vcredist_2012_x86.exe";
        ProductCode = "{33D1FD90-4274-48A1-9BC1-97E33D9C2D6F}";
    },
    @{
        Name = "Microsoft Visual C++ 2012 Redistributable (x64)";
        URL = "https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x64.exe";
        Filename = "vcredist_2012_x64.exe";
        ProductCode = "{CA67548A-5EBE-413A-B50C-4B9CEB6D66C6}";
    },
    # 2013
    @{
        Name = "Microsoft Visual C++ 2013 Redistributable (x86)";
        URL = "https://download.microsoft.com/download/2/E/6/2E61CFA4-993B-4DD4-91DA-3737CD5CD6E3/vcredist_x86.exe";
        Filename = "vcredist_2013_x86.exe";
        ProductCode = "{E59FD5FB-5A54-3B5C-B04E-7D638C0CFD35}";
    },
    @{
        Name = "Microsoft Visual C++ 2013 Redistributable (x64)";
        URL = "https://download.microsoft.com/download/2/E/6/2E61CFA4-993B-4DD4-91DA-3737CD5CD6E3/vcredist_x64.exe";
        Filename = "vcredist_2013_x64.exe";
        ProductCode = "{050D4FC8-5D48-4B8F-8972-47C82C46020F}";
    },
    # 2015-2022 (latest versions - same installers used for 2015, 2017, 2019, 2022)
    @{
        Name = "Microsoft Visual C++ 2015-2022 Redistributable (x86)";
        URL = "https://aka.ms/vs/17/release/vc_redist.x86.exe";
        Filename = "vc_redist_2015_2022_x86.exe";
        ProductCode = "{d1a19398-f088-40b5-a0b9-0bdb31d480b7}";
    },
    @{
        Name = "Microsoft Visual C++ 2015-2022 Redistributable (x64)";
        URL = "https://aka.ms/vs/17/release/vc_redist.x64.exe";
        Filename = "vc_redist_2015_2022_x64.exe";
        ProductCode = "{57a73df6-4ba9-4c45-947a-f635fddeb65c}";
    }
)

# Function to get installed programs
function Get-InstalledPrograms {
    param()

    $uninstallKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    $installedPrograms = @()

    foreach ($key in $uninstallKeys) {
        $installedPrograms += Get-ItemProperty -Path $key -ErrorAction SilentlyContinue |
            Where-Object { ($_.DisplayName -like "*Microsoft Visual C++*" -or $_.DisplayName -like "*C++ Runtime*") -and $null -eq $_.ParentDisplayName }
    }

    return $installedPrograms | Sort-Object DisplayName
}

# Function to format a list of programs for display and logging
function Format-ProgramList {
    param(
        [array]$Programs,
        [string]$Title,
        [System.ConsoleColor]$TitleColor = "Cyan",
        [System.ConsoleColor]$ItemColor = "DarkGray"
    )

    if ($Programs.Count -gt 0) {
        Write-Host "`n$Title ($($Programs.Count)):" -ForegroundColor $TitleColor
        Write-Log "$Title ($($Programs.Count))" -NoConsole

        $formattedList = @()

        foreach ($program in $Programs) {
            Write-Host "  - $($program.DisplayName)" -ForegroundColor $ItemColor
            $formattedList += "  - $($program.DisplayName)"
        }

        # Log the full list to the log file
        foreach ($item in $formattedList) {
            Write-Log $item -NoConsole
        }
    }
    else {
        Write-Host "`n${Title}: None found" -ForegroundColor $TitleColor
        Write-Log "${Title}: None found" -NoConsole
    }
}

# Function to properly handle different installer types with silent options
function Invoke-SilentInstallation {
    param (
        [string]$FilePath,
        [string]$DisplayName,
        [switch]$Uninstall,
        [string]$CustomArgs = ""
    )

    $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
    $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    $action = if($Uninstall){'uninstall'}else{'install'}

    # Determine arguments based on file type and action
    switch ($extension) {
        ".msi" {
            # For MSI files use msiexec with appropriate switches
            $action = if ($Uninstall) { "/x" } else { "/i" }
            $arguments = "$action `"$FilePath`" /quiet /norestart ALLUSERS=1"
            if ($CustomArgs) {
                $arguments += " $CustomArgs"
            }
        }
        ".exe" {
            # Use custom args if provided, otherwise use default
            if ($CustomArgs) {
                $arguments = $CustomArgs
            } else {
                # Most Visual C++ redistributables use these parameters
                if ($Uninstall) {
                    $arguments = "/uninstall /quiet /norestart"
                } else {
                    $arguments = "/install /quiet /norestart"
                }
            }
        }
        default {
            Write-Host "    Unsupported file type: $extension for $DisplayName" -ForegroundColor Red
            Write-Log "Unsupported file type: $extension for $DisplayName"
            return $false
        }
    }

    try {
        # Log the command being executed
        Write-Log "Executing: $(if($extension -eq '.msi'){'msiexec.exe'}else{$FilePath}) with args: $arguments" -NoConsole

        if ($extension -eq ".msi") {
            $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $arguments -Wait -PassThru -ErrorAction Stop
        } else {
            $process = Start-Process -FilePath $FilePath -ArgumentList $arguments -Wait -PassThru -ErrorAction Stop
        }

        # Check if reboot is recommended
        if ($process.ExitCode -eq 3010) {
            $global:rebootRecommended = $true
        }

        return $process
    }
    catch {
        Write-Host "    Error executing ${DisplayName}: $_" -ForegroundColor Red
        Write-Log "Error executing ${DisplayName}: $_"
        return $false
    }
}

# Function to download file with retry logic
function Invoke-DownloadWithRetry {
    param (
        [string]$Url,
        [string]$OutFile,
        [string]$DisplayName,
        [int]$MaxRetries = 3,
        [int]$RetryDelaySeconds = 5
    )

    $retryCount = 0
    $success = $false

    while (-not $success -and $retryCount -lt $MaxRetries) {
        try {
            if ($retryCount -gt 0) {
                Write-Host "    Retry attempt $retryCount for $DisplayName..." -ForegroundColor Yellow
                Write-Log "Retry attempt $retryCount for download: $DisplayName"
                Start-Sleep -Seconds $RetryDelaySeconds
            }

            Invoke-WebRequest -Uri $Url -OutFile $OutFile -ErrorAction Stop -UseBasicParsing
            $success = $true
        }
        catch {
            $retryCount++
            if ($retryCount -ge $MaxRetries) {
                Write-Host "    Failed to download $DisplayName after $MaxRetries attempts: $_" -ForegroundColor Red
                Write-Log "Download failed after $MaxRetries attempts: $DisplayName - $_"
                return $false
            }
        }
    }

    return $true
}

# Start by creating the log file
if ($PSCmdlet.ShouldProcess("Log file", "Create")) {
    Write-Host "Log file created: $global:logFile" -ForegroundColor Cyan
}

# Get currently installed Visual C++ Redistributables
Write-Host "Gathering information about installed Microsoft Visual C++ Redistributables..." -ForegroundColor Cyan
$installedVCRedists = Get-InstalledPrograms

# Use Format-ProgramList to display found redistributables
Format-ProgramList -Programs $installedVCRedists -Title "Found Microsoft Visual C++ Redistributable(s)" -TitleColor Yellow

if ($installedVCRedists.Count -gt 0) {
    # Uninstall existing Visual C++ Redistributables
    Write-Host "`nRemoving existing Microsoft Visual C++ Redistributables..." -ForegroundColor Cyan

    foreach ($program in $installedVCRedists) {
        if ($program.UninstallString) {
            $uninstallString = $program.UninstallString

            # Extract the executable path and any existing arguments
            if ($uninstallString -match '"([^"]+)"(.*)') {
                $executable = $matches[1]
                $existingArgs = $matches[2]
            }
            elseif ($uninstallString -match '([^\s]+)(.*)') {
                $executable = $matches[1]
                $existingArgs = $matches[2]
            }

            if ($PSCmdlet.ShouldProcess("$($program.DisplayName)", "Uninstall")) {
                Write-Host "  Uninstalling: $($program.DisplayName)" -ForegroundColor Yellow
                Write-Log "Removing: $($program.DisplayName)"

                try {
                    # Add a short delay between uninstallations to prevent conflicts
                    Start-Sleep -Seconds 2

                    # Handle MSI uninstallations differently
                    if ($executable -like "*msiexec*" -or $program.UninstallString -like "*msiexec*") {
                        # Extract ProductCode if it's an MSI uninstallation
                        $productCode = ""
                        if ($uninstallString -match "{[0-9A-F]{8}-([0-9A-F]{4}-){3}[0-9A-F]{12}}") {
                            $productCode = $matches[0]
                            $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $productCode /quiet /norestart" -Wait -PassThru -ErrorAction Stop
                        } else {
                            # If we can't extract the product code, use the original uninstall string with quiet/passive parameters
                            $process = Start-Process -FilePath $executable -ArgumentList "$existingArgs /quiet /norestart" -Wait -PassThru -ErrorAction Stop
                        }
                    }
                    else {
                        # For EXE uninstallers
                        if ($existingArgs -notlike "*/quiet*" -and $existingArgs -notlike "*/passive*") {
                            $existingArgs += " /quiet /norestart"
                        }
                        $process = Start-Process -FilePath $executable -ArgumentList $existingArgs -Wait -PassThru -ErrorAction Stop
                    }

                    if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
                        Write-Host "    Successfully uninstalled $($program.DisplayName)" -ForegroundColor Green
                        Write-Log "Successfully uninstalled: $($program.DisplayName)"

                        # Check if reboot required
                        if ($process.ExitCode -eq 3010) {
                            $global:rebootRecommended = $true
                        }
                    }
                    # Specific error handling for common error codes
                    elseif ($process.ExitCode -eq 1605) {
                        # Error 1605: This action is only valid for products that are currently installed
                        Write-Host "    Product $($program.DisplayName) was already uninstalled or not properly installed" -ForegroundColor Yellow
                        Write-Log "Product already uninstalled or not properly installed: $($program.DisplayName)"
                    }
                    elseif ($process.ExitCode -eq 1618) {
                        # Error 1618: Another installation is in progress
                        Write-Host "    Another installation is in progress. Waiting 10 seconds before retry..." -ForegroundColor Yellow
                        Write-Log "Another installation in progress for: $($program.DisplayName). Waiting before retry."
                        Start-Sleep -Seconds 10

                        # Try again after waiting
                        $retryProcess = Start-Process -FilePath $executable -ArgumentList $existingArgs -Wait -PassThru -ErrorAction Stop
                        if ($retryProcess.ExitCode -eq 0 -or $retryProcess.ExitCode -eq 3010) {
                            Write-Host "    Successfully uninstalled $($program.DisplayName) on retry" -ForegroundColor Green
                            Write-Log "Successfully uninstalled on retry: $($program.DisplayName)"
                        } else {
                            Write-Host "    Failed to uninstall $($program.DisplayName) on retry (Exit code: $($retryProcess.ExitCode))" -ForegroundColor Red
                            Write-Log "Failed to uninstall on retry: $($program.DisplayName) with exit code: $($retryProcess.ExitCode)"
                        }
                    }
                    else {
                        Write-Host "    Failed to uninstall $($program.DisplayName) (Exit code: $($process.ExitCode))" -ForegroundColor Red
                        Write-Log "Failed to uninstall: $($program.DisplayName) with exit code: $($process.ExitCode)"
                    }
                }
                catch {
                    Write-Host "    Error uninstalling $($program.DisplayName): $_" -ForegroundColor Red
                    Write-Log "Error uninstalling: $($program.DisplayName) - $_"
                }
            }
        }
        else {
            Write-Host "  Unable to uninstall $($program.DisplayName) - No uninstall string found" -ForegroundColor Red
            Write-Log "Unable to uninstall: $($program.DisplayName) - No uninstall string found"
        }
    }
}
else {
    Write-Host "No Microsoft Visual C++ Redistributables found on the system." -ForegroundColor Cyan
    Write-Log "No Microsoft Visual C++ Redistributables found on the system."
}

# Download all redistributables
Write-Host "`nDownloading Microsoft Visual C++ Redistributables..." -ForegroundColor Cyan
Write-Log "Downloading Microsoft Visual C++ Redistributables..." -NoConsole

foreach ($redist in $redistributables) {
    if ($PSCmdlet.ShouldProcess("$($redist.Name)", "Download")) {
        try {
            Write-Host "  Downloading $($redist.Name)..." -ForegroundColor Cyan
            Write-Log "Downloading: $($redist.Name) from $($redist.URL)"

            $downloadSuccess = Invoke-DownloadWithRetry -Url $redist.URL -OutFile "$downloadPath\$($redist.Filename)" -DisplayName $redist.Name

            if ($downloadSuccess) {
                Write-Host "    Download complete for $($redist.Name)" -ForegroundColor Green
                Write-Log "Download complete: $($redist.Name)"
            }
        }
        catch {
            Write-Host "    Failed to download $($redist.Name): $_" -ForegroundColor Red
            Write-Log "Download failed: $($redist.Name) - $_"
        }
    }
}

# Install all redistributables
Write-Host "`nInstalling Microsoft Visual C++ Redistributables..." -ForegroundColor Cyan
Write-Log "Installing Microsoft Visual C++ Redistributables..." -NoConsole

foreach ($redist in $redistributables) {
    if ($PSCmdlet.ShouldProcess("$($redist.Name)", "Install")) {
        $filePath = "$downloadPath\$($redist.Filename)"

        if (Test-Path -Path $filePath) {
            try {
                Write-Host "  Installing $($redist.Name)..." -ForegroundColor Cyan
                Write-Log "Installing: $($redist.Name)"

                $customArgs = if ($redist.Args) { $redist.Args } else { "" }
                $process = Invoke-SilentInstallation -FilePath $filePath -DisplayName $redist.Name -CustomArgs $customArgs

                if ($process -and ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010)) {
                    Write-Host "    Successfully installed $($redist.Name)" -ForegroundColor Green
                    Write-Log "Successfully installed: $($redist.Name)"

                    # Check if reboot required
                    if ($process.ExitCode -eq 3010) {
                        Write-Host "    Note: A system reboot is recommended after installation" -ForegroundColor Yellow
                        Write-Log "Reboot recommended after installing: $($redist.Name)"
                        $global:rebootRecommended = $true
                    }
                }
                # Handle specific error codes
                elseif ($process -and $process.ExitCode -eq 5100) {
                    Write-Host "    Cannot install $($redist.Name) because a newer version is already installed" -ForegroundColor Yellow
                    Write-Log "Cannot install: $($redist.Name) - newer version already installed (code 5100)"
                }
                elseif ($process -and $process.ExitCode -eq 4096) {
                    # For 2008 packages that fail with 4096, try alternate installation parameters
                    Write-Host "    First attempt failed for $($redist.Name), trying alternate parameters..." -ForegroundColor Yellow
                    Write-Log "First attempt failed for $($redist.Name), trying alternate parameters"

                    $process = Invoke-SilentInstallation -FilePath $filePath -DisplayName $redist.Name -CustomArgs "/q /norestart"

                    if ($process -and ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010)) {
                        Write-Host "    Successfully installed $($redist.Name) with alternate parameters" -ForegroundColor Green
                        Write-Log "Successfully installed with alternate parameters: $($redist.Name)"

                        if ($process.ExitCode -eq 3010) {
                            $global:rebootRecommended = $true
                        }
                    }
                    else {
                        $exitCode = if ($process) { $process.ExitCode } else { "Unknown" }
                        Write-Host "    Failed to install $($redist.Name) with alternate parameters (Exit code: $exitCode)" -ForegroundColor Red
                        Write-Log "Failed to install with alternate parameters: $($redist.Name) with exit code: $exitCode"
                    }
                }
                else {
                    $exitCode = if ($process) { $process.ExitCode } else { "Unknown" }
                    Write-Host "    Failed to install $($redist.Name) (Exit code: $exitCode)" -ForegroundColor Red
                    Write-Log "Failed to install: $($redist.Name) with exit code: $exitCode"
                }
            }
            catch {
                Write-Host "    Error installing $($redist.Name): $_" -ForegroundColor Red
                Write-Log "Error installing: $($redist.Name) - $_"
            }
        }
        else {
            Write-Host "    Installation file for $($redist.Name) not found at $filePath" -ForegroundColor Red
            Write-Log "Installation file not found: $($redist.Name) at path $filePath"
        }
    }
}

# Verify installation
Write-Host "`nVerifying installations..." -ForegroundColor Cyan
Write-Log "Verifying installations..." -NoConsole

if ($PSCmdlet.ShouldProcess("Microsoft Visual C++ Redistributables", "Verify installation")) {
    $installedAfter = Get-InstalledPrograms

    # Use Format-ProgramList to display installed redistributables
    Format-ProgramList -Programs $installedAfter -Title "Successfully installed Microsoft Visual C++ Redistributable(s)" -TitleColor Green

    # Compare before and after installation
    if ($installedAfter.Count -eq 0) {
        Write-Host "No Microsoft Visual C++ Redistributables were found after installation. This may indicate an installation problem." -ForegroundColor Red
        Write-Log "No Microsoft Visual C++ Redistributables found after installation - possible installation failure."
    }
    elseif ($installedAfter.Count -lt $redistributables.Count) {
        Write-Host "Warning: Not all expected redistributables were installed. Expected $($redistributables.Count) but found $($installedAfter.Count)." -ForegroundColor Yellow
        Write-Log "Warning: Not all expected redistributables were installed. Expected $($redistributables.Count) but found $($installedAfter.Count)."
    }

    # Check if specific versions are missing
    $installedProducts = $installedAfter | ForEach-Object { $_.DisplayName }
    $missingVersions = @()

    foreach ($redist in $redistributables) {
        $found = $false
        foreach ($installed in $installedProducts) {
            if ($installed -like "*$($redist.Name)*" -or ($redist.Name -match '(\d{4})' -and $installed -match $matches[1])) {
                $found = $true
                break
            }
        }

        if (-not $found) {
            $missingVersions += $redist.Name
        }
    }

    if ($missingVersions.Count -gt 0) {
        Write-Host "`nPotentially missing redistributables:" -ForegroundColor Yellow
        foreach ($missing in $missingVersions) {
            Write-Host "  - $missing" -ForegroundColor Yellow
            Write-Log "Potentially missing: $missing"
        }
    }
}

# Clean up downloaded files if requested
if (-not $NoCleanup) {
    if ($PSCmdlet.ShouldProcess("Downloaded Installation Files", "Clean up")) {
        Write-Host "`nCleaning up downloaded files..." -ForegroundColor Cyan
        Write-Log "Cleaning up downloaded files..." -NoConsole

        try {
            Remove-Item -Path $downloadPath -Recurse -Force -ErrorAction Stop
            Write-Host "  Successfully removed downloaded files" -ForegroundColor Green
            Write-Log "Successfully removed downloaded files" -NoConsole
        }
        catch {
            Write-Host "  Failed to remove downloaded files: $_" -ForegroundColor Yellow
            Write-Log "Failed to remove downloaded files: $_" -NoConsole
        }
    }
}
else {
    Write-Host "`nDownloaded files remain in: $downloadPath" -ForegroundColor Cyan
    Write-Log "Downloaded files remain in: $downloadPath" -NoConsole
}

# Display reboot recommendation at the end if needed
if ($global:rebootRecommended) {
    Write-Host "`nA system reboot is recommended to complete the installation process." -ForegroundColor Yellow
    Write-Log "A system reboot is recommended to complete the installation process." -NoConsole

    if ($Restart) {
        Write-Host "System will restart in 15 seconds. Press Ctrl+C to cancel." -ForegroundColor Yellow
        Write-Log "System will restart in 15 seconds." -NoConsole

        if ($PSCmdlet.ShouldProcess("System", "Restart")) {
            Start-Sleep -Seconds 15
            Restart-Computer -Force
        }
    }
}

Write-Host "`nProcess complete. Log file saved to: $global:logFile" -ForegroundColor Green
Write-Log "Process complete" -NoConsole
