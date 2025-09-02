# =============================================================================
# Script: Reinstall-MicrosoftC++.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 2.6.1
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
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

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$NoCleanup,
    [switch]$Restart
)

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


# Check for administrative privileges
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-ColorOutput -Message "This script requires administrative privileges. Please run as Administrator." -Color 'Red'
    exit
}

# Get the directory where the script is located
$scriptDirectory = $PSScriptRoot
# Define the path where the redistributable installers will be saved
$downloadPath = "$scriptDirectory\Redistributables"
# Get computer name for log files
$computerName = $env:COMPUTERNAME
# Add a script variable to track if reboot is recommended
$script:rebootRecommended = $false

# Create a single log file name with timestamp and computer name
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$script:logFile = "$scriptDirectory\reinstall_vcredist_${ computerName}_${ timestamp}.log"
$script:logFileCreated = $false

# Function to create a log file
function Write-ScriptLog {
    param (
        [string]$Message,
        [switch]$NoConsole
    )

    # Create log file if it doesn't exist
    if (-not $script:logFileCreated) {
        "$([DateTime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ss UTC')) - Starting Microsoft Visual C++ Redistributable Reinstallation" |
            Out-File -FilePath $script:logFile -Force
        $script:logFileCreated = $true
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] $Message"

    Add-Content -Path $script:logFile -Value $logMessage

    if (-not $NoConsole) {
        Write-ColorOutput -Message $Message -Color "White"
    }
}

# Create the download directory if it doesn't exist
if (!(Test-Path -Path $downloadPath)) {
    if ($PSCmdlet.ShouldProcess("Directory $downloadPath", "Create")) {
        New-Item -ItemType Directory -Path $downloadPath -Force | Out-Null
        Write-ColorOutput -Message "Created temporary directory: $downloadPath" -Color 'Cyan'
    }
}

# Define the URLs and filenames for all Visual C++ Redistributables and Runtimes
$redistributables = @(
    # 2008 SP1
    @{
        Name = "Microsoft Visual C++ 2008 SP1 Redistributable (x86)";
        URL = "https://download.microsoft.com/download/5/D/8/5D8C65CB-C849-4025-8E95-C3966CAFD8AE/vcredist_x86.exe";
        Filename = "vcredist_2008_x86.exe";
        ProductCode = " { FF66E9F6-83E7-3A3E-AF14-8DE9A809A6A4}";
        # Add specific arguments for 2008 SP1 x86
        Args = "/q";
    },
    @{
        Name = "Microsoft Visual C++ 2008 SP1 Redistributable (x64)";
        URL = "https://download.microsoft.com/download/5/D/8/5D8C65CB-C849-4025-8E95-C3966CAFD8AE/vcredist_x64.exe";
        Filename = "vcredist_2008_x64.exe";
        ProductCode = " { 350AA351-21FA-3270-8B7A-835434E766AD}";
        # Add specific arguments for 2008 SP1 x64
        Args = "/q";
    },
    # 2010 SP1
    @{
        Name = "Microsoft Visual C++ 2010 SP1 Redistributable (x86)";
        URL = "https://download.microsoft.com/download/C/6/D/C6D0FD4E-9E53-4897-9B91-836EBA2AACD3/vcredist_x86.exe";
        Filename = "vcredist_2010_x86.exe";
        ProductCode = " { F0C3E5D1-1ADE-321E-8167-68EF0DE699A5}";
    },
    @{
        Name = "Microsoft Visual C++ 2010 SP1 Redistributable (x64)";
        URL = "https://download.microsoft.com/download/A/8/0/A80747C3-41BD-45DF-B505-E9710D2744E0/vcredist_x64.exe";
        Filename = "vcredist_2010_x64.exe";
        ProductCode = " { 1D8E6291-B0D5-35EC-8441-6616F567A0F7}";
    },
    # 2012 Update 4
    @{
        Name = "Microsoft Visual C++ 2012 Redistributable (x86)";
        URL = "https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x86.exe";
        Filename = "vcredist_2012_x86.exe";
        ProductCode = " { 33D1FD90-4274-48A1-9BC1-97E33D9C2D6F}";
    },
    @{
        Name = "Microsoft Visual C++ 2012 Redistributable (x64)";
        URL = "https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x64.exe";
        Filename = "vcredist_2012_x64.exe";
        ProductCode = " { CA67548A-5EBE-413A-B50C-4B9CEB6D66C6}";
    },
    # 2013
    @{
        Name = "Microsoft Visual C++ 2013 Redistributable (x86)";
        URL = "https://download.microsoft.com/download/2/E/6/2E61CFA4-993B-4DD4-91DA-3737CD5CD6E3/vcredist_x86.exe";
        Filename = "vcredist_2013_x86.exe";
        ProductCode = " { E59FD5FB-5A54-3B5C-B04E-7D638C0CFD35}";
    },
    @{
        Name = "Microsoft Visual C++ 2013 Redistributable (x64)";
        URL = "https://download.microsoft.com/download/2/E/6/2E61CFA4-993B-4DD4-91DA-3737CD5CD6E3/vcredist_x64.exe";
        Filename = "vcredist_2013_x64.exe";
        ProductCode = " { 050D4FC8-5D48-4B8F-8972-47C82C46020F}";
    },
    # 2015-2022 (latest versions - same installers used for 2015, 2017, 2019, 2022)
    @{
        Name = "Microsoft Visual C++ 2015-2022 Redistributable (x86)";
        URL = "https://aka.ms/vs/17/release/vc_redist.x86.exe";
        Filename = "vc_redist_2015_2022_x86.exe";
        ProductCode = " { d1a19398-f088-40b5-a0b9-0bdb31d480b7}";
    },
    @{
        Name = "Microsoft Visual C++ 2015-2022 Redistributable (x64)";
        URL = "https://aka.ms/vs/17/release/vc_redist.x64.exe";
        Filename = "vc_redist_2015_2022_x64.exe";
        ProductCode = " { 57a73df6-4ba9-4c45-947a-f635fddeb65c}";
    }
)

# Function to get installed programs
function Get-InstalledProgram {
    param()

    $uninstallKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    $installedPrograms = @()

    foreach ($key in $uninstallKeys) {
        $installedPrograms += Get-ItemProperty -Path $key -ErrorAction SilentlyContinue |
            here-Object { ($_.DisplayName -like "*Microsoft Visual C++*" -or $_.DisplayName -like "*C++ Runtime*") -and $null -eq $_.ParentDisplayName }
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
        Write-ColorOutput -Message "`n$Title ($($Programs.Count)):" -Color $TitleColor
        Write-ScriptLog "$Title ($($Programs.Count))" -NoConsole

        $formattedList = @()

        foreach ($program in $Programs) {
            Write-ColorOutput -Message "  - $($program.DisplayName)" -Color $ItemColor
            $formattedList += "  - $($program.DisplayName)"
        }

        # Log the full list to the log file
        foreach ($item in $formattedList) {
            Write-ScriptLog $item -NoConsole
        }
    } else {
        Write-ColorOutput -Message "`n${ Title}: None found" -Color $TitleColor
        Write-ScriptLog "${ Title}: None found" -NoConsole
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
    $action = if ($Uninstall) { 'uninstall' }else { 'install' }

    # Determine arguments based on file type and action
    switch ($extension) {
        ".msi" {
            # For MSI files use msiexec with appropriate switches
            $action = if ($Uninstall) { "/x" } else { "/i" }
            $arguments = "$action `"$FilePath`" /quiet /norestart ALLUSERS = 1"
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
            Write-ColorOutput -Message "    Unsupported file type: $extension for $DisplayName" -Color 'Red'
            Write-ScriptLog "Unsupported file type: $extension for $DisplayName"
            return $false
        }
    }

    try {
        # Log the command being executed
        Write-ScriptLog "Executing: $(if($extension -eq '.msi') { 'msiexec.exe'}else { $FilePath}) with args: $arguments" -NoConsole

        if ($extension -eq ".msi") {
            $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $arguments -Wait -PassThru -ErrorAction Stop
        } else {
            $process = Start-Process -FilePath $FilePath -ArgumentList $arguments -Wait -PassThru -ErrorAction Stop
        }

        # Check if reboot is recommended
        if ($process.ExitCode -eq 3010) {
            $script:rebootRecommended = $true
        }

        return $process
    } catch {
        Write-ColorOutput -Message "    Error executing ${ DisplayName}: $_" -Color 'Red'
        Write-ScriptLog "Error executing ${ DisplayName}: $_"
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
                Write-ColorOutput -Message "    Retry attempt $retryCount for $DisplayName..." -Color 'Yellow'
                Write-ScriptLog "Retry attempt $retryCount for download: $DisplayName"
                Start-Sleep -Seconds $RetryDelaySeconds
            }

            Invoke-WebRequest -Uri $Url -OutFile $OutFile -ErrorAction Stop -UseBasicParsing
            $success = $true
        } catch {
            $retryCount++
            if ($retryCount -ge $MaxRetries) {
                Write-ColorOutput -Message "    Failed to download $DisplayName after $MaxRetries attempts: $_" -Color 'Red'
                Write-ScriptLog "Download failed after $MaxRetries attempts: $DisplayName - $_"
                return $false
            }
        }
    }

    return $true
}

# Start by creating the log file
if ($PSCmdlet.ShouldProcess("Log file", "Create")) {
    Write-ColorOutput -Message "Log file: $script:logFile" -Color 'Cyan'
}

# Get currently installed Visual C++ Redistributables
Write-ColorOutput -Message "Gathering information about installed Microsoft Visual C++ Redistributables..." -Color 'Cyan'
$installedVCRedists = Get-InstalledProgram

# Use Format-ProgramList to display found redistributables
Format-ProgramList -Programs $installedVCRedists -Title "Found Microsoft Visual C++ Redistributable(s)" -TitleColor Yellow

if ($installedVCRedists.Count -gt 0) {
    # Uninstall existing Visual C++ Redistributables
    Write-ColorOutput -Message "`nRemoving existing Microsoft Visual C++ Redistributables..." -Color 'Cyan'

    foreach ($program in $installedVCRedists) {
        if ($program.UninstallString) {
            $uninstallString = $program.UninstallString

            # Extract the executable path and any existing arguments
            if ($uninstallString -match '"([^"]+)"(.*)') {
                $executable = $matches[1]
                $existingArgs = $matches[2]
            } elseif ($uninstallString -match '([^\s]+)(.*)') {
                $executable = $matches[1]
                $existingArgs = $matches[2]
            }

            if ($PSCmdlet.ShouldProcess("$($program.DisplayName)", "Uninstall")) {
                Write-ColorOutput -Message "  Uninstalling: $($program.DisplayName)" -Color 'Yellow'
                Write-ScriptLog "Removing: $($program.DisplayName)"

                try {
                    # Add a short delay between uninstallations to prevent conflicts
                    Start-Sleep -Seconds 2

                    # Handle MSI uninstallations differently
                    if ($executable -like "*msiexec*" -or $program.UninstallString -like "*msiexec*") {
                        # Extract ProductCode if it's an MSI uninstallation
                        $productCode = ""
                        if ($uninstallString -match " { [0-9A-F]{ 8}-([0-9A-F]{ 4}-) { 3}[0-9A-F]{ 12}}") {
                            $productCode = $matches[0]
                            $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $productCode /quiet /norestart" -Wait -PassThru -ErrorAction Stop
                        } else {
                            # If we can't extract the product code, use the original uninstall string with quiet/passive parameters
                            $process = Start-Process -FilePath $executable -ArgumentList "$existingArgs /quiet /norestart" -Wait -PassThru -ErrorAction Stop
                        }
                    } else {
                        # For EXE uninstallers
                        if ($existingArgs -notlike "*/quiet*" -and $existingArgs -notlike "*/passive*") {
                            $existingArgs += " /quiet /norestart"
                        }
                        $process = Start-Process -FilePath $executable -ArgumentList $existingArgs -Wait -PassThru -ErrorAction Stop
                    }

                    if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
                        Write-ColorOutput -Message "    Successfully uninstalled $($program.DisplayName)" -Color 'Green'
                        Write-ScriptLog "Successfully uninstalled: $($program.DisplayName)"

                        # Check if reboot required
                        if ($process.ExitCode -eq 3010) {
                            $script:rebootRecommended = $true
                        }
                    }
                    # Specific error handling for common error codes
                    elseif ($process.ExitCode -eq 1605) {
                        # Error 1605: This action is only valid for products that are currently installed
                        Write-ColorOutput -Message "    Product $($program.DisplayName) was already uninstalled or not properly installed" -Color 'Yellow'
                        Write-ScriptLog "Product already uninstalled or not properly installed: $($program.DisplayName)"
                    } elseif ($process.ExitCode -eq 1618) {
                        # Error 1618: Another installation is in progress
                        Write-ColorOutput -Message "    Another installation is in progress. Waiting 10 seconds before retry..." -Color 'Yellow'
                        Write-ScriptLog "Another installation in progress for: $($program.DisplayName). Waiting before retry."
                        Start-Sleep -Seconds 10

                        # Try again after waiting
                        $retryProcess = Start-Process -FilePath $executable -ArgumentList $existingArgs -Wait -PassThru -ErrorAction Stop
                        if ($retryProcess.ExitCode -eq 0 -or $retryProcess.ExitCode -eq 3010) {
                            Write-ColorOutput -Message "    Successfully uninstalled $($program.DisplayName) on retry" -Color 'Green'
                            Write-ScriptLog "Successfully uninstalled on retry: $($program.DisplayName)"
                        } else {
                            Write-ColorOutput -Message "    Failed to uninstall $($program.DisplayName) on retry (Exit code: $($retryProcess.ExitCode))" -Color 'Red'
                            Write-ScriptLog "Failed to uninstall on retry: $($program.DisplayName) with exit code: $($retryProcess.ExitCode)"
                        }
                    } else {
                        Write-ColorOutput -Message "    Failed to uninstall $($program.DisplayName) (Exit code: $($process.ExitCode))" -Color 'Red'
                        Write-ScriptLog "Failed to uninstall: $($program.DisplayName) with exit code: $($process.ExitCode)"
                    }
                } catch {
                    Write-ColorOutput -Message "    Error uninstalling $($program.DisplayName): $_" -Color 'Red'
                    Write-ScriptLog "Error uninstalling: $($program.DisplayName) - $_"
                }
            }
        } else {
            Write-ColorOutput -Message "  Unable to uninstall $($program.DisplayName) - No uninstall string found" -Color 'Red'
            Write-ScriptLog "Unable to uninstall: $($program.DisplayName) - No uninstall string found"
        }
    }
} else {
    Write-ColorOutput -Message "No Microsoft Visual C++ Redistributables found on the system." -Color 'Cyan'
    Write-ScriptLog "No Microsoft Visual C++ Redistributables found on the system."
}

# Download all redistributables
Write-ColorOutput -Message "`nDownloading Microsoft Visual C++ Redistributables..." -Color 'Cyan'
Write-ScriptLog "Downloading Microsoft Visual C++ Redistributables..." -NoConsole

foreach ($redist in $redistributables) {
    if ($PSCmdlet.ShouldProcess("$($redist.Name)", "Download")) {
        try {
            Write-ColorOutput -Message "  Downloading $($redist.Name)..." -Color 'Cyan'
            Write-ScriptLog "Downloading: $($redist.Name) from $($redist.URL)"

            $downloadSuccess = Invoke-DownloadWithRetry -Url $redist.URL -OutFile "$downloadPath\$($redist.Filename)" -DisplayName $redist.Name

            if ($downloadSuccess) {
                Write-ColorOutput -Message "    Download complete for $($redist.Name)" -Color 'Green'
                Write-ScriptLog "Download complete: $($redist.Name)"
            }
        } catch {
            Write-ColorOutput -Message "    Failed to download $($redist.Name): $_" -Color 'Red'
            Write-ScriptLog "Download failed: $($redist.Name) - $_"
        }
    }
}

# Install all redistributables
Write-ColorOutput -Message "`nInstalling Microsoft Visual C++ Redistributables..." -Color 'Cyan'
Write-ScriptLog "Installing Microsoft Visual C++ Redistributables..." -NoConsole

foreach ($redist in $redistributables) {
    if ($PSCmdlet.ShouldProcess("$($redist.Name)", "Install")) {
        $filePath = "$downloadPath\$($redist.Filename)"

        if (Test-Path -Path $filePath) {
            try {
                Write-ColorOutput -Message "  Installing $($redist.Name)..." -Color 'Cyan'
                Write-ScriptLog "Installing: $($redist.Name)"

                $customArgs = if ($redist.Args) { $redist.Args } else { "" }
                $process = Invoke-SilentInstallation -FilePath $filePath -DisplayName $redist.Name -CustomArgs $customArgs

                if ($process -and ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010)) {
                    Write-ColorOutput -Message "    Successfully installed $($redist.Name)" -Color 'Green'
                    Write-ScriptLog "Successfully installed: $($redist.Name)"

                    # Check if reboot required
                    if ($process.ExitCode -eq 3010) {
                        Write-ColorOutput -Message "    Note: A system reboot is recommended after installation" -Color 'Yellow'
                        Write-ScriptLog "Reboot recommended after installing: $($redist.Name)"
                        $script:rebootRecommended = $true
                    }
                }
                # Handle specific error codes
                elseif ($process -and $process.ExitCode -eq 5100) {
                    Write-ColorOutput -Message "    Cannot install $($redist.Name) because a newer version is already installed" -Color 'Yellow'
                    Write-ScriptLog "Cannot install: $($redist.Name) - newer version already installed (code 5100)"
                } elseif ($process -and $process.ExitCode -eq 4096) {
                    # For 2008 packages that fail with 4096, try alternate installation parameters
                    Write-ColorOutput -Message "    First attempt failed for $($redist.Name), trying alternate parameters..." -Color 'Yellow'
                    Write-ScriptLog "First attempt failed for $($redist.Name), trying alternate parameters"

                    $process = Invoke-SilentInstallation -FilePath $filePath -DisplayName $redist.Name -CustomArgs "/q /norestart"

                    if ($process -and ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010)) {
                        Write-ColorOutput -Message "    Successfully installed $($redist.Name) with alternate parameters" -Color 'Green'
                        Write-ScriptLog "Successfully installed with alternate parameters: $($redist.Name)"

                        if ($process.ExitCode -eq 3010) {
                            $script:rebootRecommended = $true
                        }
                    } else {
                        $exitCode = if ($process) { $process.ExitCode } else { "Unknown" }
                        Write-ColorOutput -Message "    Failed to install $($redist.Name) with alternate parameters (Exit code: $exitCode)" -Color 'Red'
                        Write-ScriptLog "Failed to install with alternate parameters: $($redist.Name) with exit code: $exitCode"
                    }
                } else {
                    $exitCode = if ($process) { $process.ExitCode } else { "Unknown" }
                    Write-ColorOutput -Message "    Failed to install $($redist.Name) (Exit code: $exitCode)" -Color 'Red'
                    Write-ScriptLog "Failed to install: $($redist.Name) with exit code: $exitCode"
                }
            } catch {
                Write-ColorOutput -Message "    Error installing $($redist.Name): $_" -Color 'Red'
                Write-ScriptLog "Error installing: $($redist.Name) - $_"
            }
        } else {
            Write-ColorOutput -Message "    Installation file for $($redist.Name) not found at $filePath" -Color 'Red'
            Write-ScriptLog "Installation file not found: $($redist.Name) at path $filePath"
        }
    }
}

# Verify installation
Write-ColorOutput -Message "`nVerifying installations..." -Color 'Cyan'
Write-ScriptLog "Verifying installations..." -NoConsole

if ($PSCmdlet.ShouldProcess("Microsoft Visual C++ Redistributables", "Verify installation")) {
    $installedAfter = Get-InstalledProgram

    # Use Format-ProgramList to display installed redistributables
    Format-ProgramList -Programs $installedAfter -Title "Successfully installed Microsoft Visual C++ Redistributable(s)" -TitleColor Green

    # Compare before and after installation
    if ($installedAfter.Count -eq 0) {
        Write-ColorOutput -Message "No Microsoft Visual C++ Redistributables were found after installation. This may indicate an installation problem." -Color 'Red'
        Write-ScriptLog "No Microsoft Visual C++ Redistributables found after installation - possible installation failure."
    } elseif ($installedAfter.Count -lt $redistributables.Count) {
        Write-ColorOutput -Message "Warning: Not all expected redistributables were installed. Expected $($redistributables.Count) but found $($installedAfter.Count)." -Color 'Yellow'
        Write-ScriptLog "Warning: Not all expected redistributables were installed. Expected $($redistributables.Count) but found $($installedAfter.Count)."
    }

    # Check if specific versions are missing
    $installedProducts = $installedAfter | ForEach-Object { $_.DisplayName }
    $missingVersions = @()

    foreach ($redist in $redistributables) {
        $found = $false
        foreach ($installed in $installedProducts) {
            if ($installed -like "*$($redist.Name)*" -or ($redist.Name -match '(\d { 4})' -and $installed -match $matches[1])) {
                $found = $true
                break
            }
        }

        if (-not $found) {
            $missingVersions += $redist.Name
        }
    }

    if ($missingVersions.Count -gt 0) {
        Write-ColorOutput -Message "`nPotentially missing redistributables:" -Color 'Yellow'
        foreach ($missing in $missingVersions) {
            Write-ColorOutput -Message "  - $missing" -Color 'Yellow'
            Write-ScriptLog "Potentially missing: $missing"
        }
    }
}

# Clean up downloaded files if requested
if (-not $NoCleanup) {
    if ($PSCmdlet.ShouldProcess("Downloaded Installation Files", "Clean up")) {
        Write-ColorOutput -Message "`nCleaning up downloaded files..." -Color 'Cyan'
        Write-ScriptLog "Cleaning up downloaded files..." -NoConsole

        try {
            Remove-Item -Path $downloadPath -Recurse -Force -ErrorAction Stop
            Write-ColorOutput -Message "  Successfully removed downloaded files" -Color 'Green'
            Write-ScriptLog "Successfully removed downloaded files" -NoConsole
        } catch {
            Write-ColorOutput -Message "  Failed to remove downloaded files: $_" -Color 'Yellow'
            Write-ScriptLog "Failed to remove downloaded files: $_" -NoConsole
        }
    }
} else {
    Write-ColorOutput -Message "`nDownloaded files remain in: $downloadPath" -Color 'Cyan'
    Write-ScriptLog "Downloaded files remain in: $downloadPath" -NoConsole
}

# Display reboot recommendation at the end if needed
if ($script:rebootRecommended) {
    Write-ColorOutput -Message "`nA system reboot is recommended to complete the installation process." -Color 'Yellow'
    Write-ScriptLog "A system reboot is recommended to complete the installation process." -NoConsole

    if ($Restart) {
        Write-ColorOutput -Message "System will restart in 15 seconds. Press Ctrl+C to cancel." -Color 'Yellow'
        Write-ScriptLog "System will restart in 15 seconds." -NoConsole

        if ($PSCmdlet.ShouldProcess("System", "Restart")) {
            Start-Sleep -Seconds 15
            Restart-Computer -Force
        }
    }
}

Write-ColorOutput -Message "`nProcess complete. Log file saved to: $script:logFile" -Color 'Green'
Write-ScriptLog "Process complete" -NoConsole
