# =============================================================================
# Script: Reinstall-ForticlientVPN.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2.1
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
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

    By default, runs in silent mode with no UI. Use parameters to show installation window or run interactively.
.PARAMETER Verbose
    When specified, provides detailed information about script execution.
.PARAMETER ShowWindow
    Shows the installation window instead of running hidden.
.PARAMETER Interactive
    Runs the installer in interactive mode instead of silent mode.
.EXAMPLE
    .\Reinstall-ForticlientVPN.ps1
    Runs completely silently with no UI
.EXAMPLE
    .\Reinstall-ForticlientVPN.ps1 -ShowWindow
    Shows the installation window but still runs silently
.EXAMPLE
    .\Reinstall-ForticlientVPN.ps1 -Interactive
    Runs the installer in interactive mode with UI
#>

[CmdletBinding()]
param(
    [switch]$ShowWindow,
    [switch]$Interactive
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


# Set verbose preference and transcript
$VerbosePreference = 'Continue'
$logFile = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Definition) "FortiClientVPN_Install.log"
Start-Transcript -Path $logFile -Append

# Ensure running as administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Verbose "Restarting script with elevated privileges..."
    Start-Process PowerShell -ArgumentList "-NoProfile -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

function Stop-ForticlientService {
    [CmdletBinding(SupportsShouldProcess)]
    param()

    Write-Verbose "Searching for Forticlient processes..."
    Get-Process | Where-Object { $_.Name -like "*Forti*" } | ForEach-Object {
        Write-Verbose "Attempting to stop process: $($_.Name)"
        if ($PSCmdlet.ShouldProcess("Process $($_.Name)", "Stop Process")) {
            $_ | Stop-Process -Force -ErrorAction SilentlyContinue
        }
    }

    Write-Verbose "Searching for Forticlient services..."
    $services = Get-Service -Name "Forticlient*" -ErrorAction SilentlyContinue
    foreach ($service in $services) {
        Write-Verbose "Attempting to stop service: $($service.Name)"
        if ($PSCmdlet.ShouldProcess("Service $($service.Name)", "Stop Service")) {
            Stop-Service -Name $service.Name -Force -ErrorAction SilentlyContinue
            Write-ColorOutput -Message "Stopped service: $($service.Name)" -Color "Green"
        }
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
                    $productCode = $uninstallCmd -replace ".*({ .*})", '$1'
                    Write-Verbose "Found product code: $productCode"
                    Write-ColorOutput -Message "Uninstalling: $($app.DisplayName)" -Color "White"
                    Write-Verbose "Executing: msiexec.exe /x $productCode /qn /norestart"
                    Start-Process "msiexec.exe" -ArgumentList "/x $productCode /qn /norestart" -Wait
                }
            }
        }
    }
}

function Test-Installation {
    Write-Verbose "Verifying FortiClient installation..."
    $maxAttempts = 3
    # seconds
    $retryDelay = 10

    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        Write-Verbose "Verification attempt $attempt of $maxAttempts"

        # Check registry
        $installed = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |
            Where-Object { $_.DisplayName -like "*FortiClient*" }

        if (-not $installed) {
            $installed = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |
                Where-Object { $_.DisplayName -like "*FortiClient*" }
        }

        # Check file system
        $programFiles = @(
            "${ env:ProgramFiles}\Fortinet\FortiClient",
            "${ env:ProgramFiles(x86)}\Fortinet\FortiClient"
        )
        $filesExist = $programFiles | Where-Object { Test-Path $_ } | Select-Object -First 1

        # Check services
        $serviceExists = Get-Service -Name "FortiClient*" -ErrorAction SilentlyContinue

        if ($installed -and $filesExist -and $serviceExists) {
            Write-ColorOutput -Message "FortiClient installation verified successfully" -Color "White"
            return $true
        }

        if ($attempt -lt $maxAttempts) {
            Write-Verbose "Waiting $retryDelay seconds before next verification attempt..."
            Start-Sleep -Seconds $retryDelay
        }
    }

    Write-Warning "FortiClient installation verification failed after $maxAttempts attempts"
    Write-Verbose "Registry check: $($null -ne $installed)"
    Write-Verbose "Files check: $($null -ne $filesExist)"
    Write-Verbose "Service check: $($null -ne $serviceExists)"
    return $false
}

function Install-ForticlientVPN {
    param(
        [switch]$ShowWindow,
        [switch]$Interactive
    )

    $tempPath = Join-Path $env:LOCALAPPDATA "Temp\ForticlientVPN"
    Write-Verbose "Creating temporary directory: $tempPath"
    New-Item -ItemType Directory -Force -Path $tempPath | Out-Null
    $exePath = Join-Path $tempPath "FortiClientVPN.exe"

    try {
        Write-ColorOutput -Message "Downloading Forticlient VPN installer..." -Color "White"
        $downloadUrl = "https://links.fortinet.com/forticlient/win/vpnagent"

        Write-Verbose "Downloading from: $downloadUrl"
        Write-Verbose "Saving to: $exePath"

        # Download with retry logic
        $downloadAttempts = 3
        $success = $false

        for ($i = 1; $i -le $downloadAttempts; $i++) {
            try {
                Invoke-WebRequest -Uri $downloadUrl -OutFile $exePath -TimeoutSec 60
                $success = $true
                break
            } catch {
                Write-Warning "Download attempt $i failed: $_"
                if ($i -lt $downloadAttempts) {
                    Start-Sleep -Seconds 10
                }
            }
        }

        if (-not $success) {
            throw "Failed to download installer after $downloadAttempts attempts"
        }

        Write-ColorOutput -Message "Installing Forticlient VPN..." -Color "White"

        # Install directly using EXE with proper switches
        $installArgs = if ($Interactive) {
            "/passive"
        } else {
            "/quiet /norestart ALLUSERS = 1"
        }

        if ($ShowWindow) {
            $windowStyle = "Normal"
        } else {
            $windowStyle = "Hidden"
        }

        Write-Verbose "Install arguments: $installArgs"
        Write-Verbose "Window style: $windowStyle"

        $installProcess = Start-Process -FilePath $exePath -ArgumentList $installArgs -PassThru -WindowStyle $windowStyle -Wait

        if ($installProcess.ExitCode -ne 0) {
            $errorMessage = Get-InstallerError $installProcess.ExitCode
            throw "Installation failed with exit code: $($installProcess.ExitCode) - $errorMessage"
        }

        # Allow time for installation to complete
        Write-Verbose "Waiting for installation to settle..."
        Start-Sleep -Seconds 30

        # Verify installation
        $verified = $false
        $verificationAttempts = 6
        $verificationDelay = 30

        for ($i = 1; $i -le $verificationAttempts; $i++) {
            Write-Verbose "Verification attempt $i of $verificationAttempts"
            if (Test-Installation) {
                $verified = $true
                Write-ColorOutput -Message "FortiClient VPN installation verified successfully" -Color "White"
                break
            }

            if ($i -lt $verificationAttempts) {
                Write-Verbose "Waiting $verificationDelay seconds before next verification..."
                Start-Sleep -Seconds $verificationDelay
            }
        }

        if (-not $verified) {
            throw "Installation verification failed after $verificationAttempts attempts"
        }
    } catch {
        Write-LogEntry "Error during FortiClient VPN installation: $_" -Type "Error"
        Write-Error $_
        throw
    } finally {
        Write-Verbose "Cleaning up temporary files..."
        if (Test-Path $tempPath) {
            Remove-Item -Path $tempPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Write-LogEntry {
    param(
        [string]$Message,
        [string]$Type = "Info"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Type] $Message"

    # Write to both verbose stream and log file
    Write-Verbose $logMessage

    # Ensure log directory exists
    $scriptDir = Get-ScriptDirectory
    $logPath = Join-Path $scriptDir "FortiClientVPN_Install.log"

    # Write to log file with retry logic and file locking prevention
    $maxAttempts = 3
    $retryDelay = 2

    for ($i = 1; $i -le $maxAttempts; $i++) {
        try {
            $fs = [System.IO.FileStream]::new($logPath, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
            $sw = [System.IO.StreamWriter]::new($fs)
            $sw.WriteLine($logMessage)
            $sw.Close()
            $fs.Close()
            break
        } catch {
            if ($i -eq $maxAttempts) {
                Write-Warning "Failed to write to log file after $maxAttempts attempts: $_"
            } else {
                Start-Sleep -Seconds $retryDelay
            }
        } finally {
            if ($sw) { $sw.Dispose() }
            if ($fs) { $fs.Dispose() }
        }
    }
}

function Get-InstallerError {
    param(
        [int]$ExitCode
    )

    $errorCodes = @{
        1602 = "User cancel installation"
        1603 = "Fatal error during installation"
        1618 = "Another installation is already in progress"
        1619 = "Installation package could not be opened"
        1620 = "Installation package invalid"
        1622 = "Error opening installation log file"
        1623 = "Language not supported"
        1625 = "This installation is forbidden by system policy"
    }

    if ($errorCodes.ContainsKey($ExitCode)) {
        return $errorCodes[$ExitCode]
    }
    return "Unknown error code: $ExitCode"
}

function Get-ScriptDirectory {
    $scriptPath = $PSScriptRoot
    if ($null -eq $scriptPath) {
        $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
    }
    return $scriptPath
}

# Main execution
try {
    Write-LogEntry "=== Starting Forticlient VPN reinstallation process ===" -Type "Info"
    Write-ColorOutput -Message "Starting Forticlient VPN reinstallation..." -Color "White"
    Write-LogEntry "Stopping Forticlient services..." -Type "Info"
    Stop-ForticlientService
    Write-LogEntry "Uninstalling existing Forticlient..." -Type "Info"
    Uninstall-ExistingForticlient
    Write-LogEntry "Installing new Forticlient VPN..." -Type "Info"
    Install-ForticlientVPN -ShowWindow:$ShowWindow -Interactive:$Interactive
    Write-LogEntry "=== Forticlient VPN reinstallation process completed ===" -Type "Info"
    if (Test-Installation) {
        Write-LogEntry "FortiClient VPN reinstallation completed and verified" -Type "Info"
        Write-ColorOutput -Message "FortiClient VPN reinstallation completed and verified" -Color "White"
    } else {
        Write-LogEntry "FortiClient VPN reinstallation could not be verified" -Type "Error"
        Write-Error "FortiClient VPN reinstallation could not be verified"
    }
} finally {
    Stop-Transcript
}
