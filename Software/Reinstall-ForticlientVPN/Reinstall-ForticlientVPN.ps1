# =============================================================================
# Script: Reinstall-ForticlientVPN.ps1
# Created: 2024-02-13 18:30:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2024-02-13 18:30:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1
# Additional Info: Script to remove and reinstall Forticlient VPN
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

# Set verbose preference and transcript
$VerbosePreference = 'Continue'
$logFile = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Definition) "FortiClientVPN_Install.log"
Start-Transcript -Path $logFile -Append

# Ensure running as administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Verbose "Restarting script with elevated privileges..."
    Start-Process PowerShell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

function Stop-ForticlientServices {
    Write-Verbose "Searching for Forticlient processes..."
    Get-Process | Where-Object { $_.Name -like "*Forti*" } | ForEach-Object {
        Write-Verbose "Attempting to stop process: $($_.Name)"
        $_ | Stop-Process -Force -ErrorAction SilentlyContinue
    }
    
    Write-Verbose "Searching for Forticlient services..."
    $services = Get-Service -Name "Forticlient*" -ErrorAction SilentlyContinue
    foreach ($service in $services) {
        Write-Verbose "Attempting to stop service: $($service.Name)"
        Stop-Service -Name $service.Name -Force -ErrorAction SilentlyContinue
        Write-Host "Stopped service: $($service.Name)"
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
                    $productCode = $uninstallCmd -replace ".*({.*})", '$1'
                    Write-Verbose "Found product code: $productCode"
                    Write-Host "Uninstalling: $($app.DisplayName)"
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
    $retryDelay = 10 # seconds
    
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
            "${env:ProgramFiles}\Fortinet\FortiClient",
            "${env:ProgramFiles(x86)}\Fortinet\FortiClient"
        )
        $filesExist = $programFiles | Where-Object { Test-Path $_ } | Select-Object -First 1
        
        # Check services
        $serviceExists = Get-Service -Name "FortiClient*" -ErrorAction SilentlyContinue

        if ($installed -and $filesExist -and $serviceExists) {
            Write-Host "FortiClient installation verified successfully"
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
    $tempPath = Join-Path $env:TEMP "ForticlientVPN"
    Write-Verbose "Creating temporary directory: $tempPath"
    New-Item -ItemType Directory -Force -Path $tempPath | Out-Null
    $installerPath = Join-Path $tempPath "ForticlientVPN.exe"

    Write-Host "Downloading Forticlient VPN installer..."
    $downloadUrl = "https://links.fortinet.com/forticlient/win/vpnagent"
    
    try {
        Write-Verbose "Downloading from: $downloadUrl"
        Write-Verbose "Saving to: $installerPath"
        
        # Add timeout and retry for download
        $downloadAttempts = 3
        $success = $false
        
        for ($i = 1; $i -le $downloadAttempts; $i++) {
            try {
                Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath -TimeoutSec 60
                $success = $true
                break
            }
            catch {
                Write-Warning "Download attempt $i failed: $_"
                if ($i -lt $downloadAttempts) {
                    Start-Sleep -Seconds 10
                }
            }
        }
        
        if (-not $success) {
            throw "Failed to download installer after $downloadAttempts attempts"
        }
        
        if (-not (Test-Path $installerPath)) {
            throw "Installer file not found at $installerPath"
        }
        
        Write-Host "Installing Forticlient VPN..."

        # Define installation arguments
        $installArgs = if ($Interactive) {
            @()
        } else {
            @(
                "/S",
                "/v`"/qn REBOOT=ReallySuppress`""
            )
        }
        
        Write-Verbose "Install arguments: $($installArgs -join ' ')"
        
        if (-not $ShowWindow -and -not $Interactive) {
            # Create process start info
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = $installerPath
            $psi.Arguments = $installArgs -join ' '
            $psi.UseShellExecute = $true  # Changed to true
            $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
            $psi.CreateNoWindow = $true

            # Start process with no window
            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $psi
            
            Write-Verbose "Starting installation process with no window"
            [void]$process.Start()
            $process.WaitForExit()
            
            if ($process.ExitCode -ne 0) {
                $errorDescription = Get-InstallerError -ExitCode $process.ExitCode
                Write-LogEntry "Installation failed with exit code $($process.ExitCode): $errorDescription" -Type "Error"
                throw "Installation failed: $errorDescription"
            }
        } else {
            # Interactive or ShowWindow mode
            if ($ShowWindow) {
                Start-Process -FilePath $installerPath -ArgumentList $installArgs -Wait
            } else {
                Start-Process -FilePath $installerPath -ArgumentList $installArgs -Wait -WindowStyle Hidden
            }
        }

        # Add small delay after installation
        Start-Sleep -Seconds 15

        # Wait for installation to complete and verify
        Start-Sleep -Seconds 10  # Initial wait for installer to start
        $timeout = 300  # 5 minutes timeout
        $timer = [Diagnostics.Stopwatch]::StartNew()
        
        while ($timer.Elapsed.TotalSeconds -lt $timeout) {
            if (-not (Get-Process | Where-Object { $_.Name -like "*FortiClient*Installer*" })) {
                Write-Verbose "Installation process completed"
                break
            }
            Start-Sleep -Seconds 5
        }
        
        if ($timer.Elapsed.TotalSeconds -ge $timeout) {
            throw "Installation timed out after 5 minutes"
        }
        
        # Extended wait and verification
        $verificationAttempts = 6
        $verificationDelay = 30
        
        for ($i = 1; $i -le $verificationAttempts; $i++) {
            Write-Verbose "Verification attempt $i of $verificationAttempts"
            if (Test-Installation) {
                Write-Host "Installation completed and verified successfully"
                return
            }
            if ($i -lt $verificationAttempts) {
                Write-Verbose "Waiting $verificationDelay seconds before next verification..."
                Start-Sleep -Seconds $verificationDelay
            }
        }
        
        throw "Installation could not be verified after $verificationAttempts attempts"
    }
    catch {
        Write-LogEntry "Error during installation: $_" -Type "Error"
        Write-LogEntry "Stack trace: $($_.ScriptStackTrace)" -Type "Error"
        $logPath = Join-Path (Get-ScriptDirectory) "FortiClientVPN_Install.log"
        Write-Host "Installation log can be found at: $logPath"
        throw
    }
    finally {
        Write-Verbose "Cleaning up temporary directory: $tempPath"
        Remove-Item -Path $tempPath -Recurse -Force -ErrorAction SilentlyContinue
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
    
    # Write to log file
    try {
        $logMessage | Out-File -FilePath $logPath -Append -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to write to log file: $_"
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
    Write-Host "Starting Forticlient VPN reinstallation..."
    Write-LogEntry "Stopping Forticlient services..." -Type "Info"
    Stop-ForticlientServices
    Write-LogEntry "Uninstalling existing Forticlient..." -Type "Info"
    Uninstall-ExistingForticlient
    Write-LogEntry "Installing new Forticlient VPN..." -Type "Info"
    Install-ForticlientVPN
    Write-LogEntry "=== Forticlient VPN reinstallation process completed ===" -Type "Info"
    if (Test-Installation) {
        Write-LogEntry "FortiClient VPN reinstallation completed and verified" -Type "Info"
        Write-Host "FortiClient VPN reinstallation completed and verified"
    } else {
        Write-LogEntry "FortiClient VPN reinstallation could not be verified" -Type "Error"
        Write-Error "FortiClient VPN reinstallation could not be verified"
    }
}
finally {
    Stop-Transcript
}
