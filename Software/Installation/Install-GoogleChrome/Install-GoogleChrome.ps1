# =============================================================================
# Script: Install-GoogleChrome.ps1
# Created: 2025-05-07 14:31:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-05-07 14:31:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.0
# Additional Info: Initial script creation
# =============================================================================

<#
.SYNOPSIS
Silently installs Google Chrome on a Windows system.

.DESCRIPTION
This script downloads the latest version of Google Chrome Enterprise installer
and performs a silent installation. It includes error handling, logging, and
-WhatIf functionality for testing purposes.

The script performs the following actions:
1. Creates a temporary directory for downloading the installer
2. Downloads the latest Chrome Enterprise MSI installer
3. Installs Chrome silently with specified parameters
4. Logs all activities and errors
5. Cleans up temporary files

.PARAMETER WhatIf
If specified, shows what would happen if the script runs without actually making changes.

.EXAMPLE
.\Install-GoogleChrome.ps1
# Installs Google Chrome silently with default settings

.EXAMPLE
.\Install-GoogleChrome.ps1 -WhatIf
# Shows what the script would do without making actual changes
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param()

# Script variables
$logFile = "$PSScriptRoot\Install-GoogleChrome.log"
$tempDir = "$env:TEMP\ChromeInstall"
$downloadUrl = "https://dl.google.com/edgedl/chrome/install/GoogleChromeEnterpriseBundle64.zip"
$zipFile = "$tempDir\GoogleChromeEnterpriseBundle64.zip"
$msiPath = "$tempDir\Installers\GoogleChromeStandaloneEnterprise64.msi"

# Function to write log entries
function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS", "DEBUG")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [$Level] $Message"

    # Output to console with colors
    switch ($Level) {
        "INFO"    { Write-Host $logEntry -ForegroundColor White }
        "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
        "ERROR"   { Write-Host $logEntry -ForegroundColor Red }
        "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
        "DEBUG"   { Write-Host $logEntry -ForegroundColor Magenta }
    }

    # Write to log file
    Add-Content -Path $logFile -Value $logEntry
}

# Start script
Write-Log "Starting Google Chrome silent installation script" "INFO"
Write-Log "Script version: 1.0.0" "INFO"

try {
    # Create temporary directory if it doesn't exist
    if (-not (Test-Path -Path $tempDir)) {
        if ($PSCmdlet.ShouldProcess("Create temporary directory: $tempDir", "New-Item")) {
            Write-Log "Creating temporary directory: $tempDir" "INFO"
            New-Item -Path $tempDir -ItemType Directory -Force | Out-Null
        }
        else {
            Write-Log "WhatIf: Would create temporary directory: $tempDir" "INFO"
        }
    }

    # Download Chrome Enterprise Bundle
    if ($PSCmdlet.ShouldProcess("Download Chrome Enterprise installer", "Invoke-WebRequest")) {
        Write-Log "Downloading Chrome Enterprise installer from $downloadUrl" "INFO"
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $downloadUrl -OutFile $zipFile
        Write-Log "Download completed successfully" "SUCCESS"
    }
    else {
        Write-Log "WhatIf: Would download Chrome Enterprise installer from $downloadUrl" "INFO"
    }

    # Extract the zip file
    if ($PSCmdlet.ShouldProcess("Extract Chrome Enterprise installer", "Expand-Archive")) {
        Write-Log "Extracting Chrome Enterprise installer" "INFO"
        Expand-Archive -Path $zipFile -DestinationPath $tempDir -Force
        Write-Log "Extraction completed successfully" "SUCCESS"
    }
    else {
        Write-Log "WhatIf: Would extract Chrome Enterprise installer" "INFO"
    }

    # Check if MSI exists
    if (-not ($WhatIfPreference) -and -not (Test-Path -Path $msiPath)) {
        Write-Log "MSI file not found at expected location: $msiPath" "ERROR"
        throw "MSI file not found at expected location: $msiPath"
    }

    # Install Chrome silently
    if ($PSCmdlet.ShouldProcess("Install Google Chrome", "Start-Process msiexec.exe")) {
        Write-Log "Installing Google Chrome silently" "INFO"

        $arguments = "/i `"$msiPath`" /qn /norestart /l*v `"$tempDir\chrome_install_log.txt`""
        $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $arguments -Wait -PassThru

        if ($process.ExitCode -eq 0) {
            Write-Log "Google Chrome installed successfully" "SUCCESS"
        }
        else {
            Write-Log "Installation failed with exit code: $($process.ExitCode)" "ERROR"
            Write-Log "Check log file for details: $tempDir\chrome_install_log.txt" "INFO"
            throw "Installation failed with exit code: $($process.ExitCode)"
        }
    }
    else {
        Write-Log "WhatIf: Would install Google Chrome silently" "INFO"
    }

    # Clean up temporary files
    if ($PSCmdlet.ShouldProcess("Remove temporary files", "Remove-Item")) {
        Write-Log "Cleaning up temporary files" "INFO"
        Remove-Item -Path $tempDir -Recurse -Force
        Write-Log "Cleanup completed successfully" "SUCCESS"
    }
    else {
        Write-Log "WhatIf: Would remove temporary files" "INFO"
    }

    Write-Log "Google Chrome installation completed successfully" "SUCCESS"
}
catch {
    Write-Log "An error occurred: $_" "ERROR"
    Write-Log "Exception details: $($_.Exception.Message)" "ERROR"
    Write-Log "Stack trace: $($_.ScriptStackTrace)" "DEBUG"

    # Ensure we attempt to clean up even if there was an error
    if (Test-Path -Path $tempDir) {
        try {
            if ($PSCmdlet.ShouldProcess("Remove temporary files after error", "Remove-Item")) {
                Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                Write-Log "Cleaned up temporary files after error" "INFO"
            }
        }
        catch {
            Write-Log "Failed to clean up temporary files: $_" "WARNING"
        }
    }

    exit 1
}
