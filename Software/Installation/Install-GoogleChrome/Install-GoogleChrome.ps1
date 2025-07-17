# =============================================================================
# Script: Install-GoogleChrome.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1.1
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
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

# Color configuration for PowerShell compatibility
$Script:UseAnsiColors = $PSVersionTable.PSVersion.Major -ge 7
if ($Script:UseAnsiColors) {
    # PowerShell 7+ ANSI escape codes
    $Script:Colors = @{
        "White"    = "`e[37m"
        "Cyan"     = "`e[36m"
        "Green"    = "`e[32m"
        "Yellow"   = "`e[33m"
        "Red"      = "`e[31m"
        "Magenta"  = "`e[35m"
        "DarkGray" = "`e[90m"
        "Reset"    = "`e[0m"
    }
} else {
    # PowerShell 5.1 console colors
    $Script:Colors = @{
        "White"    = [System.ConsoleColor]::White
        "Cyan"     = [System.ConsoleColor]::Cyan
        "Green"    = [System.ConsoleColor]::Green
        "Yellow"   = [System.ConsoleColor]::Yellow
        "Red"      = [System.ConsoleColor]::Red
        "Magenta"  = [System.ConsoleColor]::Magenta
        "DarkGray" = [System.ConsoleColor]::DarkGray
        "Reset"    = ""
    }
}

# Function to write colored output
function Write-ColorOutput {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$Color = "White"
    )

    if ($Script:UseAnsiColors) {
        # PowerShell 7+ with ANSI escape codes
        $colorCode = $Script:Colors[$Color]
        $resetCode = $Script:Colors.Reset
        Write-Output "${colorCode}${Message}${resetCode}"
    } else {
        # PowerShell 5.1 - Change console color, write output, then reset
        $originalColor = $Host.UI.RawUI.ForegroundColor
        try {
            if ($Script:Colors[$Color] -and $Script:Colors[$Color] -ne "") {
                $Host.UI.RawUI.ForegroundColor = $Script:Colors[$Color]
            }
            Write-Output $Message
        } finally {
            $Host.UI.RawUI.ForegroundColor = $originalColor
        }
    }
}

# Function to write log entries
function Write-ScriptLog {
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
        "INFO" { Write-ColorOutput -Message $logEntry -Color "White" }
        "WARNING" { Write-ColorOutput -Message $logEntry -Color "Yellow" }
        "ERROR" { Write-ColorOutput -Message $logEntry -Color "Red" }
        "SUCCESS" { Write-ColorOutput -Message $logEntry -Color "Green" }
        "DEBUG" { Write-ColorOutput -Message $logEntry -Color "Magenta" }
    }

    # Write to log file
    Add-Content -Path $logFile -Value $logEntry
}

# Start script
Write-ScriptLog "Starting Google Chrome silent installation script" "INFO"
Write-ScriptLog "Script version: 1.1.0" "INFO"

try {
    # Create temporary directory if it doesn't exist
    if (-not (Test-Path -Path $tempDir)) {
        if ($PSCmdlet.ShouldProcess("Create temporary directory: $tempDir", "New-Item")) {
            Write-ScriptLog "Creating temporary directory: $tempDir" "INFO"
            New-Item -Path $tempDir -ItemType Directory -Force | Out-Null
        } else {
            Write-ScriptLog "WhatIf: Would create temporary directory: $tempDir" "INFO"
        }
    }

    # Download Chrome Enterprise Bundle
    if ($PSCmdlet.ShouldProcess("Download Chrome Enterprise installer", "Invoke-WebRequest")) {
        Write-ScriptLog "Downloading Chrome Enterprise installer from $downloadUrl" "INFO"
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $downloadUrl -OutFile $zipFile
        Write-ScriptLog "Download completed successfully" "SUCCESS"
    } else {
        Write-ScriptLog "WhatIf: Would download Chrome Enterprise installer from $downloadUrl" "INFO"
    }

    # Extract the zip file
    if ($PSCmdlet.ShouldProcess("Extract Chrome Enterprise installer", "Expand-Archive")) {
        Write-ScriptLog "Extracting Chrome Enterprise installer" "INFO"
        Expand-Archive -Path $zipFile -DestinationPath $tempDir -Force
        Write-ScriptLog "Extraction completed successfully" "SUCCESS"
    } else {
        Write-ScriptLog "WhatIf: Would extract Chrome Enterprise installer" "INFO"
    }

    # Check if MSI exists
    if (-not ($WhatIfPreference) -and -not (Test-Path -Path $msiPath)) {
        Write-ScriptLog "MSI file not found at expected location: $msiPath" "ERROR"
        throw "MSI file not found at expected location: $msiPath"
    }

    # Install Chrome silently
    if ($PSCmdlet.ShouldProcess("Install Google Chrome", "Start-Process msiexec.exe")) {
        Write-ScriptLog "Installing Google Chrome silently" "INFO"

        $arguments = "/i `"$msiPath`" /qn /norestart /l*v `"$tempDir\chrome_install_log.txt`""
        $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $arguments -Wait -PassThru

        if ($process.ExitCode -eq 0) {
            Write-ScriptLog "Google Chrome installed successfully" "SUCCESS"
        } else {
            Write-ScriptLog "Installation failed with exit code: $($process.ExitCode)" "ERROR"
            Write-ScriptLog "Check log file for details: $tempDir\chrome_install_log.txt" "INFO"
            throw "Installation failed with exit code: $($process.ExitCode)"
        }
    } else {
        Write-ScriptLog "WhatIf: Would install Google Chrome silently" "INFO"
    }

    # Clean up temporary files
    if ($PSCmdlet.ShouldProcess("Remove temporary files", "Remove-Item")) {
        Write-ScriptLog "Cleaning up temporary files" "INFO"
        Remove-Item -Path $tempDir -Recurse -Force
        Write-ScriptLog "Cleanup completed successfully" "SUCCESS"
    } else {
        Write-ScriptLog "WhatIf: Would remove temporary files" "INFO"
    }

    Write-ScriptLog "Google Chrome installation completed successfully" "SUCCESS"
} catch {
    Write-ScriptLog "An error occurred: $_" "ERROR"
    Write-ScriptLog "Exception details: $($_.Exception.Message)" "ERROR"
    Write-ScriptLog "Stack trace: $($_.ScriptStackTrace)" "DEBUG"

    # Ensure we attempt to clean up even if there was an error
    if (Test-Path -Path $tempDir) {
        try {
            if ($PSCmdlet.ShouldProcess("Remove temporary files after error", "Remove-Item")) {
                Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                Write-ScriptLog "Cleaned up temporary files after error" "INFO"
            }
        } catch {
            Write-ScriptLog "Failed to clean up temporary files: $_" "WARNING"
        }
    }

    exit 1
}
