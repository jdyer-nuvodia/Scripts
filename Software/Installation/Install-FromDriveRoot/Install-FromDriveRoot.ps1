# =============================================================================
# Script: Install-FromDriveRoot.ps1
# Created: 2025-05-07 21:45:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-06-06 17:08:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.2
# Additional Info: Replaced Write-Host with Write-Information for automated execution
# =============================================================================

<#
.SYNOPSIS
Automatically finds and installs the first EXE file located in the root of the drive.

.DESCRIPTION
This script identifies the first EXE file in the root directory of the drive containing
the script, then runs it with silent installation parameters. It is designed for use
on removable media or network shares where software installers are placed in the root.

The script performs the following actions:
1. Determines the drive root where the script is located
2. Searches for the first .exe file in that root directory
3. Runs the installer with silent parameters
4. Returns appropriate exit codes based on the installation result

Dependencies:
- Must be run with appropriate permissions to execute the installer
- Assumes silent install is supported by the target EXE

.PARAMETER WhatIf
If specified, shows what would happen if the script runs without actually making changes.

.EXAMPLE
.\Install-FromDriveRoot.ps1
# Finds and runs the first EXE in the drive root with silent parameters

.EXAMPLE
.\Install-FromDriveRoot.ps1 -WhatIf
# Shows what installer would be run without actually installing
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param()

# Get the directory where the script is located
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$driveRoot = (Get-Item $scriptDir).PSDrive.Root

# Create log file in the same directory as the script
$computerName = $env:COMPUTERNAME
$utcTimestamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd_HH-mm-ss")
$logFile = Join-Path -Path $scriptDir -ChildPath "Install-FromDriveRoot_${computerName}_${utcTimestamp}.log"

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
    $logEntry = "$timestamp [$Level] $Message"    # Output to console for automated execution
    switch ($Level) {
        "INFO"    { Write-Information $logEntry -InformationAction Continue }
        "WARNING" { Write-Warning $logEntry }
        "ERROR"   { Write-Error $logEntry }
        "SUCCESS" { Write-Information $logEntry -InformationAction Continue }
        "DEBUG"   { Write-Verbose $logEntry }
        Default   { Write-Information $logEntry -InformationAction Continue }
    }

    # Create log directory if it does not exist
    $logDir = Split-Path -Path $logFile -Parent
    if (-not [string]::IsNullOrEmpty($logDir) -and -not (Test-Path -Path $logDir)) {
        try {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        catch {
            # Log creation failed, but continue execution
            Write-Verbose "Could not create log directory: $($_.Exception.Message)"
        }
    }

    # Write to log file
    if (-not [string]::IsNullOrEmpty($logFile)) {
        try {
            Add-Content -Path $logFile -Value $logEntry -ErrorAction SilentlyContinue
        }
        catch {
            # Silently continue if we cannot write to log file
            Write-Verbose "Could not write to log file: $($_.Exception.Message)"
        }
    }
}

Write-Log "Script started. Looking for installers in $driveRoot" "INFO"

# Find the first .exe file in the root of the drive
$installer = Get-ChildItem -Path $driveRoot -Filter *.exe | Select-Object -First 1

if ($null -eq $installer) {
    Write-Log "No installer found in the drive root." "ERROR"
    exit 1
} else {
    Write-Log "Found installer: $($installer.FullName)" "INFO"
    # Adjust the silent switch as needed for your installer
    $arguments = "/silent"

    if ($PSCmdlet.ShouldProcess($installer.FullName, "Run installer with arguments: $arguments")) {
        Write-Log "Starting installation process with arguments: $arguments" "INFO"
        $process = Start-Process -FilePath $installer.FullName -ArgumentList $arguments -Wait -PassThru

        if ($process.ExitCode -eq 0) {
            Write-Log "Installation completed successfully." "SUCCESS"
            exit 0
        } else {
            Write-Log "Installer exited with code $($process.ExitCode)." "ERROR"
            exit $process.ExitCode
        }
    } else {
        Write-Log "WhatIf: Would have run installer: $($installer.FullName) with arguments: $arguments" "INFO"
    }
}
