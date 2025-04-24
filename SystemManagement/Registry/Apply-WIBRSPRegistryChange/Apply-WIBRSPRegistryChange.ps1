# =============================================================================
# Script: Apply-WIBRSPRegistryChange.ps1
# Created: 2025-04-24 13:45:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-24 13:45:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.0
# Additional Info: Initial creation of script to apply SharePoint_Auto_Mount.reg
# =============================================================================

<#
.SYNOPSIS
Applies registry changes from a .reg file in the same directory.

.DESCRIPTION
This script imports registry settings from a file named 'SharePoint_Auto_Mount.reg' located
in the same directory as this script. It verifies the file exists, validates it has a .reg
extension, and applies the registry changes with detailed logging and error handling.
The script includes -WhatIf functionality to preview changes before applying them.

.PARAMETER WhatIf
Shows what would happen if the script runs without actually applying the changes.

.EXAMPLE
.\Apply-RegistryChange.ps1
Applies registry changes from the SharePoint_Auto_Mount.reg file.

.EXAMPLE
.\Apply-RegistryChange.ps1 -WhatIf
Shows the registry file that would be imported without actually applying the changes.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param()

# Set error action preference
$ErrorActionPreference = "Stop"

# Define log file path in the same directory as script
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$logFile = Join-Path -Path $scriptPath -ChildPath "Apply-RegistryChange_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$regFilePath = Join-Path -Path $scriptPath -ChildPath "SharePoint_Auto_Mount.reg"

# Function to write to log file and console with color-coding
function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "PROCESS", "SUCCESS", "WARNING", "ERROR", "DEBUG", "DETAIL")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Write to log file
    Add-Content -Path $logFile -Value $logMessage
    
    # Write to console with color-coding
    switch ($Level) {
        "INFO"    { Write-Host $logMessage -ForegroundColor White }
        "PROCESS" { Write-Host $logMessage -ForegroundColor Cyan }
        "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        "ERROR"   { Write-Host $logMessage -ForegroundColor Red }
        "DEBUG"   { Write-Host $logMessage -ForegroundColor Magenta }
        "DETAIL"  { Write-Host $logMessage -ForegroundColor DarkGray }
    }
}

try {
    # Log script start
    Write-Log "Starting registry change application script" "INFO"
    Write-Log "Script version: 1.0.0" "DETAIL"
    
    # Check if registry file exists
    if (-not (Test-Path -Path $regFilePath)) {
        Write-Log "Registry file not found: $regFilePath" "ERROR"
        throw "Registry file 'SharePoint_Auto_Mount.reg' not found in script directory."
    }
    
    # Verify file has .reg extension
    if ([System.IO.Path]::GetExtension($regFilePath) -ne ".reg") {
        Write-Log "File does not have .reg extension: $regFilePath" "ERROR"
        throw "The specified file is not a registry (.reg) file."
    }
    
    # Log file details
    $fileInfo = Get-Item -Path $regFilePath
    Write-Log "Registry file found: $($fileInfo.Name)" "PROCESS"
    Write-Log "File size: $([Math]::Round($fileInfo.Length / 1KB, 2)) KB" "DETAIL"
    Write-Log "Last modified: $($fileInfo.LastWriteTime)" "DETAIL"
    
    # Preview first 5 lines of registry file (for information purposes)
    Write-Log "Registry file preview:" "DETAIL"
    Get-Content -Path $regFilePath -TotalCount 5 | ForEach-Object {
        Write-Log "  $_" "DETAIL"
    }
    Write-Log "..." "DETAIL"
    
    # Import registry file
    if ($PSCmdlet.ShouldProcess("Registry", "Apply changes from $regFilePath")) {
        Write-Log "Applying registry changes from file..." "PROCESS"
        
        # Use reg.exe to import the registry file
        $regImportOutput = reg.exe import $regFilePath 2>&1
        
        if ($LASTEXITCODE -eq 0) {
            Write-Log "Registry changes applied successfully" "SUCCESS"
        }
        else {
            Write-Log "Failed to apply registry changes. Exit code: $LASTEXITCODE" "ERROR"
            Write-Log "Error details: $regImportOutput" "ERROR"
            throw "Failed to apply registry changes: $regImportOutput"
        }
    }
    else {
        Write-Log "WhatIf: Would apply registry changes from: $regFilePath" "PROCESS"
    }
}
catch {
    Write-Log "An error occurred: $_" "ERROR"
    Write-Log "Exception details: $($_.Exception)" "DEBUG"
    exit 1
}
finally {
    Write-Log "Script execution completed" "INFO"
}
