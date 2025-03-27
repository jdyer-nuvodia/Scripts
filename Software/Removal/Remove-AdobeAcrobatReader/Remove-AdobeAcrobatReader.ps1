# =============================================================================
# Script: Remove-AdobeAcrobatReader.ps1
# Created: 2025-02-27 18:52:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-27 18:52:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0
# Additional Info: Initial script with standardized header
# =============================================================================

<#
.SYNOPSIS
    Removes Adobe Acrobat Reader and related components from a Windows system.
.DESCRIPTION
    This script performs a comprehensive removal of Adobe Acrobat Reader by:
     - Uninstalling via MSI product codes
     - Removing associated directories
     - Cleaning registry entries
     - Removing Creative Cloud Files shortcuts
     
    Dependencies:
     - Must be run with administrative privileges
     - Windows PowerShell 5.1 or later
     
    Security considerations:
     - Requires registry modification permissions
     - Requires file system modification permissions
     
    Performance impact:
     - Minimal system impact
     - May take several minutes depending on installed versions
.EXAMPLE
    .\Remove-AdobeAcrobatReader.ps1
    Performs complete removal of Adobe Acrobat Reader from the system.
.NOTES
    Security Level: Medium
    Required Permissions: Local Administrator
    Validation Requirements: 
    - Verify Adobe Reader is uninstalled
    - Check for removal of specified directories
    - Validate registry cleanup
#>

# Run this script as an administrator
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))  
{  
    Write-Warning "You do not have Administrator rights to run this script!`nPlease re-run this script as an Administrator!"
    Break
}

# Function to write log messages
function Write-Log {
    param([string]$Message)
    $logPath = "C:\Temp\AdobeReaderRemoval_$(Get-Date -Format 'yyyyMMdd').log"
    $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message"
    Add-Content -Path $logPath -Value $logMessage
    Write-Host $logMessage
}

Write-Log "Starting Adobe Acrobat Reader removal process"

# Uninstall Adobe Acrobat Reader using MSI
$uninstallKeys = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
)

$foundInstallations = $false

foreach ($key in $uninstallKeys) {
    Write-Log "Searching for Adobe Reader in $key"
    $adobeReaderEntries = Get-ChildItem -Path $key -ErrorAction SilentlyContinue | 
        Get-ItemProperty | 
        Where-Object { 
            $_.DisplayName -like "*Adobe Acrobat Reader*" -or
            $_.DisplayName -like "*Adobe Reader*" -and 
            $_.DisplayName -notlike "*Standard*" -and 
            $_.DisplayName -notlike "*Professional*" 
        }

    if ($adobeReaderEntries) {
        $foundInstallations = $true
        foreach ($entry in $adobeReaderEntries) {
            $productCode = $entry.PSChildName
            $displayName = $entry.DisplayName
            Write-Log "Found installation: $displayName with product code: $productCode"
            try {
                $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $productCode /qn" -Wait -NoNewWindow -PassThru
                if ($process.ExitCode -eq 0) {
                    Write-Log "Successfully uninstalled $displayName"
                } else {
                    Write-Log "Failed to uninstall $displayName. Exit code: $($process.ExitCode)"
                }
            } catch {
                Write-Log "Error occurred while uninstalling $displayName. Error: $_"
            }
        }
    } else {
        Write-Log "No Adobe Reader installations found in $key"
    }
}

if (-not $foundInstallations) {
    Write-Log "No Adobe Acrobat Reader installations were found to uninstall."
}

# Remove Adobe Acrobat Reader directories
$directories = @(
    "${env:ProgramFiles}\Adobe\Acrobat Reader",
    "${env:ProgramFiles(x86)}\Adobe\Acrobat Reader",
    "${env:APPDATA}\Adobe\Acrobat",
    "${env:LOCALAPPDATA}\Adobe\Acrobat"
)

foreach ($dir in $directories) {
    if (Test-Path $dir) {
        try {
            Remove-Item -Path $dir -Recurse -Force
            Write-Log "Removed directory: $dir"
        } catch {
            Write-Log "Failed to remove directory $dir. Error: $_"
        }
    } else {
        Write-Log "Directory not found: $dir"
    }
}

# Comprehensive search for Adobe Reader registry keys
$adobeKeys = @(
    "HKLM:\SOFTWARE\Adobe\Acrobat Reader",
    "HKLM:\SOFTWARE\WOW6432Node\Adobe\Acrobat Reader",
    "HKCU:\Software\Adobe\Acrobat Reader",
    "HKLM:\SOFTWARE\Adobe\Adobe Acrobat",
    "HKLM:\SOFTWARE\WOW6432Node\Adobe\Adobe Acrobat",
    "HKCU:\Software\Adobe\Adobe Acrobat"
)

Write-Log "Searching for Adobe Reader/Acrobat registry keys:"
foreach ($key in $adobeKeys) {
    if (Test-Path $key) {
        $versions = Get-ChildItem -Path $key -ErrorAction SilentlyContinue
        foreach ($version in $versions) {
            $versionNumber = $version.PSChildName
            Write-Log "Found Adobe Reader/Acrobat version: $versionNumber in $key"
            
            $subKeys = Get-ChildItem -Path $version.PSPath -Recurse -ErrorAction SilentlyContinue
            foreach ($subKey in $subKeys) {
                $values = Get-ItemProperty -Path $subKey.PSPath -ErrorAction SilentlyContinue
                foreach ($value in $values.PSObject.Properties) {
                    Write-Log "  $($subKey.Name)\$($value.Name): $($value.Value)"
                }
            }
            
            try {
                Remove-Item -Path $version.PSPath -Recurse -Force
                Write-Log "Removed registry key: $($version.PSPath)"
            } catch {
                Write-Log "Failed to remove registry key $($version.PSPath). Error: $_"
            }
        }
    } else {
        Write-Log "Registry key not found: $key"
    }
}

# Remove Creative Cloud Files shortcut from Explorer
$ccfPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{0E270DAA-1BE6-48F2-AC49-5CE0DBECC398}"
if (Test-Path $ccfPath) {
    try {
        Remove-Item -Path $ccfPath -Recurse -Force
        Write-Log "Removed Creative Cloud Files shortcut from Explorer"
    } catch {
        Write-Log "Failed to remove Creative Cloud Files shortcut. Error: $_"
    }
} else {
    Write-Log "Creative Cloud Files shortcut not found"
}

Write-Log "Adobe Acrobat Reader removal process completed"
