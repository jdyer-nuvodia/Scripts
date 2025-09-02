# =============================================================================
# Script: Remove-AdobeAcrobatReader.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2.1
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
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

    Supports -WhatIf parameter to preview changes without making them.

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

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
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

# Run this script as an administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You do not have Administrator rights to run this script!`nPlease re-run this script as an Administrator!"
    break
}

# Function to write log messages
function Write-LogMessage {
    param([string]$Message)
    $logPath = "C:\Temp\AdobeReaderRemoval_$(Get-Date -Format 'yyyyMMdd').log"
    $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message"
    Add-Content -Path $logPath -Value $logMessage
    Write-ColorOutput -Message $logMessage -Color "White"
}

Write-LogMessage "Starting Adobe Acrobat Reader removal process"

# Uninstall Adobe Acrobat Reader using MSI
$uninstallKeys = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
)

$foundInstallations = $false

foreach ($key in $uninstallKeys) {
    Write-LogMessage "Searching for Adobe Reader in $key"
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
            Write-LogMessage "Found installation: $displayName with product code: $productCode"
            if ($PSCmdlet.ShouldProcess($displayName, "Uninstall using MSI")) {
                try {
                    $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $productCode /qn" -Wait -NoNewWindow -PassThru
                    if ($process.ExitCode -eq 0) {
                        Write-LogMessage "Successfully uninstalled $displayName"
                    } else {
                        Write-LogMessage "Failed to uninstall $displayName. Exit code: $($process.ExitCode)"
                    }
                } catch {
                    Write-LogMessage "Error occurred while uninstalling $displayName. Error: $_"
                }
            }
        }
    } else {
        Write-LogMessage "No Adobe Reader installations found in $key"
    }
}

if (-not $foundInstallations) {
    Write-LogMessage "No Adobe Acrobat Reader installations were found to uninstall."
}

# Remove Adobe Acrobat Reader directories
$directories = @(
    "${ env:ProgramFiles}\Adobe\Acrobat Reader",
    "${ env:ProgramFiles(x86)}\Adobe\Acrobat Reader",
    "${ env:APPDATA}\Adobe\Acrobat",
    "${ env:LOCALAPPDATA}\Adobe\Acrobat"
)

foreach ($dir in $directories) {
    if (Test-Path $dir) {
        if ($PSCmdlet.ShouldProcess($dir, "Remove directory")) {
            try {
                Remove-Item -Path $dir -Recurse -Force
                Write-LogMessage "Removed directory: $dir"
            } catch {
                Write-LogMessage "Failed to remove directory $dir. Error: $_"
            }
        }
    } else {
        Write-LogMessage "Directory not found: $dir"
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

Write-LogMessage "Searching for Adobe Reader/Acrobat registry keys:"
foreach ($key in $adobeKeys) {
    if (Test-Path $key) {
        $versions = Get-ChildItem -Path $key -ErrorAction SilentlyContinue
        foreach ($version in $versions) {
            $versionNumber = $version.PSChildName
            Write-LogMessage "Found Adobe Reader/Acrobat version: $versionNumber in $key"

            if ($PSCmdlet.ShouldProcess("$key\$versionNumber", "Remove registry key")) {
                try {
                    Remove-Item -Path $version.PSPath -Recurse -Force
                    Write-LogMessage "Removed registry key: $($version.PSPath)"
                } catch {
                    Write-LogMessage "Failed to remove registry key $($version.PSPath). Error: $_"
                }
            }
        }
    } else {
        Write-LogMessage "Registry key not found: $key"
    }
}

# Remove Creative Cloud Files shortcut from Explorer
$ccfPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{ 0E270DAA-1BE6-48F2-AC49-5CE0DBECC398}"
if (Test-Path $ccfPath) {
    if ($PSCmdlet.ShouldProcess("Creative Cloud Files shortcut", "Remove from Explorer")) {
        try {
            Remove-Item -Path $ccfPath -Recurse -Force
            Write-LogMessage "Removed Creative Cloud Files shortcut from Explorer"
        } catch {
            Write-LogMessage "Failed to remove Creative Cloud Files shortcut. Error: $_"
        }
    }
} else {
    Write-LogMessage "Creative Cloud Files shortcut not found"
}

Write-LogMessage "Adobe Acrobat Reader removal process completed"
