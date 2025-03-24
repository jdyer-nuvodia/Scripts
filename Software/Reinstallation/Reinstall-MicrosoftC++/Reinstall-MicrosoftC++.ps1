# =============================================================================
# Script: Reinstall-MicrosoftC++.ps1
# Created: 2025-02-27 18:51:00 UTC
# Author: nunya-nunya
# Last Updated: 2025-02-27 18:51:00 UTC
# Updated By: nunya-nunya
# Version: 1.0
# Additional Info: Initial script with standard header format
# =============================================================================

<#
.SYNOPSIS
    Downloads and installs Microsoft Visual C++ Redistributables (x86 and x64).
.DESCRIPTION
    This script automates the process of downloading and installing the latest Microsoft Visual C++ Redistributables.
    Key actions:
     - Creates a temporary directory for downloads
     - Downloads both x86 and x64 versions of Visual C++ Redistributables
     - Installs the redistributables silently
     - No system restart is forced after installation
    Dependencies:
     - Requires internet connection
     - Requires administrative privileges
.EXAMPLE
    .\Reinstall-MicrosoftC++.ps1
    Downloads and installs both x86 and x64 Visual C++ Redistributables
.NOTES
    Security Level: Medium
    Required Permissions: Administrative privileges
    Validation Requirements: Verify successful installation in Programs and Features
#>

# Define the path where the redistributable installers will be saved
$downloadPath = "$env:TEMP\Redistributables"

# Create the download directory if it doesn't exist
if (!(Test-Path -Path $downloadPath)) {
    New-Item -ItemType Directory -Path $downloadPath
}

# URLs for the latest redistributables
$urls = @(
    "https://aka.ms/vs/17/release/vc_redist.x86.exe",
    "https://aka.ms/vs/17/release/vc_redist.x64.exe"
)

# Filenames for the redistributables
$filenames = @(
    "vc_redist.x86.exe",
    "vc_redist.x64.exe"
)

# Download the redistributables
for ($i = 0; $i -lt $urls.Count; $i++) {
    Write-Host "Downloading $($filenames[$i])..."
    Invoke-WebRequest -Uri $urls[$i] -OutFile "$downloadPath\$($filenames[$i])"
}

# Install the redistributables
foreach ($filename in $filenames) {
    Write-Host "Installing $filename..."
    Start-Process -FilePath "$downloadPath\$filename" -ArgumentList "/install /passive /norestart" -Wait
}

Write-Host "Installation complete."
