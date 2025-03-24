# =============================================================================
# Script: Get-WizTreePortable.ps1
# Created: 2024-02-08 15:30:00 UTC
# Author: nunya-nunya
# Last Updated: 2024-02-27 17:45:00 UTC
# Updated By: nunya-nunya
# Version: 2.1
# Additional Info: Updated to version 4.24 and added architecture verification
# =============================================================================

<#
.SYNOPSIS
    Downloads and runs the latest version of WizTree Portable silently.
.DESCRIPTION
    This script automatically fetches the latest version of WizTree Portable,
    downloads it, and runs it silently in administrator mode.
.EXAMPLE
    .\Get-WizTreePortable.ps1
#>

function Get-LatestWizTreeUrl {
    try {
        Write-Host "Checking for latest WizTree version..." -ForegroundColor Cyan
        $webResponse = Invoke-WebRequest -Uri "https://wiztree.co.uk/download/" -UseBasicParsing
        if ($webResponse.Content -match 'href="([^"]*wiztree_4_24.*portable\.zip)"') {
            return $Matches[1]
        }
        throw "Could not find download URL"
    }
    catch {
        Write-Warning "Failed to get latest version URL. Using fallback URL..."
        return "https://wiztree.co.uk/wp-content/uploads/2024/02/wiztree_4_24_portable.zip"
    }
}

# Verify x64 architecture
if (-not [Environment]::Is64BitOperatingSystem) {
    Write-Error "This script requires a 64-bit operating system."
    exit 1
}

# Verify running with admin privileges
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "This script requires Administrator privileges. Please run as Administrator."
    exit 1
}

try {
    [string]$downloadUrl = Get-LatestWizTreeUrl
    [string]$zipFilePath = "C:\temp\wiztreeportable.zip"
    [string]$extractPath = "C:\temp\WizTree"
    [string]$exePath = "$extractPath\WizTree64.exe"

    # Create temp directory if it doesn't exist
    if (-Not (Test-Path -Path "C:\temp")) {
        Write-Host "Creating temp directory..." -ForegroundColor Cyan
        New-Item -ItemType Directory -Path "C:\temp" | Out-Null
    }

    # Download WizTree Portable
    Write-Host "Downloading WizTree Portable..." -ForegroundColor Cyan
    Invoke-WebRequest -Uri $downloadUrl -OutFile $zipFilePath -UseBasicParsing

    # Extract the ZIP file
    Write-Host "Extracting files..." -ForegroundColor Cyan
    Expand-Archive -Path $zipFilePath -DestinationPath $extractPath -Force

    # Run WizTree silently as Administrator (ensuring x64 version)
    Write-Host "Starting WizTree x64..." -ForegroundColor Cyan
    Start-Process -FilePath $exePath -Verb RunAs -ArgumentList "/quiet" -WindowStyle Hidden

    Write-Host "WizTree has been successfully launched!" -ForegroundColor Green
}
catch {
    Write-Error "An error occurred: $_"
    exit 1
}
finally {
    # Cleanup
    if (Test-Path $zipFilePath) {
        Remove-Item $zipFilePath -Force
    }
}
