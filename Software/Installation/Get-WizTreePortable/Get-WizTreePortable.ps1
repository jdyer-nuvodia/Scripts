# =============================================================================
# Script: Get-WizTreePortable.ps1
# Created: 2025-02-08 15:30:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-02 15:07:00 UTC
# Updated By: jdyer-nuvodia
# Version: 2.2.1
# Additional Info: Modified to start WizTree with visible window
# =============================================================================

<#
.SYNOPSIS
    Downloads and runs the latest version of WizTree Portable.
.DESCRIPTION
    This script automatically fetches the latest version of WizTree Portable,
    downloads it, and runs it in administrator mode with a visible window.
    Automatically detects the latest version from the WizTree website.
.EXAMPLE
    .\Get-WizTreePortable.ps1
#>

function Get-LatestWizTreeUrl {
    try {
        Write-Host "Checking for latest WizTree version..." -ForegroundColor Cyan
        $webResponse = Invoke-WebRequest -Uri "https://wiztree.co.uk/download/" -UseBasicParsing
        $pattern = 'href="([^"]*wiztree_\d+_\d+.*portable\.zip)"'
        if ($webResponse.Content -match $pattern) {
            Write-Host "Found latest version URL" -ForegroundColor Green
            return $Matches[1]
        }
        throw "Could not find download URL"
    }
    catch {
        Write-Warning "Failed to get latest version URL: $_"
        Write-Warning "Please check https://wiztree.co.uk/download/ for the latest version"
        exit 1
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

    # Run WizTree as Administrator (ensuring x64 version)
    Write-Host "Starting WizTree x64..." -ForegroundColor Cyan
    Start-Process -FilePath $exePath -Verb RunAs

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
