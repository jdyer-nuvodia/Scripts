# =============================================================================
# Script: Get-WizTreePortable.ps1
# Created: 2024-02-08 15:30:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-27 17:30:00 UTC
# Updated By: jdyer-nuvodia
# Version: 2.0
# Additional Info: Added silent installation and latest version check
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
        if ($webResponse.Content -match 'href="([^"]*wiztree[^"]*portable\.zip)"') {
            return $Matches[1]
        }
        throw "Could not find download URL"
    }
    catch {
        Write-Warning "Failed to get latest version URL. Using fallback URL..."
        return "https://wiztree.co.uk/wp-content/uploads/2024/05/wiztree_4_19_portable.zip"
    }
}

# Define variables
$downloadUrl = Get-LatestWizTreeUrl
$zipFilePath = "C:\temp\wiztreeportable.zip"
$extractPath = "C:\temp\WizTree"
$exePath = "$extractPath\WizTree64.exe"

try {
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

    # Run WizTree silently as Administrator
    Write-Host "Starting WizTree..." -ForegroundColor Cyan
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
