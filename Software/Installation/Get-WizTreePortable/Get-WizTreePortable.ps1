# =============================================================================
# Script: Get-WizTreePortable.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 2.3.2
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
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

# Initialize color support for cross-platform compatibility
$script:UseAnsiColors = $PSVersionTable.PSVersion.Major -ge 7
$script:Colors = @{
    Reset    = if ($script:UseAnsiColors) { "`e[0m" } else { "" }
    White    = if ($script:UseAnsiColors) { "`e[37m" } else { "White" }
    Cyan     = if ($script:UseAnsiColors) { "`e[36m" } else { "Cyan" }
    Green    = if ($script:UseAnsiColors) { "`e[32m" } else { "Green" }
    Yellow   = if ($script:UseAnsiColors) { "`e[33m" } else { "Yellow" }
    Red      = if ($script:UseAnsiColors) { "`e[31m" } else { "Red" }
    Magenta  = if ($script:UseAnsiColors) { "`e[35m" } else { "Magenta" }
    DarkGray = if ($script:UseAnsiColors) { "`e[90m" } else { "DarkGray" }
}

function Write-ColorOutput {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$Color = "White"
    )

    if ($script:UseAnsiColors) {
        # PowerShell 7+ with ANSI escape codes
        $colorCode = $script:Colors[$Color]
        $resetCode = $script:Colors.Reset
        Write-Output "${colorCode}${Message}${resetCode}"
    } else {
        # PowerShell 5.1 - Change console color, write output, then reset
        $originalColor = $Host.UI.RawUI.ForegroundColor
        try {
            if ($script:Colors[$Color] -and $script:Colors[$Color] -ne "") {
                $Host.UI.RawUI.ForegroundColor = $script:Colors[$Color]
            }
            Write-Output $Message
        } finally {
            $Host.UI.RawUI.ForegroundColor = $originalColor
        }
    }
}

function Get-LatestWizTreeUrl {
    try {
        Write-ColorOutput -Message "Checking for latest WizTree version..." -Color "Cyan"
        $webResponse = Invoke-WebRequest -Uri "https://wiztree.co.uk/download/" -UseBasicParsing
        $pattern = 'href = "([^"]*wiztree_\d+_\d+.*portable\.zip)"'
        if ($webResponse.Content -match $pattern) {
            Write-ColorOutput -Message "Found latest version URL" -Color "Green"
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
    if (-not (Test-Path -Path "C:\temp")) {
        Write-ColorOutput -Message "Creating temp directory..." -Color "Cyan"
        New-Item -ItemType Directory -Path "C:\temp" | Out-Null
    }

    # Download WizTree Portable
    Write-ColorOutput -Message "Downloading WizTree Portable..." -Color "Cyan"
    Invoke-WebRequest -Uri $downloadUrl -OutFile $zipFilePath -UseBasicParsing

    # Extract the ZIP file
    Write-ColorOutput -Message "Extracting files..." -Color "Cyan"
    Expand-Archive -Path $zipFilePath -DestinationPath $extractPath -Force

    # Run WizTree as Administrator (ensuring x64 version)
    Write-ColorOutput -Message "Starting WizTree x64..." -Color "Cyan"
    Start-Process -FilePath $exePath -Verb RunAs

    Write-ColorOutput -Message "WizTree has been successfully launched!" -Color "Green"
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
