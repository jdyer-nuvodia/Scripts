# =============================================================================
# Script: Reinstall-OneDrive.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-15 23:30:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.1
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
# =============================================================================


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

<#
.SYNOPSIS
    Uninstalls existing OneDrive installations and installs the latest version.
.DESCRIPTION
    This script performs the following actions:
     - Uninstalls all existing OneDrive installations
     - Downloads and installs the latest version of OneDrive
     - Includes timeout protection for long-running operations
     - Requires administrator privileges
.EXAMPLE
    .\Reinstall-OneDrive.ps1
    Completely reinstalls OneDrive with the latest version
.NOTES
    Security Level: Medium
    Required Permissions: Administrator
    Validation Requirements: Verify OneDrive is running after installation
#>

# Requires administrator privileges
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-ColorOutput -Message "Script is not running as Administrator. Restarting with elevated privileges..." -Color 'Yellow'
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-File `"$($MyInvocation.MyCommand.Path)`""
    exit
}

# Main script execution with timeout
$job = Start-Job -ScriptBlock {
    # Function to uninstall OneDrive
    function Uninstall-OneDrive {
        Write-ColorOutput -Message "Stopping OneDrive processes..." -Color 'Cyan'
        Stop-Process -Name OneDrive* -Force -ErrorAction SilentlyContinue

        $oneDrivePaths = @(
            "$env:SystemRoot\System32\OneDriveSetup.exe",
            "$env:SystemRoot\SysWOW64\OneDriveSetup.exe",
            "${ env:ProgramFiles}\Microsoft OneDrive\OneDriveSetup.exe",
            "${ env:ProgramFiles(x86)}\Microsoft OneDrive\OneDriveSetup.exe"
        )

        foreach ($StartPath in $oneDrivePaths) {
            if (Test-Path $StartPath) {
                Write-ColorOutput -Message "Uninstalling OneDrive from $StartPath" -Color 'Cyan'
                Start-Process $StartPath -ArgumentList "/uninstall" -Wait
            }
        }

        Write-ColorOutput -Message "Uninstalling OneDrive using WinGet..." -Color 'Cyan'
        winget uninstall Microsoft.OneDrive
    }

    # Function to download and install the latest OneDrive
    function Install-LatestOneDrive {
        $url = "https://go.microsoft.com/fwlink/p/?LinkID = 2182910"
        $outPath = "$env:TEMP\OneDriveSetup.exe"

        Write-ColorOutput -Message "Downloading the latest OneDrive installer..." -Color 'Cyan'
        Invoke-WebRequest -Uri $url -OutFile $outPath

        Write-ColorOutput -Message "Installing the latest version of OneDrive..." -Color 'Cyan'
        Start-Process $outPath -ArgumentList "/allusers" -Wait
    }

    # Execute the functions
    Write-ColorOutput -Message "Starting OneDrive reinstallation process..." -Color 'Cyan'
    Uninstall-OneDrive

    Write-ColorOutput -Message "Installing the latest version of OneDrive..." -Color 'Cyan'
    Install-LatestOneDrive

    Write-ColorOutput -Message "OneDrive update process completed successfully." -Color 'Green'
}

# 5 minutes in seconds
$timeout = 300
$completed = Wait-Job $job -Timeout $timeout

if ($null -eq $completed) {
    Write-ColorOutput -Message "The script did not complete within 5 minutes. Stopping the process..." -Color 'Red'
    Stop-Job $job
    Remove-Job $job
} else {
    Receive-Job $job
    Remove-Job $job
}
