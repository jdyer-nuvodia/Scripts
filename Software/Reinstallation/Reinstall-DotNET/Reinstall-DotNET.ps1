# =============================================================================
# Script: Reinstall-DotNET.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1.1
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
# =============================================================================

<#
.SYNOPSIS
    Updates all installed versions of .NET Framework.
.DESCRIPTION
    This script performs the following actions:
    - Detects installed .NET Framework versions
    - Downloads required updates from Microsoft
    - Installs updates silently
    - Verifies successful installation
.PARAMETER Verbose
    When specified, provides detailed information about script execution.
.EXAMPLE
    .\Reinstall-DotNET.ps1
    Updates all installed .NET Framework versions silently
.NOTES
    Security Level: Medium
    Required Permissions: Administrator
    Validation Requirements: Verify .NET versions after update
#>

[CmdletBinding()]
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

# Ensure running as administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Verbose "Restarting script with elevated privileges..."
    Start-Process PowerShell -ArgumentList "-NoProfile -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

function Get-InstalledDotNetVersion {
    Write-Verbose "Detecting installed .NET Framework versions..."
    $versions = @()

    # Check .NET Framework 4.x
    $net4 = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction SilentlyContinue
    if ($net4) {
        $version = $net4.GetValue("Version")
        $release = $net4.GetValue("Release")
        if ($version) {
            $versions += @{
                Version = $version
                Release = $release
                Family  = "4"
            }
        }
    }

    # Check .NET Framework 3.5
    $net35 = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5' -ErrorAction SilentlyContinue
    if ($net35) {
        $version = $net35.GetValue("Version")
        if ($version) {
            $versions += @{
                Version = $version
                Family  = "3.5"
            }
        }
    }

    return $versions
}

function Update-DotNetFramework {
    [CmdletBinding(SupportsShouldProcess)]
    param()

    $tempPath = Join-Path $env:TEMP "DotNetUpdates"
    New-Item -ItemType Directory -Force -Path $tempPath | Out-Null

    try {
        # Update .NET Framework 4.x
        Write-ColorOutput -Message "Updating .NET Framework 4.x..." -Color "Cyan"
        $url = "https://go.microsoft.com/fwlink/?LinkId = 2085155"
        $installer = Join-Path $tempPath "ndp48-x86-x64-allos-enu.exe"

        Write-Verbose "Downloading .NET Framework 4.8 update..."
        Invoke-WebRequest -Uri $url -OutFile $installer

        Write-Verbose "Installing .NET Framework 4.8 update..."
        if ($PSCmdlet.ShouldProcess("Install .NET Framework 4.8", "Installing .NET Framework 4.8 update")) {
            $process = Start-Process -FilePath $installer -ArgumentList "/quiet /norestart" -Wait -PassThru

            if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
                Write-ColorOutput -Message ".NET Framework 4.x update completed successfully" -Color "Green"
            } else {
                Write-Warning "Failed to update .NET Framework 4.x. Exit code: $($process.ExitCode)"
            }
        }

        # Enable .NET Framework 3.5 if needed
        Write-ColorOutput -Message "Checking .NET Framework 3.5..." -Color "Cyan"
        $state = (Get-WindowsOptionalFeature -Online -FeatureName "NetFx3").State
        if ($state -ne "Enabled") {
            Write-Verbose "Enabling .NET Framework 3.5..."
            if ($PSCmdlet.ShouldProcess("Enable .NET Framework 3.5", "Enabling .NET Framework 3.5 feature")) {
                Enable-WindowsOptionalFeature -Online -FeatureName "NetFx3" -NoRestart
            }
        }

        # Run Windows Update to get latest .NET updates
        Write-ColorOutput -Message "Checking for additional .NET updates through Windows Update..." -Color "Cyan"
        if ($PSCmdlet.ShouldProcess("Install PSWindowsUpdate module and get .NET updates", "Installing Windows Updates for .NET Framework")) {
            Install-Module PSWindowsUpdate -Force -AllowClobber
            Get-WindowsUpdate -MicrosoftUpdate -Category "Microsoft .NET Framework" -AcceptAll -Install
        }
    } catch {
        Write-Error "Error during .NET Framework update: $_"
        Write-Verbose "Stack trace: $($_.ScriptStackTrace)"
    } finally {
        Write-Verbose "Cleaning up temporary files..."
        Remove-Item -Path $tempPath -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Test-DotNetUpdate {
    Write-Verbose "Verifying .NET Framework updates..."
    $versions = Get-InstalledDotNetVersion

    foreach ($version in $versions) {
        Write-ColorOutput -Message "Detected .NET Framework $($version.Family) - Version: $($version.Version)" -Color "White"
    }
}

# Main execution
Write-Verbose "=== Starting .NET Framework update process ==="
Write-ColorOutput -Message "Starting .NET Framework updates..." -Color "Cyan"
Update-DotNetFramework
Test-DotNetUpdate

Write-ColorOutput -Message "`nUpdate process completed. A system restart may be required." -Color "Yellow"
$restartPrompt = Read-Host "Would you like to restart now? (Y/N)"
if ($restartPrompt -eq 'Y' -or $restartPrompt -eq 'y') {
    Restart-Computer -Force
}
