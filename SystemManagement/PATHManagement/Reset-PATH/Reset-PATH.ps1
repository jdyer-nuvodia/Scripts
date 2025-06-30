# =============================================================================
# Script: Reset-PATH.ps1
# Created: 2025-01-09 15:30:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-05-08 17:25:00 UTC
# Updated By: jdyer-nuvodia
# Version: 2.1.0
# Additional Info: Changed parameters to use -Scope instead of separate -Machine and -User switches
# =============================================================================

<#
.SYNOPSIS
    Resets the PATH environment variable to a predefined list of directories.
.DESCRIPTION
    This script modifies the PATH environment variable for either the current user or machine-wide.
    It creates a backup of the existing PATH and sets a new predefined list of directories.
    When using the Machine scope, the script requires administrative privileges.

    Key actions:
     - Verifies administrative privileges when needed
     - Creates backup of current PATH
     - Sets new PATH with predefined directories for either User or Machine scope

    Dependencies:
     - Windows Operating System
     - Administrative privileges (for Machine PATH only)

    Security considerations:
     - Modifies environment variables at specified scope
     - Machine scope requires elevation to run
     - Creates backup file in script directory
     - WhatIf parameter allows previewing changes without applying them

    Performance impact:
     - Minimal system impact
     - One-time environment variable modification
     - No ongoing resource usage
.PARAMETER Scope
    Specifies the scope for the PATH reset operation.
    Valid values are "Machine" (system-wide) or "User" (current user).
    Default value is "User".
    The "Machine" scope requires administrative privileges.
.PARAMETER WhatIf
    Shows what would happen if the script runs without making any actual changes.
.EXAMPLE
    .\Reset-PATH.ps1
    Resets the User PATH to the predefined list of directories.
.EXAMPLE
    .\Reset-PATH.ps1 -Scope Machine
    Resets the Machine PATH to the predefined list of directories.
.EXAMPLE
    .\Reset-PATH.ps1 -Scope User -WhatIf
    Shows what the User PATH would be reset to without making any changes.
.NOTES
    Security Level: High
    Required Permissions: Administrative privileges (for Machine PATH only)
    Validation Requirements:
     - Verify PATH after modification
     - Ensure critical system paths are included
     - Test environment variable accessibility
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param (
    [Parameter(Mandatory=$false)]
    [ValidateSet("Machine", "User")]
    [string]$Scope = "User"
)

$pathType = $Scope

# Verify running as Administrator when modifying Machine PATH
if ($Scope -eq "Machine" -and -not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "Administrative privileges required to modify Machine PATH!"
    exit 1
}

# Set up logging
$scriptName = $MyInvocation.MyCommand.Name
$computerName = $env:COMPUTERNAME
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logPath = Join-Path -Path $PSScriptRoot -ChildPath "${computerName}_${scriptName}_${timestamp}.log"

function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"

    Add-Content -Path $logPath -Value $logMessage

    switch ($Level) {
        "INFO"    { Write-Host $Message -ForegroundColor White }
        "PROCESS" { Write-Host $Message -ForegroundColor Cyan }
        "SUCCESS" { Write-Host $Message -ForegroundColor Green }
        "WARNING" { Write-Host $Message -ForegroundColor Yellow }
        "ERROR"   { Write-Host $Message -ForegroundColor Red }
        "DEBUG"   { Write-Host $Message -ForegroundColor Magenta }
        "DETAIL"  { Write-Host $Message -ForegroundColor DarkGray }
        default   { Write-Host $Message }
    }
}

# Define the default PATH entries
$defaultMachinePaths = @(
    'C:\WINDOWS\system32',
    'C:\WINDOWS',
    'C:\WINDOWS\System32\Wbem',
    'C:\WINDOWS\System32\WindowsPowerShell\v1.0\',
    'C:\WINDOWS\System32\OpenSSH\',
    'C:\Program Files\PowerShell\7\',
    'C:\Program Files\dotnet\',
    'C:\Program Files\Git\cmd',
    'C:\Program Files (x86)\Windows Kits\10\Windows Performance Toolkit\'
)

$defaultUserPaths = @(
    'C:\Users\{0}\AppData\Local\Microsoft\WindowsApps' -f $env:USERNAME,
    'C:\Users\{0}\AppData\Local\GitHubDesktop\bin' -f $env:USERNAME
)

# Select appropriate paths based on PATH type
$newPathEntries = if ($Scope -eq "Machine") {
    $defaultMachinePaths
} else {
    $defaultUserPaths
}

try {
    Write-Log "Starting PATH reset for $pathType scope" "PROCESS"

    # Backup current PATH
    $currentPath = [Environment]::GetEnvironmentVariable('PATH', $pathType)
    $backupPath = Join-Path -Path $PSScriptRoot -ChildPath "${pathType}_PATH_Backup_${timestamp}.txt"

    Write-Log "Current $pathType PATH: $currentPath" "DETAIL"

    if ($PSCmdlet.ShouldProcess("$pathType PATH", "Reset to default values")) {
        $currentPath | Out-File -FilePath $backupPath -Encoding UTF8
        Write-Log "Backup of previous $pathType PATH saved to: $backupPath" "PROCESS"

        # Join the new paths with semicolon
        $newPath = $newPathEntries -join ';'

        # Set the new PATH
        [Environment]::SetEnvironmentVariable('PATH', $newPath, $pathType)

        Write-Log "$pathType PATH has been successfully updated" "SUCCESS"
        Write-Log "Please restart your terminal/applications for the changes to take effect" "PROCESS"
    } else {
        Write-Log "WhatIf: Would reset $pathType PATH to: $($newPathEntries -join ';')" "DEBUG"
    }
}
catch {
    Write-Log "Failed to update $pathType PATH: $_" "ERROR"
    exit 1
}

# Run PSScriptAnalyzer validation
if (Get-Command -Name Invoke-ScriptAnalyzer -ErrorAction SilentlyContinue) {
    Write-Log "Running PSScriptAnalyzer..." "PROCESS"
    $scriptAnalyzerResults = Invoke-ScriptAnalyzer -Path $MyInvocation.MyCommand.Path

    if ($scriptAnalyzerResults) {
        Write-Log "PSScriptAnalyzer found issues:" "WARNING"
        foreach ($result in $scriptAnalyzerResults) {
            Write-Log ("Line {0}: {1} - {2}" -f $result.Line, $result.RuleName, $result.Message) "DETAIL"
        }
    } else {
        Write-Log "PSScriptAnalyzer found no issues" "SUCCESS"
    }
} else {
    Write-Log "PSScriptAnalyzer not available. Install with: Install-Module -Name PSScriptAnalyzer -Force" "DETAIL"
}
