# =============================================================================
# Script: Reset-PATH.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 2.1.1
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
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

[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $false)]
    [ValidateSet("Machine", "User")]
    [string]$Scope = "User"
)

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

function Write-LogMessage {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"

    Add-Content -Path $logPath -Value $logMessage

    switch ($Level) {
        "INFO" { Write-ColorOutput -Message  -Color 'White' }
        "PROCESS" { Write-ColorOutput -Message  -Color 'Cyan' }
        "SUCCESS" { Write-ColorOutput -Message  -Color 'Green' }
        "WARNING" { Write-ColorOutput -Message  -Color 'Yellow' }
        "ERROR" { Write-ColorOutput -Message  -Color 'Red' }
        "DEBUG" { Write-ColorOutput -Message  -Color 'Magenta' }
        "DETAIL" { Write-ColorOutput -Message  -Color 'DarkGray' }
        default { Write-ColorOutput -Message $Message -Color "White" }
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
    'C:\Users\{ 0}\AppData\Local\Microsoft\WindowsApps' -f $env:USERNAME,
    'C:\Users\{ 0}\AppData\Local\GitHubDesktop\bin' -f $env:USERNAME
)

# Select appropriate paths based on PATH type
$newPathEntries = if ($Scope -eq "Machine") {
    $defaultMachinePaths
} else {
    $defaultUserPaths
}

try {
    Write-LogMessage "Starting PATH reset for $pathType scope" "PROCESS"

    # Backup current PATH
    $currentPath = [Environment]::GetEnvironmentVariable('PATH', $pathType)
    $backupPath = Join-Path -Path $PSScriptRoot -ChildPath "${ pathType}_PATH_Backup_${ timestamp}.txt"

    Write-LogMessage "Current $pathType PATH: $currentPath" "DETAIL"

    if ($PSCmdlet.ShouldProcess("$pathType PATH", "Reset to default values")) {
        $currentPath | Out-File -FilePath $backupPath -Encoding UTF8
        Write-LogMessage "Backup of previous $pathType PATH saved to: $backupPath" "PROCESS"

        # Join the new paths with semicolon
        $newPath = $newPathEntries -join ';'

        # Set the new PATH
        [Environment]::SetEnvironmentVariable('PATH', $newPath, $pathType)

        Write-LogMessage "$pathType PATH has been successfully updated" "SUCCESS"
        Write-LogMessage "Please restart your terminal/applications for the changes to take effect" "PROCESS"
    } else {
        Write-LogMessage "WhatIf: Would reset $pathType PATH to: $($newPathEntries -join ';')" "DEBUG"
    }
} catch {
    Write-LogMessage "Failed to update $pathType PATH: $_" "ERROR"
    exit 1
}

# Run PSScriptAnalyzer validation
if (Get-Command -Name Invoke-ScriptAnalyzer -ErrorAction SilentlyContinue) {
    Write-LogMessage "Running PSScriptAnalyzer..." "PROCESS"
    $scriptAnalyzerResults = Invoke-ScriptAnalyzer -Path $MyInvocation.MyCommand.Path

    if ($scriptAnalyzerResults) {
        Write-LogMessage "PSScriptAnalyzer found issues:" "WARNING"
        foreach ($result in $scriptAnalyzerResults) {
            Write-LogMessage ("Line { 0}: { 1} - { 2}" -f $result.Line, $result.RuleName, $result.Message) "DETAIL"
        }
    } else {
        Write-LogMessage "PSScriptAnalyzer found no issues" "SUCCESS"
    }
} else {
    Write-LogMessage "PSScriptAnalyzer not available. Install with: Install-Module -Name PSScriptAnalyzer -Force" "DETAIL"
}
