# =============================================================================
# Script: Get-DriveInfo.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.3
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
# =============================================================================

<#
.SYNOPSIS
    Displays detailed drive space information for all available volumes.
.DESCRIPTION
    Retrieves and displays comprehensive drive information including:
    - Drive letter and label
    - File system type
    - Total, used, and free space
    - Health status
    Uses PowerShell's Get-Volume cmdlet and custom formatting.
.EXAMPLE
    .\Get-DriveInfo.ps1
    Displays information for all available drive volumes
#>

[CmdletBinding()]
param ()


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

# Initialize logging
$scriptPath = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Path)
$systemName = [System.Environment]::MachineName
$logFile = [System.IO.Path]::Combine($scriptPath, "Get-DriveInfo_${systemName}_$([DateTime]::UtcNow.ToString('yyyyMMdd_HHmmss')).log")
$script:logStream = [System.IO.StreamWriter]::new($logFile, $true, [System.Text.Encoding]::UTF8)

function Write-LogMessage {
    param(
        [string]$Message,
        [string]$Color = 'White',
        [switch]$NoConsole
    )

    $timestamp = [DateTime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ss UTC')
    $logMessage = "[$timestamp] $Message"

    # Write to log file
    $script:logStream.WriteLine($logMessage)
    $script:logStream.Flush()

    # Write to console if not suppressed
    if (-not $NoConsole) {
        Write-ColorOutput -Message $Message -Color $Color
    }
}

function Show-DriveInfo {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Volume
    )

    Write-ColorOutput -Message "`nDrive Volume Details:" -Color 'Green'
    Write-ColorOutput -Message "------------------------" -Color 'Green'
    Write-ColorOutput -Message "Drive Letter: $($Volume.DriveLetter)" -Color 'Cyan'
    Write-ColorOutput -Message "Drive Label: $($Volume.FileSystemLabel)" -Color 'Cyan'
    Write-ColorOutput -Message "File System: $($Volume.FileSystem)" -Color 'Cyan'
    Write-ColorOutput -Message "Drive Type: $($Volume.DriveType)" -Color 'Cyan'
    Write-ColorOutput -Message "Size: $([math]::Round($Volume.Size/1GB, 2)) GB" -Color 'Cyan'
    Write-ColorOutput -Message "Free Space: $([math]::Round($Volume.SizeRemaining/1GB, 2)) GB" -Color 'Cyan'
    Write-ColorOutput -Message "Used Space: $([math]::Round(($Volume.Size - $Volume.SizeRemaining)/1GB, 2)) GB" -Color 'Cyan'
    Write-ColorOutput -Message "Free Space %: $([math]::Round(($Volume.SizeRemaining/$Volume.Size) * 100, 2))%" -Color 'Cyan'
    Write-ColorOutput -Message "Health Status: $($Volume.HealthStatus)" -Color 'Cyan'
    Write-ColorOutput -Message "Operational Status: $($Volume.OperationalStatus)" -Color 'Cyan'
    Write-ColorOutput -Message "" -Color "White"
}

function Write-StatusMessage {
    param(
        [string]$Message,
        [string]$Color = 'White'
    )
    Write-LogMessage -Message $Message -Color $Color
}

# Main execution
try {
    Write-StatusMessage "Getting drive space information..." -Color Cyan
    $volumes = Get-Volume | Where-Object { $_.DriveLetter } | Sort-Object DriveLetter

    if ($volumes.Count -eq 0) {
        Write-LogMessage "No drives with letters found on the system." -Color Red
        exit 1
    }

    Write-StatusMessage "Found $($volumes.Count) drive volumes." -Color Green
    foreach ($volume in $volumes) {
        Write-StatusMessage "Drive space for $($volume.DriveLetter):" -Color Yellow
        Show-DriveInfo -Volume $volume
    }

    Write-LogMessage "Drive information collection completed successfully. See log file for details: $logFile" -Color Green
} catch {
    Write-LogMessage "Error during execution: $($_.Exception.Message)" -Color Red
    Write-LogMessage "Stack trace: $($_.ScriptStackTrace)" -Color Red
    exit 1
} finally {
    if ($null -ne $script:logStream) {
        $script:logStream.Close()
        $script:logStream.Dispose()
    }
}
