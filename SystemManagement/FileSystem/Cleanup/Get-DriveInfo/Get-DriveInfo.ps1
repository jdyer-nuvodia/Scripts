# =============================================================================
# Script: Get-DriveInfo.ps1
# Created: 2025-03-19 22:27:00 UTC
# Author: nunya-nunya
# Last Updated: 2025-03-19 22:27:00 UTC
# Updated By: nunya-nunya
# Version: 1.0.0
# Additional Info: Initial creation from Clear-SystemStorage.ps1
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

# Initialize logging
$scriptPath = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Path)
$systemName = [System.Environment]::MachineName
$logFile = [System.IO.Path]::Combine($scriptPath, "Get-DriveInfo_${systemName}_$([DateTime]::UtcNow.ToString('yyyyMMdd_HHmmss')).log")
$script:logStream = [System.IO.StreamWriter]::new($logFile, $true, [System.Text.Encoding]::UTF8)

function Write-Log {
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
        Write-Host $Message -ForegroundColor $Color
    }
}

function Show-DriveInfo {
    param (
        [Parameter(Mandatory=$true)]
        [object]$Volume
    )
    
    Write-Host "`nDrive Volume Details:" -ForegroundColor Green
    Write-Host "------------------------" -ForegroundColor Green
    Write-Host "Drive Letter: $($Volume.DriveLetter)" -ForegroundColor Cyan
    Write-Host "Drive Label: $($Volume.FileSystemLabel)" -ForegroundColor Cyan
    Write-Host "File System: $($Volume.FileSystem)" -ForegroundColor Cyan
    Write-Host "Drive Type: $($Volume.DriveType)" -ForegroundColor Cyan
    Write-Host "Size: $([math]::Round($Volume.Size/1GB, 2)) GB" -ForegroundColor Cyan
    Write-Host "Free Space: $([math]::Round($Volume.SizeRemaining/1GB, 2)) GB" -ForegroundColor Cyan
    Write-Host "Used Space: $([math]::Round(($Volume.Size - $Volume.SizeRemaining)/1GB, 2)) GB" -ForegroundColor Cyan
    Write-Host "Free Space %: $([math]::Round(($Volume.SizeRemaining/$Volume.Size) * 100, 2))%" -ForegroundColor Cyan
    Write-Host "Health Status: $($Volume.HealthStatus)" -ForegroundColor Cyan
    Write-Host "Operational Status: $($Volume.OperationalStatus)" -ForegroundColor Cyan
    Write-Host ""
}

function Write-StatusMessage {
    param(
        [string]$Message,
        [string]$Color = 'White'
    )
    Write-Log -Message $Message -Color $Color
}

# Main execution
try {
    Write-StatusMessage "Getting drive space information..." -Color Cyan
    $volumes = Get-Volume | Where-Object { $_.DriveLetter } | Sort-Object DriveLetter
    
    if ($volumes.Count -eq 0) {
        Write-Log "No drives with letters found on the system." -Color Red
        exit 1
    }
    
    Write-StatusMessage "Found $($volumes.Count) drive volumes." -Color Green
    foreach ($volume in $volumes) {
        Write-StatusMessage "Drive space for $($volume.DriveLetter):" -Color Yellow
        Show-DriveInfo -Volume $volume
    }
    
    Write-Log "Drive information collection completed successfully. See log file for details: $logFile" -Color Green
}
catch {
    Write-Log "Error during execution: $($_.Exception.Message)" -Color Red
    Write-Log "Stack trace: $($_.ScriptStackTrace)" -Color Red
    exit 1
}
finally {
    if ($null -ne $script:logStream) {
        $script:logStream.Close()
        $script:logStream.Dispose()
    }
}
