# Script: Set-MachinePath.ps1
# Version: 1.0
# Description: Sets the machine PATH to a predefined set of directories
# Author: jdyer-nuvodia
# Last Modified: 2025-02-05 22:57:12
#
# .SYNOPSIS
#   Sets the machine PATH environment variable to a specific set of directories
#
# .DESCRIPTION
#   This script sets the machine (system) PATH environment variable to a predefined
#   set of directories. It includes safety checks, creates a backup of the current
#   PATH, and verifies the existence of directories before setting them.
#
# .PARAMETER BackupOnly
#   If specified, only creates a backup of the current PATH without making changes
#
# .PARAMETER RestoreFromBackup
#   If specified, restores the PATH from the most recent backup file
#
# .EXAMPLE
#   # Set the machine PATH
#   .\Set-MachinePath.ps1
#
# .EXAMPLE
#   # Create backup only
#   .\Set-MachinePath.ps1 -BackupOnly
#
# .EXAMPLE
#   # Restore from backup
#   .\Set-MachinePath.ps1 -RestoreFromBackup
#
# .NOTES
#   - Requires administrative privileges
#   - Creates backup before making changes
#   - Verifies directory existence
#   - Maintains proper PATH format
#

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(Mandatory=$false)]
    [switch]$BackupOnly,
    
    [Parameter(Mandatory=$false)]
    [switch]$RestoreFromBackup
)

# Define the desired PATH directories
$pathDirs = @(
    "C:\Program Files\PowerShell\7",
    "C:\AzCopy",
    "C:\Users\Nuvodialocal\AppData\Local\Microsoft\WindowsApps",
    "C:\Program Files (x86)\Microsoft SDKs\Azure\CLI2\wbin",
    "C:\WINDOWS\system32",
    "C:\WINDOWS",
    "C:\WINDOWS\System32\Wbem",
    "C:\WINDOWS\System32\WindowsPowerShell\v1.0\",
    "C:\WINDOWS\System32\OpenSSH\",
    "C:\Program Files\dotnet\",
    "C:\Program Files (x86)\Windows Kits\10\Windows Performance Toolkit\",
    "C:\Program Files\PowerShell\7\",
    "C:\Users\jdyer\AppData\Local\Microsoft\WindowsApps",
    "C:\Users\jdyer\AppData\Local\GitHubDesktop\bin"
)

function Test-AdminPrivileges {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Backup-Path {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupPath = Join-Path $env:TEMP "PATH_backup_$timestamp.txt"
    
    try {
        [Environment]::GetEnvironmentVariable("PATH", "Machine") | Out-File -FilePath $backupPath -Encoding UTF8
        Write-Host "PATH backup created: $backupPath" -ForegroundColor Green
        return $backupPath
    }
    catch {
        throw "Failed to create PATH backup: $_"
    }
}

function Get-LatestBackup {
    $backups = Get-ChildItem -Path $env:TEMP -Filter "PATH_backup_*.txt" | Sort-Object LastWriteTime -Descending
    if ($backups.Count -eq 0) {
        throw "No backup files found in $env:TEMP"
    }
    return $backups[0].FullName
}

function Restore-PathFromBackup {
    param([string]$backupFile)
    
    try {
        $oldPath = Get-Content -Path $backupFile -Raw
        if ($PSCmdlet.ShouldProcess("Machine PATH", "Restore from backup")) {
            [Environment]::SetEnvironmentVariable("PATH", $oldPath, "Machine")
            Write-Host "PATH restored from backup successfully" -ForegroundColor Green
        }
    }
    catch {
        throw "Failed to restore PATH from backup: $_"
    }
}

function Test-PathDirs {
    $nonExistentDirs = @()
    foreach ($dir in $pathDirs) {
        if (!(Test-Path $dir)) {
            $nonExistentDirs += $dir
        }
    }
    return $nonExistentDirs
}

try {
    # Check for admin privileges
    if (!(Test-AdminPrivileges)) {
        throw "This script requires administrative privileges."
    }

    # Handle backup only mode
    if ($BackupOnly) {
        Backup-Path
        exit 0
    }

    # Handle restore from backup mode
    if ($RestoreFromBackup) {
        $latestBackup = Get-LatestBackup
        Restore-PathFromBackup -backupFile $latestBackup
        exit 0
    }

    # Create backup before making changes
    $backupFile = Backup-Path

    # Check for non-existent directories
    $nonExistentDirs = Test-PathDirs
    if ($nonExistentDirs.Count -gt 0) {
        Write-Warning "The following directories do not exist:"
        $nonExistentDirs | ForEach-Object { Write-Warning "  $_" }
        $choice = Read-Host "Do you want to continue anyway? (y/n)"
        if ($choice -ne 'y') {
            throw "Operation cancelled by user."
        }
    }

    # Set the new PATH
    $newPath = $pathDirs -join ";"
    
    if ($PSCmdlet.ShouldProcess("Machine PATH", "Set new value")) {
        [Environment]::SetEnvironmentVariable("PATH", $newPath, "Machine")
        Write-Host "Machine PATH updated successfully" -ForegroundColor Green
        Write-Host "Backup file created at: $backupFile" -ForegroundColor Cyan
        Write-Host "`nNew PATH value:" -ForegroundColor Yellow
        $pathDirs | ForEach-Object { Write-Host "  $_" }
    }
}
catch {
    Write-Error "Error occurred: $_"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    
    if ($backupFile) {
        Write-Host "`nTo restore the previous PATH, run:" -ForegroundColor Yellow
        Write-Host ".\Set-MachinePath.ps1 -RestoreFromBackup" -ForegroundColor Yellow
    }
    
    exit 1
}