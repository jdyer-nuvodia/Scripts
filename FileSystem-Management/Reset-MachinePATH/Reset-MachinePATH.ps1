# =============================================================================
# Script: Reset-MachinePATH.ps1
# Created: 2025-01-09 15:30:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-25 23:26:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1
# Additional Info: Added security considerations and performance impact details
# =============================================================================

<#
.SYNOPSIS
    Resets the Machine PATH environment variable to a predefined list of directories.
.DESCRIPTION
    This script modifies the system-wide (Machine) PATH environment variable.
    It creates a backup of the existing PATH and sets a new predefined list of directories.
    The script requires administrative privileges to modify the Machine PATH.
    
    Key actions:
     - Verifies administrative privileges
     - Creates backup of current PATH
     - Sets new Machine PATH with predefined directories
    
    Dependencies:
     - Windows Operating System
     - Administrative privileges
     
    Security considerations:
     - Modifies system-wide environment variables
     - Requires elevation to run
     - Creates backup file in script directory
     
    Performance impact:
     - Minimal system impact
     - One-time environment variable modification
     - No ongoing resource usage
.EXAMPLE
    .\Reset-MachinePATH.ps1
    Resets the Machine PATH to the predefined list of directories.
.NOTES
    Security Level: High
    Required Permissions: Administrative privileges
    Validation Requirements: 
     - Verify PATH after modification
     - Ensure critical system paths are included
     - Test environment variable accessibility
#>

# Verify running as Administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "This script must be run as Administrator!"
    exit 1
}

# Define the new PATH entries
$newPathEntries = @(
    'C:\Program Files\PowerShell\7',
    'C:\AzCopy',
    'C:\Users\Nuvodialocal\AppData\Local\Microsoft\WindowsApps',
    'C:\Program Files (x86)\Microsoft SDKs\Azure\CLI2\wbin',
    'C:\WINDOWS\system32',
    'C:\WINDOWS',
    'C:\WINDOWS\System32\Wbem',
    'C:\WINDOWS\System32\WindowsPowerShell\v1.0\',
    'C:\WINDOWS\System32\OpenSSH\',
    'C:\Program Files\dotnet\',
    'C:\Program Files (x86)\Windows Kits\10\Windows Performance Toolkit\',
    'C:\Program Files\PowerShell\7\',
    'C:\Users\jdyer\AppData\Local\Microsoft\WindowsApps',
    'C:\Users\jdyer\AppData\Local\GitHubDesktop\bin'
)

try {
    # Backup current PATH
    $currentPath = [Environment]::GetEnvironmentVariable('PATH', 'Machine')
    $backupPath = "PATH_Backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    $currentPath | Out-File -FilePath $backupPath -Encoding UTF8

    # Join the new paths with semicolon
    $newPath = $newPathEntries -join ';'

    # Set the new Machine PATH
    [Environment]::SetEnvironmentVariable('PATH', $newPath, 'Machine')

    Write-Host "Machine PATH has been successfully updated."
    Write-Host "Backup of previous PATH saved to: $backupPath"
    Write-Host "Please restart your terminal/applications for the changes to take effect."
}
catch {
    Write-Error "Failed to update PATH: $_"
    exit 1
}
