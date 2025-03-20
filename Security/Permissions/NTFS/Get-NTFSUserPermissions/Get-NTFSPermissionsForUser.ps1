# =============================================================================
# Script: Get-NTFSPermissions.ps1
# Created: 2025-02-25 23:15:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-20 23:46:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2.0
# Additional Info: Added SID file input support, progress bar, and logging
# =============================================================================

<#
.SYNOPSIS
    Gets NTFS permissions for specified users or SIDs in a directory structure.
.DESCRIPTION
    This script recursively checks NTFS permissions for specified users or SIDs
    across all folders under a given root directory. It can accept either a direct
    user input or read from a SIDS.txt file. Results are displayed in the console
    and saved to a transcript file, with detailed scanning logged to a debug file.
.PARAMETER User
    The username to check permissions for. Must include domain name if on a domain (e.g., "DOMAIN\username")
.PARAMETER RootFolder
    The starting folder path to begin the recursive permission check
.PARAMETER SIDFile
    Optional path to a text file containing SIDs to check (one per line)
.EXAMPLE
    .\Get-NTFSPermissions.ps1 -User "DOMAIN\jsmith" -RootFolder "D:\"
    Checks permissions for user DOMAIN\jsmith starting from D:\ drive
.EXAMPLE
    .\Get-NTFSPermissions.ps1 -RootFolder "D:\" -SIDFile "C:\SIDS.txt"
    Checks permissions for all SIDs listed in SIDS.txt starting from D:\ drive
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$User,
    
    [Parameter(Mandatory=$true)]
    [string]$RootFolder,
    
    [Parameter(Mandatory=$false)]
    [string]$SIDFile
)

# Initialize Logging
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logPath = Join-Path $PSScriptRoot "Logs"
$debugLogFile = Join-Path $logPath "Debug_$timestamp.log"
$transcriptFile = Join-Path $logPath "Transcript_$timestamp.txt"

# Create Logs directory if it does not exist
if (-not (Test-Path $logPath)) {
    New-Item -ItemType Directory -Path $logPath | Out-Null
}

# Start Transcript
Start-Transcript -Path $transcriptFile

# Function to write debug log
function Write-DebugLog {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $debugLogFile -Append
}

Write-Host "Initializing permission scan..." -ForegroundColor Cyan

# Get list of identities to check
$identities = @()
if ($User) {
    $identities += $User
}
if ($SIDFile -and (Test-Path $SIDFile)) {
    $identities += Get-Content $SIDFile
    Write-Host "Loaded $(($identities | Measure-Object).Count) identities from $SIDFile" -ForegroundColor Cyan
}
elseif ($SIDFile) {
    Write-Warning "SID file not found: $SIDFile"
}

if ($identities.Count -eq 0) {
    Write-Error "No identities specified. Please provide either a User or a SIDFile."
    Stop-Transcript
    return
}

# Get total folder count for progress bar
$totalFolders = (Get-ChildItem -Directory -Path $RootFolder -Recurse -Force | Measure-Object).Count
$currentFolder = 0

Get-ChildItem -Directory -Path $RootFolder -Recurse -Force | ForEach-Object {
    $folder = $_.FullName
    $currentFolder++
    
    # Update progress bar
    $percentComplete = ($currentFolder / $totalFolders) * 100
    Write-Progress -Activity "Scanning folders for permissions" -Status "Checking $folder" -PercentComplete $percentComplete
    
    # Log detailed scanning info to debug file
    Write-DebugLog "Scanning folder: $folder"
    
    try {
        $acl = Get-Acl $folder
        foreach ($identity in $identities) {
            $userAccess = $acl.Access | Where-Object { $_.IdentityReference -eq $identity }
            
            if ($userAccess) {
                Write-DebugLog "Found permissions for $identity in $folder"
                [PSCustomObject]@{
                    Folder = $folder
                    Identity = $identity
                    Permissions = $userAccess.FileSystemRights
                    IsInherited = $userAccess.IsInherited
                }
            }
        }
    }
    catch {
        Write-DebugLog "Error accessing $folder : $_"
        Write-Warning "Could not access $folder"
    }
} | Tee-Object -Variable results | Format-Table -AutoSize

Write-Progress -Activity "Scanning folders for permissions" -Completed

# Output summary
Write-Host "`nScan Complete" -ForegroundColor Green
Write-Host "Total folders scanned: $totalFolders" -ForegroundColor White
Write-Host "Results saved to: $transcriptFile" -ForegroundColor White
Write-Host "Debug log saved to: $debugLogFile" -ForegroundColor White

Stop-Transcript
