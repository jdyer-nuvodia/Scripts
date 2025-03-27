# =============================================================================
# Script: Get-NTFSPermissionsForUser.ps1
# Created: 2025-02-07 21:21:53 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-27 21:36:00 UTC 
# Updated By: jdyer-nuvodia
# Version: 1.3.8
# Additional Info: Enhanced permission output formatting
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
.PARAMETER StartPath
    The starting folder path to begin the recursive permission check
.PARAMETER SIDFile
    Optional path to a text file containing SIDs to check (one per line)
.EXAMPLE
    .\Get-NTFSPermissionsForUser.ps1 -User "DOMAIN\jsmith" -StartPath "D:\"
    Checks permissions for user DOMAIN\jsmith starting from D:\ drive
.EXAMPLE
    .\Get-NTFSPermissionsForUser.ps1 -StartPath "D:\" -SIDFile "C:\SIDS.txt"
    Checks permissions for all SIDs listed in SIDS.txt starting from D:\ drive
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$User,
    
    [Parameter(Mandatory=$true)]
    [string]$StartPath,
    
    [Parameter(Mandatory=$false)]
    [string]$SIDFile
)

# Initialize Logging
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$systemName = $env:COMPUTERNAME

# Determine user string for filename
$userString = if ($User) {
    $userParts = $User.Split('\')
    $userParts[-1]  # Take the last part after splitting on backslash
} elseif ($SIDFile) {
    "multiple_users"
} else {
    "no_user"
}

# Construct log file names
$baseLogName = "NTFSPermissionsForUser_${systemName}_${userString}_${timestamp}"
$debugLogFile = Join-Path $PSScriptRoot "${baseLogName}_debug.log"
$transcriptFile = Join-Path $PSScriptRoot "${baseLogName}_transcript.log"

# Start Transcript
Start-Transcript -Path $transcriptFile

# Function to write debug log with error handling
function Write-DebugLog {
    param(
        [string]$Message,
        [switch]$IsError
    )
    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "$timestamp - $Message"
        if ($IsError) {
            $logEntry = "$logEntry [ERROR]"
            Write-Warning $Message
        }
        $logEntry | Out-File -FilePath $debugLogFile -Append -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to write to debug log: $_"
    }
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
$totalFolders = (Get-ChildItem -Directory -Path $StartPath -Recurse -Force | Measure-Object).Count

# Function to check and report folder permissions
function Test-FolderPermissions {
    param(
        [string]$FolderPath,
        [array]$IdentityList,
        [ref]$ProcessedCount
    )
    try {
        $acl = Get-Acl -Path $FolderPath -ErrorAction Stop
        foreach ($identity in $IdentityList) {
            $accessRules = $acl.Access | Where-Object { $_.IdentityReference -eq $identity }
            if ($accessRules) {
                foreach ($rule in $accessRules) {
                    Write-PermissionInfo -Identity $identity `
                                       -Path $FolderPath `
                                       -Access $rule.FileSystemRights `
                                       -AccessType $rule.AccessControlType
                }
            }
        }
        $ProcessedCount.Value++
        $percentComplete = [math]::Min(($ProcessedCount.Value / $totalFolders) * 100, 100)
        Write-Progress -Activity "Scanning folders for permissions" -Status "Processing: $FolderPath" -PercentComplete $percentComplete
    }
    catch {
        Write-DebugLog "Error checking permissions for $FolderPath : $_" -IsError
    }
}

# Enhanced folder scanning function with progress tracking
function Find-Folder {
    param(
        [string]$FolderPath,
        [array]$IdentityList,
        [ref]$ProcessedCount
    )
    try {
        Write-Host "Scanning: $FolderPath" -ForegroundColor DarkGray
        Write-DebugLog "Scanning folder: $FolderPath"
        Test-FolderPermissions -FolderPath $FolderPath -IdentityList $IdentityList -ProcessedCount $ProcessedCount

        Get-ChildItem -Path $FolderPath -Directory -ErrorAction Stop | ForEach-Object {
            try {
                Find-Folder -FolderPath $_.FullName -IdentityList $IdentityList -ProcessedCount $ProcessedCount
            }
            catch {
                Write-Host "Error scanning subfolder $($_.FullName)" -ForegroundColor Red
                Write-DebugLog "Error scanning subfolder $($_.FullName): $_" -IsError
            }
        }
    }
    catch {
        Write-Host "Error accessing folder $FolderPath" -ForegroundColor Red
        Write-DebugLog "Error accessing folder $FolderPath : $_" -IsError
    }
}

# Add permission reporting function
function Write-PermissionInfo {
    param(
        [string]$Identity,
        [string]$Path,
        [string]$Access,
        [string]$AccessType
    )
    
    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $permType = if ($AccessType -eq "Allow") { "[ALLOW]" } else { "[DENY]" }
        $message = "$timestamp - $permType - Identity: $Identity - Access: $Access - Path: $Path"
        
        # Use Green for Allow and Yellow for Deny permissions
        $color = if ($AccessType -eq "Allow") { "Green" } else { "Yellow" }
        Write-Host $message -ForegroundColor $color
        
        # Log with consistent formatting
        Write-DebugLog $message
        Add-Content -Path $transcriptFile -Value $message -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to log permission info: $_"
    }
}

# Initialize progress tracking
$processedFolders = [ref]0

# Start the scan with progress tracking
Write-Host "Beginning folder scan from $StartPath" -ForegroundColor Cyan
Find-Folder -FolderPath $StartPath -IdentityList $identities -ProcessedCount $processedFolders

# Output final statistics
Write-Progress -Activity "Scanning folders for permissions" -Completed
Write-Host "`nScan Complete" -ForegroundColor Green
Write-Host "Scan Statistics:" -ForegroundColor White
Write-Host "Folders processed: $($processedFolders.Value)/$totalFolders" -ForegroundColor Cyan
Write-Host "Debug log: $debugLogFile" -ForegroundColor DarkGray
Write-Host "Transcript: $transcriptFile" -ForegroundColor DarkGray

Stop-Transcript
