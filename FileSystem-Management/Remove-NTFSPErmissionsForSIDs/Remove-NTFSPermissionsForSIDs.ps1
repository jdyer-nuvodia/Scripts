# =============================================================================
# Script: Remove-NTFSPermissionsForSIDs.ps1
# Created: 2025-03-18 17:20:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-18 17:25:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1.1
# Additional Info: Added usage of adminAccounts variable for admin SID validation
# =============================================================================

<#
.SYNOPSIS
Removes NTFS permissions for specified SIDs from folders after user confirmation.

.DESCRIPTION
Analyzes folder permissions and removes access for SIDs listed in SIDs.txt file.
Requires user confirmation before removing permissions for each unique SID.
Consolidates output into two log files:
- Main log for permission removal details
- Debug log for troubleshooting information

.PARAMETER StartPath
The folder path to analyze. Must be a valid NTFS path.

.PARAMETER MaxThreads
Maximum number of concurrent threads to use for processing.

.PARAMETER MaxDepth
Maximum folder depth to analyze. 0 means no limit.

.PARAMETER SkipUniquenessCounting
Skips the counting of unique permissions to improve performance.

.PARAMETER SkipADResolution
Skips Active Directory resolution for SIDs.

.PARAMETER EnableSIDDiagnostics
Enables detailed diagnostics for SID resolution issues.

.PARAMETER TimeoutMinutes
Maximum time in minutes to allow the script to run.

.PARAMETER EnableProgressBar
Enables progress bar display during processing.

.EXAMPLE
.\Remove-NTFSPermissionsForSIDs.ps1 -StartPath "C:\Temp"
Analyzes permissions on C:\Temp, prompts for SID removal confirmation, and removes confirmed permissions
#>

using namespace System.Security.AccessControl
using namespace System.IO
using namespace System.Security.Principal

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$StartPath,

    [Parameter(Mandatory = $false)]
    [int]$MaxThreads = 10,

    [Parameter(Mandatory = $false)]
    [int]$MaxDepth = 0,

    [Parameter(Mandatory = $false)]
    [switch]$SkipUniquenessCounting,

    [Parameter(Mandatory = $false)]
    [switch]$SkipADResolution,

    [Parameter(Mandatory = $false)]
    [bool]$EnableSIDDiagnostics = $true,

    [Parameter(Mandatory = $false)]
    [int]$TimeoutMinutes = 120,

    [Parameter(Mandatory = $false)]
    [switch]$EnableProgressBar
)

# Enable strict mode and error handling
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Script-level variables
$script:TranscriptStarted = $false
$script:SidCache = @{}
$script:FailedSids = [System.Collections.Generic.HashSet[string]]::new()
$script:SuppressedSids = [System.Collections.Generic.List[string]]::new()
$script:TotalFolders = 0
$script:ProcessedFolders = 0
$script:StartTime = Get-Date
$script:EndTime = $null
$script:ElapsedTime = $null
$script:FolderPermissions = @{}
$script:UniquePermissions = @{}
$script:PermissionGroups = @{}
$script:InheritanceStatus = @{}
$script:ParentPermissions = @{}
$script:SidTranslationAttempts = @{}
$script:WellKnownSIDs = @{}
$script:ADResolutionErrors = @{}
$script:MaxRetries = 3
$script:RetryDelay = 2
$script:ApprovedSIDRemovals = @{}
$script:TargetSIDs = @()

# Script-level cancellation token
$script:cancellationTokenSource = New-Object System.Threading.CancellationTokenSource
$script:processingTimeout = New-TimeSpan -Minutes $TimeoutMinutes

# =============================================================================
# Core Functions - Import existing functions from Get-NTFSFolderPermissions.ps1
# =============================================================================

# Function to handle progress reporting
function Write-Progress {
    param (
        [string]$Message,
        [int]$PercentComplete
    )
    Write-Progress -Activity "Processing Permissions" -Status $Message -PercentComplete $PercentComplete
}

# Function to handle performance metrics
function Write-PerformanceMetric {
    param (
        [string]$Operation,
        [datetime]$StartTime,
        [int]$ProcessedItems = 0
    )
    
    $endTime = Get-Date
    $duration = ($endTime - $StartTime).TotalMilliseconds
    $rate = if ($ProcessedItems -gt 0) { [math]::Round(($ProcessedItems * 1000) / $duration, 2) } else { 0 }
    
    Write-Log -Message "Performance: $Operation completed in $([math]::Round($duration, 2))ms ($rate items/sec)" -Level 'METRIC'
}

# Function to handle logging with colors
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = 'INFO',
        [string]$Color = 'White'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Write to console with color
    Write-Host $logMessage -ForegroundColor $Color
    
    # Write to log file
    Add-Content -Path $script:DebugLogFile -Value $logMessage
}

# New function to load target SIDs from file
function Import-TargetSIDs {
    $sidsPath = Join-Path $PSScriptRoot "SIDs.txt"
    if (-not (Test-Path $sidsPath)) {
        throw "SIDs.txt not found in script directory: $PSScriptRoot"
    }

    $script:TargetSIDs = Get-Content $sidsPath | Where-Object { $_ -match '^S-1-' }
    if ($script:TargetSIDs.Count -eq 0) {
        throw "No valid SIDs found in SIDs.txt"
    }

    Write-Log "Loaded $($script:TargetSIDs.Count) SIDs from SIDs.txt" -Level 'INFO' -Color "Cyan"
}

# New function to confirm SID removal
function Confirm-SIDRemoval {
    param (
        [Parameter(Mandatory=$true)]
        [string]$SID,
        [string]$Name
    )

    if ($script:ApprovedSIDRemovals.ContainsKey($SID)) {
        return $script:ApprovedSIDRemovals[$SID]
    }

    $displayName = if ($Name -and $Name -ne $SID) { "$Name ($SID)" } else { $SID }
    $confirmation = Read-Host -Prompt "Remove permissions for $displayName? (Y/N)"
    $approved = $confirmation -eq 'Y'
    $script:ApprovedSIDRemovals[$SID] = $approved

    Write-Log "User $('approved' * $approved + 'denied' * -not $approved) removal of permissions for $displayName" -Level $(if ($approved) { 'INFO' } else { 'WARNING' })
    return $approved
}

# New function to remove permissions for target SIDs
function Remove-TargetSIDPermissions {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [System.Security.AccessControl.FileSystemSecurity]$Acl
    )

    $modified = $false
    $removedCount = 0

    foreach ($ace in @($Acl.Access)) {
        $sid = if ($ace.IdentityReference -match '^S-1-') {
            $ace.IdentityReference.Value
        } else {
            try {
                $ntAccount = [System.Security.Principal.NTAccount]$ace.IdentityReference
                $sid = $ntAccount.Translate([System.Security.Principal.SecurityIdentifier]).Value
            } catch {
                Write-Log "Could not translate $($ace.IdentityReference) to SID" -Level 'WARNING'
                continue
            }
        }

        if ($script:TargetSIDs -contains $sid) {
            $name = Convert-SidToName -Sid $sid
            if (Confirm-SIDRemoval -SID $sid -Name $name) {
                $Acl.RemoveAccessRule($ace) | Out-Null
                $modified = $true
                $removedCount++
                Write-Log "Removed permission for $name from $Path" -Level 'SUCCESS' -Color "Green"
            }
        }
    }

    if ($modified) {
        try {
            Set-Acl -Path $Path -AclObject $Acl
            Write-Log "Successfully updated permissions on $Path (Removed $removedCount permissions)" -Level 'SUCCESS' -Color "Green"
        } catch {
            Write-Log "Failed to update permissions on ${Path}: $_" -Level 'ERROR' -Color "Red"
        }
    }

    return $modified
}

# Modified Invoke-FolderProcessing to include permission removal
function Invoke-FolderProcessing {
    param(
        [string]$Path,
        [int]$CurrentCount,
        [int]$TotalCount
    )

    $startTime = Get-Date
    try {
        Write-Log -Message "Processing folder: $Path" -Level 'VERBOSE' -Category 'FolderProcess' -NoConsole

        # Get current ACL
        $acl = Get-Acl -Path $Path
        
        # Check and remove target SID permissions
        $modified = Remove-TargetSIDPermissions -Path $Path -Acl $acl

        # Get updated permissions for logging
        if ($modified) {
            $acl = Get-Acl -Path $Path
        }

        # Continue with permission analysis
        $permissionData = Get-FolderPermissions -Folder $Path

        if ($permissionData) {
            # Process permissions as before
            // ...existing code...
        }
    }
    catch {
        Write-Log -Message "Error processing folder ${Path}: $($_.Exception.Message)" -Color "Red" -Level 'ERROR'
        Write-Log -Message "Stack trace: $($_.ScriptStackTrace)" -Level 'DEBUG' -Category 'Exception' -NoConsole
    }
    finally {
        Write-PerformanceMetric -Operation "Folder processing for $Path" -StartTime $startTime
        Write-ProgressStatus -Activity "Analyzing and Updating Folder Permissions" -Status "Processing: $Path" -Current $CurrentCount -Total $TotalCount
    }
}

# Modified function to use adminAccounts
function Test-AdministratorSID {
    param([string]$SID)
    
    if ([string]::IsNullOrEmpty($SID)) { return $false }
    
    # Check against known admin accounts first
    if ($script:adminAccounts.SID -contains $SID) {
        return $true
    }
    
    # Fallback to pattern matching
    return $SID -match '^S-1-5-21-\d+-\d+-\d+-500$'
}

# Main script execution
try {
    # Start transcript
    // ...existing code...

    # Load target SIDs
    Import-TargetSIDs

    # Initialize SIDs and get Administrator accounts
    $script:adminAccounts = Initialize-WellKnownSIDs

    # Validate admin accounts and add to well-known SIDs
    foreach ($admin in $script:adminAccounts) {
        if (-not $script:WellKnownSIDs.ContainsValue($admin.SID)) {
            $script:WellKnownSIDs["Administrator_$($admin.Domain)"] = $admin.SID
        }
    }

    Write-Log -Message "Starting folder permission analysis and removal for $StartPath" -Color "Cyan" -Level 'INFO'

    # Process folders with timeout tracking
    Invoke-FolderRecursively -Path $StartPath

    if ($script:cancellationTokenSource.Token.IsCancellationRequested) {
        Write-Log "`nProcessing terminated before completion" -Level 'WARNING' -Color "Yellow"
    }

    # Calculate elapsed time
    $script:EndTime = Get-Date
    $script:ElapsedTime = $script:EndTime - $script:StartTime

    # Display summary
    Write-Log -Message "`nAnalysis Complete" -Color "Green" -Level 'SUCCESS'
    Write-Log -Message "Total folders processed: $($script:ProcessedFolders)" -Color "Cyan" -Level 'INFO'
    Write-Log -Message "Permissions removed for: $(@($script:ApprovedSIDRemovals.Keys | Where-Object { $script:ApprovedSIDRemovals[$_] }).Count) SIDs" -Color "Cyan" -Level 'INFO'
    Write-Log -Message "Elapsed time: $($script:ElapsedTime.ToString())" -Color "Cyan" -Level 'INFO'

    // ...existing code...
}
catch [System.Exception] {
    Write-Error "An error occurred: $_"
    Write-Error $_.ScriptStackTrace
}
finally {
    // ...existing code...
}