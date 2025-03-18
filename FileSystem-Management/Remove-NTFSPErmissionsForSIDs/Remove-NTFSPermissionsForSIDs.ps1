# =============================================================================
# Script: Remove-NTFSPermissionsForSIDs.ps1
# Created: 2025-03-18 17:20:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-18 18:38:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1.2
# Additional Info: Fixed parameter naming inconsistencies, added missing functions, corrected Write-Log parameters
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

# Function to handle logging with colors
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = 'INFO',
        [string]$Color = 'White',
        [string]$Category = '',
        [switch]$NoConsole
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $categoryText = if ($Category) { "[$Category] " } else { "" }
    $logMessage = "[$timestamp] [$Level] $categoryText$Message"
    
    # Write to console with color if not suppressed
    if (-not $NoConsole) {
        Write-Host $logMessage -ForegroundColor $Color
    }
    
    # Write to log file
    Add-Content -Path $script:DebugLogFile -Value $logMessage
}

# Function to handle progress reporting
function Write-ProgressStatus {
    param (
        [string]$Activity,
        [string]$Status,
        [int]$Current,
        [int]$Total
    )
    
    if ($Total -gt 0) {
        $percentComplete = [math]::Min([math]::Round(($Current / $Total) * 100), 100)
        Write-Progress -Activity $Activity -Status $Status -PercentComplete $percentComplete
    } else {
        Write-Progress -Activity $Activity -Status $Status
    }
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

# Function to convert SID to name
function Convert-SidToName {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Sid
    )
    
    # Check if we already have this SID cached
    if ($script:SidCache.ContainsKey($Sid)) {
        return $script:SidCache[$Sid]
    }
    
    try {
        $sidObj = New-Object System.Security.Principal.SecurityIdentifier($Sid)
        $name = $sidObj.Translate([System.Security.Principal.NTAccount]).Value
        $script:SidCache[$Sid] = $name
        return $name
    }
    catch {
        if ($EnableSIDDiagnostics) {
            Write-Log "Failed to resolve SID ${Sid}: $_" -Level 'DEBUG' -Color "Magenta"
        }
        $script:SidCache[$Sid] = $Sid
        $script:FailedSids.Add($Sid) | Out-Null
        return $Sid
    }
}

# Function to get folder permissions
function Get-FolderPermissions {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Folder
    )
    
    try {
        $acl = Get-Acl -Path $Folder
        $permissions = @()
        
        foreach ($ace in $acl.Access) {
            $sid = if ($ace.IdentityReference -match '^S-1-') {
                $ace.IdentityReference.Value
            } else {
                try {
                    $ntAccount = [System.Security.Principal.NTAccount]$ace.IdentityReference
                    $sid = $ntAccount.Translate([System.Security.Principal.SecurityIdentifier]).Value
                } catch {
                    "Unknown"
                }
            }
            
            $permissions += [PSCustomObject]@{
                Path = $Folder
                Identity = $ace.IdentityReference.Value
                SID = $sid
                Type = $ace.AccessControlType
                Rights = $ace.FileSystemRights
                Inheritance = $ace.InheritanceFlags
                Propagation = $ace.PropagationFlags
                IsInherited = $ace.IsInherited
            }
        }
        
        return $permissions
    }
    catch {
        Write-Log "Error getting permissions for ${Folder}: $_" -Level 'ERROR' -Color "Red"
        return $null
    }
}

# Function to recursively process folders
function Invoke-FolderRecursively {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [int]$CurrentDepth = 0
    )
    
    if ($script:cancellationTokenSource.Token.IsCancellationRequested) {
        Write-Log "Cancellation requested. Stopping folder recursion." -Level 'WARNING' -Color "Yellow"
        return
    }
    
    # Check timeout
    $currentDuration = (Get-Date) - $script:StartTime
    if ($currentDuration -gt $script:processingTimeout) {
        Write-Log "Processing timeout reached ($TimeoutMinutes minutes). Canceling further processing." -Level 'WARNING' -Color "Yellow"
        $script:cancellationTokenSource.Cancel()
        return
    }
    
    # Check depth limit
    if ($MaxDepth -gt 0 -and $CurrentDepth -gt $MaxDepth) {
        return
    }
    
    try {
        $script:ProcessedFolders++
        
        # Process current folder
        Invoke-FolderProcessing -StartPath $Path -CurrentCount $script:ProcessedFolders -TotalCount $script:TotalFolders
        
        # Get subfolders
        $subFolders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue
        
        # Process subfolders
        foreach ($folder in $subFolders) {
            Invoke-FolderRecursively -Path $folder.FullName -CurrentDepth ($CurrentDepth + 1)
        }
    }
    catch {
        Write-Log "Error processing path $Path recursively: $_" -Level 'ERROR' -Color "Red"
    }
}

# Function to initialize well-known SIDs
function Initialize-WellKnownSIDs {
    $wellKnownAccounts = @(
        [PSCustomObject]@{ Name = "Administrator"; Domain = "BUILTIN"; SID = "S-1-5-32-544" },
        [PSCustomObject]@{ Name = "Domain Admins"; Domain = "DOMAIN"; SID = "S-1-5-21-DOMAIN-512" },
        [PSCustomObject]@{ Name = "Enterprise Admins"; Domain = "DOMAIN"; SID = "S-1-5-21-DOMAIN-519" }
    )
    
    # Try to get the actual domain SID to replace placeholders
    try {
        $domain = (Get-CimInstance Win32_ComputerSystem).Domain
        $domainSid = (New-Object System.Security.Principal.NTAccount("$domain\Domain Users")).Translate([System.Security.Principal.SecurityIdentifier]).Value
        $domainSidBase = $domainSid -replace "-513$", ""
        
        # Update domain SIDs with actual domain SID
        foreach ($account in $wellKnownAccounts) {
            if ($account.Domain -eq "DOMAIN") {
                $account.SID = $account.SID -replace "S-1-5-21-DOMAIN", $domainSidBase
            }
        }
    }
    catch {
        Write-Log "Could not determine domain SID. Using placeholder values." -Level 'WARNING' -Color "Yellow"
    }
    
    return $wellKnownAccounts
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

# New function to remove permissions for target SIDs - FIXED PARAMETER NAME
function Remove-TargetSIDPermissions {
    param (
        [Parameter(Mandatory=$true)]
        [string]$StartPath,
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
                Write-Log "Removed permission for $name from $StartPath" -Level 'SUCCESS' -Color "Green"
            }
        }
    }

    if ($modified) {
        try {
            # Backup ACL for potential rollback
            $backupAcl = Get-Acl -Path $StartPath
            
            Set-Acl -Path $StartPath -AclObject $Acl
            Write-Log "Successfully updated permissions on $StartPath (Removed $removedCount permissions)" -Level 'SUCCESS' -Color "Green"
        } catch {
            Write-Log "Failed to update permissions on ${StartPath}: $_" -Level 'ERROR' -Color "Red"
            
            # Attempt rollback
            try {
                Set-Acl -Path $StartPath -AclObject $backupAcl
                Write-Log "Rolled back permissions on $StartPath due to error" -Level 'WARNING' -Color "Yellow"
            } catch {
                Write-Log "Failed to rollback permissions on ${StartPath}: $_" -Level 'ERROR' -Color "Red"
            }
        }
    }

    return $modified
}

# Modified Invoke-FolderProcessing to include permission removal - FIXED PARAMETER NAME
function Invoke-FolderProcessing {
    param(
        [string]$StartPath,
        [int]$CurrentCount,
        [int]$TotalCount
    )

    $startTime = Get-Date
    try {
        Write-Log -Message "Processing folder: $StartPath" -Level 'VERBOSE' -Category 'FolderProcess' -NoConsole

        # Verify permissions before attempting changes
        try {
            $testFile = Join-Path $env:TEMP "ACLTest_$(Get-Random).tmp"
            New-Item -Path $testFile -ItemType File | Out-Null
            Remove-Item -Path $testFile -Force | Out-Null
        } catch {
            Write-Log "Insufficient permissions to make changes to the file system. Run as administrator." -Level 'ERROR' -Color "Red"
            return
        }

        # Get current ACL
        $acl = Get-Acl -Path $StartPath
        
        # Check and remove target SID permissions - FIXED PARAMETER NAME
        $modified = Remove-TargetSIDPermissions -StartPath $StartPath -Acl $acl

        # Get updated permissions for logging
        if ($modified) {
            $acl = Get-Acl -Path $StartPath
        }

        # Continue with permission analysis
        $permissionData = Get-FolderPermissions -Folder $StartPath

        if ($permissionData) {
            # Process permissions as before            
        }
    }
    catch {
        Write-Log -Message "Error processing folder ${StartPath}: $($_.Exception.Message)" -Color "Red" -Level 'ERROR'
        Write-Log -Message "Stack trace: $($_.ScriptStackTrace)" -Level 'DEBUG' -Category 'Exception' -NoConsole
    }
    finally {
        Write-PerformanceMetric -Operation "Folder processing for $StartPath" -StartTime $startTime
        Write-ProgressStatus -Activity "Analyzing and Updating Folder Permissions" -Status "Processing: $StartPath" -Current $CurrentCount -Total $TotalCount
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
    if (-not $script:TranscriptStarted) {
        $script:TranscriptStarted = $true
        $script:DebugLogFile = Join-Path $PSScriptRoot "Remove-NTFSPermissionsForSIDs.log"
        Start-Transcript -Path $script:DebugLogFile -Append
    }

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
}
catch [System.Exception] {
    Write-Error "An error occurred: $_"
    Write-Error $_.ScriptStackTrace
}
finally {
    # Stop transcript
    if ($script:TranscriptStarted) {
        Stop-Transcript
    }
    Write-Log -Message "Script execution completed" -Level 'INFO' -Color "Green"    
}