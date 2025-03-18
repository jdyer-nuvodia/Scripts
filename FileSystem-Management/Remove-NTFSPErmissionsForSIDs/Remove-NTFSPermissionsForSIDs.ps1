# =============================================================================
# Script: Remove-NTFSPermissionsForSIDs.ps1
# Created: 2025-03-18 17:20:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-18 23:22:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1.13
# Additional Info: Fixed $displayName variable reference and Write-Progress issues
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

# Initialize the display name variable
$script:displayName = $null  

# Enable strict mode and error handling
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Script-level variables
$script:TranscriptStarted = $false
$script:ComputerName = $env:COMPUTERNAME
$script:TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$script:DebugLogFile = Join-Path $PSScriptRoot "Remove-NTFSPermissionsForSIDs_${script:ComputerName}_${script:TimeStamp}_debug.log"
$script:TranscriptFile = Join-Path $PSScriptRoot "Remove-NTFSPermissionsForSIDs_${script:ComputerName}_${script:TimeStamp}_transcript.log"
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
    
    # Format console output to be more readable
    $consoleMessage = "$categoryText$Message"
    $debugMessage = "[$timestamp] [$Level] $categoryText$Message"
    
    # Write to console with color if not suppressed
    if (-not $NoConsole) {
        Write-Host $consoleMessage -ForegroundColor $Color
    }
    
    # Write detailed info to debug log
    Add-Content -Path $script:DebugLogFile -Value $debugMessage
}

# Function to handle progress bar with error handling
function Write-ProgressStatus {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Activity,
        [Parameter(Mandatory=$true)]
        [string]$Status,
        [Parameter(Mandatory=$true)]
        [int]$Current,
        [Parameter(Mandatory=$true)]
        [int]$Total,
        [int]$Id = 0
    )
    
    if (-not $EnableProgressBar) { return }
    
    try {
        $percentComplete = [math]::Min([math]::Round(($Current / [math]::Max($Total, 1)) * 100), 100)
        Write-Progress -Activity $Activity `
                      -Status "$Status ($Current of $Total)" `
                      -PercentComplete $percentComplete `
                      -Id $Id
    }
    catch {
        Write-Log "Error updating progress bar: $_" -Level 'WARNING' -Color "Yellow" -NoConsole
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

# Enhanced function to handle SID to name conversion with better error handling
function Convert-SidToName {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Sid
    )
    
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
        Write-Log "Failed to resolve SID ${Sid}: $_" -Level 'DEBUG' -Color "Magenta" -NoConsole
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

# Fixed Confirm-SIDRemoval function with proper variable handling
function Confirm-SIDRemoval {
    param (
        [Parameter(Mandatory=$true)]
        [string]$SID,
        [string]$Name = $null
    )

    if ($script:ApprovedSIDRemovals.ContainsKey($SID)) {
        return $script:ApprovedSIDRemovals[$SID]
    }

    $displayText = if ($Name -and $Name -ne $SID) { 
        "$Name ($SID)" 
    } else { 
        $SID 
    }
    
    Write-Host "Do you want to remove permissions for $displayText? (Y/N)" -ForegroundColor Yellow
    $confirmation = Read-Host
    $approved = $confirmation -eq 'Y'
    $script:ApprovedSIDRemovals[$SID] = $approved

    $logLevel = if ($approved) { 'INFO' } else { 'WARNING' }
    $logMessage = if ($approved) { "approved" } else { "denied" }
    Write-Log "User $logMessage removal of permissions for $displayText" -Level $logLevel

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
                Write-Log "Could not translate $($ace.IdentityReference) to SID" -Level 'WARNING' -NoConsole
                continue
            }
        }

        if ($script:TargetSIDs -contains $sid) {
            $name = Convert-SidToName -Sid $sid
            # Make sure we have a valid name before calling Confirm-SIDRemoval
            if ([string]::IsNullOrEmpty($name)) { $name = $sid }
            
            # Use try/catch to help identify any issues
            try {
                if (Confirm-SIDRemoval -SID $sid -Name $name) {
                    $Acl.RemoveAccessRule($ace) | Out-Null
                    $modified = $true
                    $removedCount++
                    Write-Log "Removed permission for $name from $StartPath" -Level 'SUCCESS' -Color "Green"
                }
            }
            catch {
                Write-Log "Error in permission removal process: $_" -Level 'ERROR' -Color "Red"
                Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level 'DEBUG' -NoConsole
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

# Enhanced folder processing with better progress tracking
function Invoke-FolderProcessing {
    param(
        [string]$StartPath,
        [int]$CurrentCount,
        [int]$TotalCount
    )

    $startTime = Get-Date
    try {
        # Update main progress bar
        Write-ProgressStatus -Activity "Analyzing and Updating Folder Permissions" `
                           -Status "Processing: $StartPath" `
                           -Current $CurrentCount `
                           -Total $TotalCount `
                           -Id 0

        # Update sub-progress for current folder
        Write-ProgressStatus -Activity "Current Folder Analysis" `
                           -Status "Checking permissions and SIDs" `
                           -Current 1 `
                           -Total 3 `
                           -Id 1

        # Get current ACL
        $acl = Get-Acl -Path $StartPath

        # Update sub-progress
        Write-ProgressStatus -Activity "Current Folder Analysis" `
                           -Status "Removing target SID permissions" `
                           -Current 2 `
                           -Total 3 `
                           -Id 1

        # Check and remove target SID permissions
        $modified = Remove-TargetSIDPermissions -StartPath $StartPath -Acl $acl

        # Final sub-progress update
        Write-ProgressStatus -Activity "Current Folder Analysis" `
                           -Status "Finalizing changes" `
                           -Current 3 `
                           -Total 3 `
                           -Id 1

        # Get updated permissions for logging
        if ($modified) {
            $acl = Get-Acl -Path $StartPath
        }

        Write-Log -Message "Processed: $StartPath" -Level 'INFO' -NoConsole
    }
    catch {
        Write-Log -Message "Error processing folder ${StartPath}: $_" -Level 'ERROR' -Color "Red"
    }
    finally {
        Write-PerformanceMetric -Operation "Folder processing for $StartPath" -StartTime $startTime
        Write-Progress -Id 1 -Completed
    }
}

# Modified function to use adminAccounts
function Test-AdministratorSID {
    param(
        [string]$SID
    )
    
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
    # Start transcript with unique name
    if (-not $script:TranscriptStarted) {
        $script:TranscriptStarted = $true
        Start-Transcript -Path $script:TranscriptFile -Force
    }

    # Initialize debug log
    $null = New-Item -ItemType File -Path $script:DebugLogFile -Force

    Write-Log "Starting permission analysis on $StartPath" -Color Cyan
    Write-Log "Debug log: $script:DebugLogFile" -Level 'DEBUG' -NoConsole
    Write-Log "Transcript: $script:TranscriptFile" -Level 'DEBUG' -NoConsole

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

    # Add detailed logging for SID initialization
    Write-Log "Starting folder permission analysis and removal for $StartPath" -Color Cyan
    Write-Log "Debug: Current well-known SIDs:" -Level 'DEBUG'
    $script:WellKnownSIDs.Keys | ForEach-Object {
        Write-Log "  $_ : $($script:WellKnownSIDs[$_])" -Level 'DEBUG' -NoConsole
    }

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
    Write-Log "Critical error encountered:" -Level 'ERROR' -Color 'Red'
    Write-Log "Error message: $_" -Level 'ERROR' -Color 'Red'
    Write-Log "Stack trace:" -Level 'ERROR' -Color 'Red'
    Write-Log $_.ScriptStackTrace -Level 'ERROR' -Color 'Red'
    
    if ($_.Exception.InnerException) {
        Write-Log "Inner exception:" -Level 'ERROR' -Color 'Red'
        Write-Log $_.Exception.InnerException.Message -Level 'ERROR' -Color 'Red'
    }
    
    # Ensure the error is properly recorded in the PowerShell error stream
    $PSCmdlet.ThrowTerminatingError($_)
}

finally {
    # Stop transcript safely
    if ($script:TranscriptStarted) {
        try {
            Stop-Transcript
        }
        catch {
            Write-Error "Failed to stop transcript: $_"
        }
    }
    Write-Log -Message "Script execution completed" -Level 'INFO' -Color "Green"    
}