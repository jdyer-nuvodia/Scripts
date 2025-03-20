# =============================================================================
# Script: Remove-NTFSPermissionsForSIDs.ps1
# Created: 2025-03-18 17:20:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-22 22:23:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2.1
# Additional Info: Fixed script hanging after SID approvals by improving queue management and tracking
# =============================================================================

<#
.SYNOPSIS
Removes permissions for specified SIDs from NTFS folder structures.
.DESCRIPTION
Analyzes folder structures and removes permissions for target SIDs listed in SIDs.txt.
Includes confirmation prompts and detailed logging of all changes.
.PARAMETER StartPath
The root folder path to begin permission analysis and removal.
.PARAMETER MaxDepth
Maximum folder depth to recurse. Default is 0 (unlimited).
.PARAMETER TimeoutMinutes
Maximum execution time in minutes before cancellation. Default is 120.
.EXAMPLE
.\Remove-NTFSPermissionsForSIDs.ps1 -StartPath "D:\Data" -MaxDepth 5
Processes the D:\Data folder structure to a maximum depth of 5 folders.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [string]$StartPath,
    
    [Parameter()]
    [int]$MaxDepth = 0,
    
    [Parameter()]
    [int]$TimeoutMinutes = 120
)

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
$script:EnableProgressBar = $true
$script:StartTime = Get-Date
$script:TotalFolders = 0
$script:ProcessedFolders = 0
$script:cancellationTokenSource = New-Object System.Threading.CancellationTokenSource
$script:TargetSIDs = [System.Collections.Generic.HashSet[string]]::new()
$script:ApprovedSIDRemovals = @{}
$script:processingTimeout = New-TimeSpan -Minutes $TimeoutMinutes
$script:FolderQueue = [System.Collections.Generic.Queue[string]]::new()

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

# Enhanced progress bar handling function
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
    
    if (-not $script:EnableProgressBar) { return }
    
    try {
        $percentComplete = [math]::Min([math]::Round(($Current / $Total) * 100), 100)
        
        $progressParams = @{
            Activity = $Activity
            Status = "$Status ($Current of $Total)"
            PercentComplete = $percentComplete
            Id = $Id
        }
        
        Write-Progress @progressParams
    }
    catch {
        Write-Log "Error updating progress bar: $_" -Level 'WARNING' -NoConsole
    }
}

# Function to handle performance metrics
function Write-PerformanceMetric {
    param (
        [string]$Operation,
        [DateTime]$StartTime
    )
    
    $endTime = Get-Date
    $duration = (New-TimeSpan -Start $StartTime -End $endTime).TotalMilliseconds
    $rate = if ($duration -gt 0) { [math]::Round(1000 / $duration, 2) } else { 0 }
    
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

# Fixed Confirm-SIDRemoval function with improved display
function Confirm-SIDRemoval {
    param (
        [Parameter(Mandatory=$true)]
        [string]$SID,
        [string]$Name = $null,
        [string]$Path
    )

    if ($script:ApprovedSIDRemovals.ContainsKey($SID)) {
        return $script:ApprovedSIDRemovals[$SID]
    }

    # Ensure we have a name to display
    $displayName = if ($Name -and $Name -ne $SID) { 
        "$Name ($SID)" 
    } else { 
        $SID 
    }
    
    Write-Host "`n=== Permission Removal Confirmation ===" -ForegroundColor Cyan
    Write-Host "Progress: $($script:ProcessedFolders) of $($script:TotalFolders) folders processed" -ForegroundColor White
    Write-Host "Current Path: $Path" -ForegroundColor White
    Write-Host "Target SID: $displayName" -ForegroundColor Yellow
    $confirmation = Read-Host "Do you want to remove these permissions? (Y/N)"
    $approved = $confirmation -eq 'Y' -or $confirmation -eq 'y'
    $script:ApprovedSIDRemovals[$SID] = $approved

    $logLevel = if ($approved) { 'INFO' } else { 'WARNING' }
    Write-Log "User $('approved' * $approved + 'denied' * -not $approved) removal of permissions for $displayName" -Level $logLevel

    return $approved
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

# Function to load target SIDs from file
function Import-TargetSIDs {
    $sidFilePath = Join-Path $PSScriptRoot "SIDs.txt"
    
    if (-not (Test-Path $sidFilePath)) {
        Write-Log "SIDs.txt file not found at $sidFilePath!" -Level 'ERROR' -Color "Red"
        throw "SIDs.txt file not found. Please create this file with one SID per line."
    }
    
    $sids = Get-Content -Path $sidFilePath | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    
    foreach ($sid in $sids) {
        $script:TargetSIDs.Add($sid.Trim()) | Out-Null
    }
    
    Write-Log "Loaded $($script:TargetSIDs.Count) SIDs from SIDs.txt" -Level 'INFO' -Color "Cyan"
}

# Function to check if SID is an administrator account
function Test-AdministratorSID {
    param(
        [string]$SID
    )
    
    if ([string]::IsNullOrEmpty($SID)) { return $false }
    
    # Check against well-known admin SIDs
    foreach ($admin in $script:AdminAccounts) {
        if ($SID -eq $admin.SID) {
            return $true
        }
    }
    
    return $false
}

# Function to remove permissions for target SIDs with improved error handling
function Remove-TargetSIDPermissions {
    param (
        [Parameter(Mandatory=$true)]
        [string]$StartPath,
        
        [Parameter(Mandatory=$true)]
        [System.Security.AccessControl.DirectorySecurity]$Acl
    )
    
    $modified = $false
    $removedCount = 0
    
    # Get all access rules
    $accessRules = $Acl.Access
    
    foreach ($rule in $accessRules) {
        $sidString = if ($rule.IdentityReference -match '^S-1-') {
            $rule.IdentityReference.Value
        } else {
            try {
                $ntAccount = [System.Security.Principal.NTAccount]$rule.IdentityReference
                $sid = $ntAccount.Translate([System.Security.Principal.SecurityIdentifier])
                $sid.Value
            } catch {
                continue
            }
        }
        
        # Skip inherited permissions
        if ($rule.IsInherited) { continue }
        
        # Check if this is one of our target SIDs
        if ($script:TargetSIDs -contains $sidString) {
            # Get display name for confirmation
            $displayName = Convert-SidToName -Sid $sidString
            
            # Check if it's an administrator SID
            if (Test-AdministratorSID -SID $sidString) {
                Write-Log "Skipping administrator SID: $displayName ($sidString)" -Level 'WARNING' -Color "Yellow"
                continue
            }
            
            # Confirm removal
            if (Confirm-SIDRemoval -SID $sidString -Name $displayName -Path $StartPath) {
                try {
                    $Acl.RemoveAccessRule($rule) | Out-Null
                    $removedCount++
                    $modified = $true
                    Write-Log "Removed permission for $displayName from $StartPath" -Level 'SUCCESS' -Color "Green"
                }
                catch {
                    Write-Log "Failed to remove permission for ${sidString}: $_" -Level 'ERROR' -Color "Red"
                }
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

# Enhanced folder processing with better progress tracking and error handling
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

        # Verify path exists before processing
        if (-not (Test-Path -Path $StartPath -ErrorAction SilentlyContinue)) {
            Write-Log "Path does not exist or is inaccessible: $StartPath" -Level 'WARNING' -Color "Yellow"
            return
        }

        # Verify permissions before attempting changes
        try {
            $testFile = Join-Path $env:TEMP "ACLTest_$(Get-Random).tmp"
            New-Item -Path $testFile -ItemType File -ErrorAction Stop | Out-Null
            Remove-Item -Path $testFile -Force -ErrorAction Stop | Out-Null
        } catch {
            Write-Log "Insufficient permissions to make changes to the file system. Run as administrator." -Level 'ERROR' -Color "Red"
            return
        }

        # Get current ACL
        try {
            $acl = Get-Acl -Path $StartPath -ErrorAction Stop
            
            # Check and remove target SID permissions
            $modified = Remove-TargetSIDPermissions -StartPath $StartPath -Acl $acl
            
            # Log whether modifications were made
            if ($modified) {
                Write-Log "Permissions were modified on $StartPath" -Level 'INFO' -Color "Cyan"
            } else {
                Write-Log "No permission changes needed for $StartPath" -Level 'INFO' -Color "DarkGray" -NoConsole
            }
        }
        catch {
            Write-Log "Error getting ACL for ${StartPath}: $_" -Level 'ERROR' -Color "Red"
        }
    }
    catch {
        Write-Log "Error processing folder ${StartPath}: $_" -Level 'ERROR' -Color "Red"
    }
    finally {
        Write-PerformanceMetric -Operation "Folder processing for $StartPath" -StartTime $startTime
    }
}

# Non-recursive folder processing using queue approach
function Invoke-FolderTree {
    param (
        [string]$RootPath
    )
    
    Write-Log "Starting folder permission analysis and removal for $RootPath" -Level 'INFO' -Color "Cyan"
    
    # First estimate total folder count for progress
    Write-Host "Estimating folder count..." -ForegroundColor Cyan
    try {
        $script:TotalFolders = (Get-ChildItem -Path $RootPath -Directory -Recurse -ErrorAction SilentlyContinue | Measure-Object).Count + 1
    }
    catch {
        # If recursion fails, set a default value
        $script:TotalFolders = 1000
        Write-Log "Could not estimate folder count: $_" -Level 'WARNING' -Color "Yellow"
    }
    Write-Host "Estimated $($script:TotalFolders) folders to process" -ForegroundColor Cyan
    
    # Initialize the queue with the root folder
    $script:FolderQueue.Clear()
    $script:FolderQueue.Enqueue($RootPath)
    $script:ProcessedFolders = 0
    
    Write-Host "Beginning folder processing..." -ForegroundColor Green
    
    # Process folders from the queue until empty
    while ($script:FolderQueue.Count -gt 0 -and -not $script:cancellationTokenSource.Token.IsCancellationRequested) {
        # Check timeout
        $currentDuration = (Get-Date) - $script:StartTime
        if ($currentDuration -gt $script:processingTimeout) {
            Write-Log "Processing timeout reached ($TimeoutMinutes minutes). Canceling further processing." -Level 'WARNING' -Color "Yellow"
            break
        }
        
        $currentFolder = $script:FolderQueue.Dequeue()
        
        # Process current folder
        Invoke-FolderProcessing -StartPath $currentFolder -CurrentCount $script:ProcessedFolders -TotalCount $script:TotalFolders
        
        # Increment the processed folders counter AFTER processing
        $script:ProcessedFolders++
        
        # Update progress bar to reflect current state
        Write-ProgressStatus -Activity "Analyzing and Updating Folder Permissions" `
                           -Status "Processed: $currentFolder" `
                           -Current $script:ProcessedFolders `
                           -Total $script:TotalFolders `
                           -Id 0
        
        # Check depth before enqueuing subfolders
        $currentDepth = ($currentFolder.Split('\').Length - $RootPath.Split('\').Length)
        if ($MaxDepth -eq 0 -or $currentDepth -lt $MaxDepth) {
            # Add subfolders to queue
            try {
                $subFolders = Get-ChildItem -Path $currentFolder -Directory -ErrorAction SilentlyContinue
                foreach ($folder in $subFolders) {
                    $script:FolderQueue.Enqueue($folder.FullName)
                }
                
                # Log number of subfolders added for debugging
                if ($subFolders.Count -gt 0) {
                    Write-Log "Added $($subFolders.Count) subfolders from $currentFolder to processing queue" -Level 'DEBUG' -Color "DarkGray" -NoConsole
                }
            }
            catch {
                Write-Log "Error accessing subfolders of $currentFolder`: $_" -Level 'WARNING' -Color "Yellow"
            }
        }
        
        # Log queue status periodically
        if ($script:ProcessedFolders % 100 -eq 0) {
            Write-Log "Current queue status: $($script:ProcessedFolders) folders processed, $($script:FolderQueue.Count) folders remaining in queue" -Level 'INFO' -Color "Cyan"
        }
    }
    
    # Complete progress bar
    Write-Progress -Activity "Analyzing and Updating Folder Permissions" -Completed
    Write-Log "Completed processing $($script:ProcessedFolders) folders" -Level 'INFO' -Color "Green"
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
    
    # Initialize admin accounts
    $script:AdminAccounts = Initialize-WellKnownSIDs
    Write-Log "Debug: Current well-known SIDs:" -Level 'DEBUG' -Color "Magenta" -NoConsole
    $script:AdminAccounts | ForEach-Object {
        Write-Log "  $($_.Name): $($_.SID)" -Level 'DEBUG' -Color "Magenta" -NoConsole
    }
    
    # Process folder tree using non-recursive approach
    Invoke-FolderTree -RootPath $StartPath
    
    # Summary
    $elapsed = (Get-Date) - $script:StartTime
    Write-Log "Permission analysis and removal completed in $($elapsed.TotalMinutes.ToString('0.00')) minutes" -Level 'SUCCESS' -Color "Green"
    Write-Log "Processed $($script:ProcessedFolders) of $($script:TotalFolders) folders" -Level 'SUCCESS' -Color "Green"
}
catch [System.Exception] {
    Write-Log "Error: $_" -Level 'ERROR' -Color "Red"
    Write-Error $_.ScriptStackTrace
}
finally {
    # Stop cancellation token source
    $script:cancellationTokenSource.Dispose()
    
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