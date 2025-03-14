# =============================================================================
# Script: Get-NTFSFolderPermissions.ps1
# Created: 2025-02-07 21:21:53 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-14 23:56:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.10.10
# Additional Info: Fixed UTF-8 character encoding in output
# =============================================================================

<#
.SYNOPSIS
Gets NTFS folder permissions for specified path.

.DESCRIPTION
Analyzes and reports NTFS permissions for specified folder path and its subfolders.
Consolidates output into two log files:
- Main log for permission details
- Debug log for troubleshooting information

.PARAMETER FolderPath
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
.\Get-NTFSFolderPermissions.ps1 -FolderPath "C:\Temp"
Analyzes permissions on C:\Temp and outputs to logs
#>

using namespace System.Security.AccessControl
using namespace System.IO
using namespace System.Security.Principal

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$FolderPath,

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

# Script-level variables - consolidated to avoid duplication
$script:TranscriptStarted = $false
$script:SidCache = @{}  # Single standardized cache
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
# Define script-level retry values for consistency
$script:MaxRetries = 3
$script:RetryDelay = 2

# Add script-level cancellation token
$script:cancellationTokenSource = New-Object System.Threading.CancellationTokenSource
$script:processingTimeout = New-TimeSpan -Minutes $TimeoutMinutes

# Add function for sanitizing path for filename
function Get-SafeFilename {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )
    
    try {
        # Initial null/empty check
        if ([string]::IsNullOrEmpty($Path)) {
            throw "Path cannot be null or empty"
        }

        # Convert path separators to underscores
        $safeName = $Path.Replace('\', '_').Replace('/', '_')
        
        # Replace invalid characters but preserve dashes and spaces
        $safeName = $safeName -replace '[:<>"|?*]', '_'
        
        # Handle multiple spaces/underscores
        $safeName = $safeName -replace '[\s_]+', '_'
        
        # Remove leading/trailing underscores
        $safeName = $safeName.Trim('_')
        
        # Length validation with proper checks
        if ([string]::IsNullOrEmpty($safeName)) {
            throw "Sanitized path resulted in empty string"
        }

        # Safe substring operation
        if ($safeName.Length -gt 50) {
            $safeName = $safeName.Substring(0, [Math]::Min(47, $safeName.Length)) + "..."
        }
        
        # Final validation
        if ([string]::IsNullOrEmpty($safeName)) {
            throw "Final sanitized path is empty"
        }
        
        return $safeName
    }
    catch {
        Write-Log -Message "Error in Get-SafeFilename: $_" -Level 'ERROR' -Color "Red"
        # Return a safe default name that includes part of the original path
        $defaultName = "DefaultLog_" + (Get-Date -Format 'yyyyMMdd_HHmmss')
        if (-not [string]::IsNullOrEmpty($Path)) {
            # Take last part of path if available
            $lastPart = $Path.Split('\')[-1]
            if (-not [string]::IsNullOrEmpty($lastPart)) {
                $defaultName = "Log_" + ($lastPart -replace '[^\w\-]', '_')
            }
        }
        return $defaultName
    }
}

# Consolidated log initialization with enhanced configuration
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$computerName = $env:COMPUTERNAME
$safePath = Get-SafeFilename -Path $FolderPath
$logBase = Join-Path $PSScriptRoot "NTFSPermissions_${computerName}_${safePath}_${timestamp}"
$script:DebugLogFile = "${logBase}_debug.log"

# Function to create a standardized log header with enhanced metadata
function New-LogHeader {
    @"
# =============================================================================
# NTFS Permissions Debug Log
# Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC")
# System: $computerName
# PowerShell Version: $($PSVersionTable.PSVersion)
# Analysis Path: $FolderPath
# Max Threads: $MaxThreads
# Max Depth: $MaxDepth
# Skip AD Resolution: $SkipADResolution
# Skip Uniqueness Counting: $SkipUniquenessCounting
# Enable SID Diagnostics: $EnableSIDDiagnostics
# Script Version: 1.10.1
# =============================================================================

"@
}

# Initialize debug log with proper header
Set-Content -Path $script:DebugLogFile -Value (New-LogHeader)

# Define Write-Log function with enhanced functionality
function Write-Log {
    param (
        [string]$Message,
        [string]$Color = "White",
        [switch]$NoConsole,
        [switch]$Debug,
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'DEBUG', 'SUCCESS')]
        [string]$Level = $(if ($Debug) { 'DEBUG' } else { 'INFO' })
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff UTC"
    $callStack = Get-PSCallStack
    $callingFunction = $callStack[1].FunctionName
    $lineNumber = $callStack[1].ScriptLineNumber
    if ($callingFunction -eq "<ScriptBlock>") { $callingFunction = "MainScript" }
    
    $logEntry = @"
[TIMESTAMP: $timestamp]
[LEVEL: $Level]
[FUNCTION: $callingFunction]
[LINE: $lineNumber]
[THREAD: $([Threading.Thread]::CurrentThread.ManagedThreadId)]
[MESSAGE: $Message]
----------------------------------------
"@
    
    # Write to debug log with error handling
    try {
        Add-Content -Path $script:DebugLogFile -Value $logEntry
    }
    catch {
        Write-Warning "Failed to write to debug log: $_"
    }
    
    if (-not $NoConsole) {
        $levelColors = @{
            'INFO' = 'White'
            'WARNING' = 'Yellow'
            'ERROR' = 'Red'
            'DEBUG' = 'Magenta'
            'SUCCESS' = 'Green'
        }
        $messageColor = if ($levelColors.ContainsKey($Level)) { $levelColors[$Level] } else { $Color }
        Write-Host $Message -ForegroundColor $messageColor
    }
}

# Import required modules
Import-Module ActiveDirectory -ErrorAction SilentlyContinue

# Initialize transcript - done only once
if ($Host.Name -eq 'ConsoleHost' -and -not $script:TranscriptStarted) {
    $transcriptPath = "${logBase}_transcript.log"
    try {
        Start-Transcript -Path $transcriptPath -Force | Out-Null
        $script:TranscriptStarted = $true
        Write-Log -Message "Initializing transcript at: $transcriptPath" -Level 'INFO' -Color "Cyan"
    }
    catch {
        Write-Log -Message "Could not start transcript: $_" -Level 'ERROR' -Color "Red"
    }
}

# Initialize well-known SIDs
function Initialize-WellKnownSIDs {
    # Get Administrator SID using WMI first
    $script:AdminSID = (Get-WmiObject Win32_UserAccount -Filter "Name='Administrator'" -ErrorAction SilentlyContinue).SID

    $script:WellKnownSIDs = @{
        "Nobody" = "S-1-0-0"
        "Everyone" = "S-1-1-0"
        "Local" = "S-1-2-0"
        "CreatorOwner" = "S-1-3-0"
        "CreatorGroup" = "S-1-3-1"
        "Network" = "S-1-5-2"
        "Interactive" = "S-1-5-4"
        "AuthenticatedUsers" = "S-1-5-11"
        "LocalSystem" = "S-1-5-18"
        "LocalService" = "S-1-5-19"
        "NetworkService" = "S-1-5-20"
        "Administrator" = if ($script:AdminSID) { $script:AdminSID } else { "S-1-5-21-domain-500" }
        "Administrators" = "S-1-5-32-544"
        "Users" = "S-1-5-32-545"
        "Guests" = "S-1-5-32-546"
    }
    Write-Log -Message "Initialized well-known SIDs collection" -Color "DarkGray" -NoConsole -Level 'DEBUG'
}

function Test-WellKnownSID {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Sid
    )
    
    if ($script:WellKnownSIDs.Values -contains $Sid) {
        $name = $script:WellKnownSIDs.GetEnumerator() | 
            Where-Object { $_.Value -eq $Sid } | 
            Select-Object -First 1 -ExpandProperty Key
        return $name
    }
    return $null
}

# Initialize suppressed SIDs
$script:SuppressedSids.Add('S-1-5-21-3715258189-2875184700-594828381-500')
$script:SuppressedSids.Add('S-1-5-21-1787995930-3758959370-1315816792-13767')
$script:SuppressedSids.Add('S-1-5-21-1787995930-3758959370-1315816792-13821')
$script:SuppressedSids.Add('S-1-5-21-1787995930-3758959370-1315816792-17638')

# Consolidated SID handling function
function Convert-SidToName {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Sid
    )
    
    if ($script:SuppressedSids -contains $Sid) {
        Write-Log -Message "Skipping suppressed SID: $Sid" -Color "DarkGray" -NoConsole -Level 'DEBUG'
        return $Sid
    }
    
    if ($script:SidCache.ContainsKey($Sid)) {
        return $script:SidCache[$Sid]
    }
    
    $wellKnownName = Test-WellKnownSID -Sid $Sid
    if ($wellKnownName) {
        $script:SidCache[$Sid] = $wellKnownName
        return $wellKnownName
    }

    # If we've already failed this SID max times, return it immediately
    if ($script:SidTranslationAttempts[$Sid] -ge $script:MaxRetries) {
        return $Sid
    }

    try {
        if (-not $script:SidTranslationAttempts.ContainsKey($Sid)) {
            $script:SidTranslationAttempts[$Sid] = 0
        }
        
        $script:SidTranslationAttempts[$Sid]++
        $attempt = $script:SidTranslationAttempts[$Sid]
        
        Write-Log -Message "Attempting SID translation (attempt $attempt of $script:MaxRetries): $Sid" -Color "DarkGray" -NoConsole -Level 'DEBUG'
        
        # Try to resolve using .NET first
        try {
            $objSID = New-Object System.Security.Principal.SecurityIdentifier($Sid)
            $objName = $objSID.Translate([System.Security.Principal.NTAccount])
            $name = $objName.Value
            $script:SidCache[$Sid] = $name
            Write-Log -Message "Successfully resolved SID on attempt ${attempt}: ${Sid} -> ${name}" -Color "Green" -NoConsole -Level 'SUCCESS'
            return $name
        }
        catch {
            # Fall back to AD lookup if .NET translation fails
            Write-Log -Message ".NET translation failed on attempt ${attempt}, trying AD lookup for SID: $Sid" -Color "DarkGray" -NoConsole -Level 'DEBUG'
            if (-not $SkipADResolution) {
                try {
                    $user = Get-ADUser -Identity $Sid -Properties SamAccountName -ErrorAction Stop
                    if ($user) {
                        $name = $user.SamAccountName
                        $script:SidCache[$Sid] = $name
                        Write-Log -Message "Successfully resolved SID via AD on attempt ${attempt}: ${Sid} -> ${name}" -Color "Green" -NoConsole -Level 'SUCCESS'
                        return $name
                    }
                }
                catch {
                    # Try as a group if user lookup failed
                    try {
                        $group = Get-ADGroup -Identity $Sid -Properties SamAccountName -ErrorAction Stop
                        if ($group) {
                            $name = $group.SamAccountName
                            $script:SidCache[$Sid] = $name
                            Write-Log -Message "Successfully resolved SID via AD (group) on attempt ${attempt}: ${Sid} -> ${name}" -Color "Green" -NoConsole -Level 'SUCCESS'
                            return $name
                        }
                    }
                    catch {
                        Write-Log -Message "AD resolution failed for SID: $Sid" -Color "Yellow" -NoConsole -Level 'WARNING'
                        throw "Failed to resolve SID via AD: $_"
                    }
                }
            }
            else {
                throw "AD resolution skipped by user"
            }
        }
    }
    catch {
        Write-Log -Message "SID translation failed on attempt ${attempt}: ${Sid}" -Color "Yellow" -NoConsole -Level 'WARNING'
        if ($script:SidTranslationAttempts[$Sid] -lt $script:MaxRetries) {
            Write-Log -Message "Retrying in $script:RetryDelay seconds (attempt ${attempt}/${script:MaxRetries})..." -Color "DarkGray" -NoConsole -Level 'DEBUG'
            Start-Sleep -Seconds $script:RetryDelay
            return Convert-SidToName -Sid $Sid
        }
        $script:ADResolutionErrors[$Sid] = $_.Exception.Message
        $script:FailedSids.Add($Sid) | Out-Null
    }
    
    # If all resolution attempts fail, return the original SID
    return $Sid
}

# Standardized function to generate permission hashes
function Get-PermissionHash {
    param (
        [Parameter(Mandatory=$true)]
        [object]$AccessRules,
        
        [Parameter(Mandatory=$false)]
        [bool]$IncludeInheritance = $true
    )
    
    if ($IncludeInheritance) {
        return ($AccessRules | ForEach-Object { 
            "$($_.IdentityReference)|$($_.FileSystemRights)|$($_.AccessControlType)|$($_.IsInherited)" 
        }) -join ';'
    }
    else {
        return ($AccessRules | ForEach-Object { 
            "$($_.IdentityReference)|$($_.FileSystemRights)|$($_.AccessControlType)" 
        }) -join ';'
    }
}

# Write progress function for consistent progress reporting
function Write-ProgressStatus {
    param (
        [string]$Activity,
        [string]$Status,
        [int]$Current,
        [int]$Total
    )
    
    $percentComplete = [math]::Round(($Current / $Total) * 100, 2)
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $percentComplete
}

# Modify Get-FolderPermissions to use standardized hash generation
function Get-FolderPermissions {
    param (
        [string]$Folder
    )
    try {
        $acl = Get-Acl -Path $Folder
        $owner = $acl.Owner
        $access = $acl.Access
        $isInherited = $acl.AreAccessRulesProtected -eq $false
        $parentPath = Split-Path -Path $Folder -Parent
        $permissionHash = Get-PermissionHash -AccessRules $acl.Access -IncludeInheritance $true
        
        # Check if permissions match parent
        $matchesParent = $false
        if ($parentPath -and $script:ParentPermissions.ContainsKey($parentPath)) {
            $matchesParent = $script:ParentPermissions[$parentPath] -eq $permissionHash
        }

        $permissionData = [PSCustomObject]@{
            Folder = $Folder
            Owner  = $owner
            Access = $access
            PermissionHash = $permissionHash
            IsInherited = $isInherited
            MatchesParent = $matchesParent
        }

        # Store current folder's permissions for child comparison
        $script:ParentPermissions[$Folder] = $permissionHash

        return $permissionData
    }
    catch {
        Write-Log -Message "Error getting permissions for ${Folder}: $($_.Exception.Message)" -Color "Red" -Level 'ERROR'
        return $null
    }
}

# Function to process each folder
function Invoke-FolderProcessing {
    param(
        [string]$Path,
        [int]$CurrentCount,
        [int]$TotalCount
    )

    try {
        # Get permissions for the folder
        $permissionData = Get-FolderPermissions -Folder $Path

        if ($permissionData) {
            if (-not $script:PermissionGroups.ContainsKey($permissionData.PermissionHash)) {
                $script:PermissionGroups[$permissionData.PermissionHash] = @{
                    Folders = @()
                    Permissions = $permissionData.Access
                    Owner = $permissionData.Owner
                    IsInherited = $permissionData.IsInherited
                    ParentPaths = @{}
                }
            }

            # Group by parent path for better organization
            $currentParent = Split-Path -Path $Path -Parent
            if (-not [string]::IsNullOrEmpty($currentParent)) {
                if (-not $script:PermissionGroups[$permissionData.PermissionHash].ParentPaths.ContainsKey($currentParent)) {
                    $script:PermissionGroups[$permissionData.PermissionHash].ParentPaths[$currentParent] = @()
                }
                $script:PermissionGroups[$permissionData.PermissionHash].ParentPaths[$currentParent] += $Path
            }

            $script:PermissionGroups[$permissionData.PermissionHash].Folders += $Path
            $script:FolderPermissions[$Path] = $permissionData

            # Update unique permissions count - only write to debug log
            if (-not $SkipUniquenessCounting) {
                $permissionHash = Get-PermissionHash -AccessRules $permissionData.Access -IncludeInheritance:$false
                
                if (-not $script:UniquePermissions.ContainsKey($permissionHash)) {
                    $script:UniquePermissions[$permissionHash] = @($Path)
                } else {
                    $script:UniquePermissions[$permissionHash] += $Path
                }
            }
        } else {
            Write-Log -Message "Could not retrieve permissions for $Path" -Color "Yellow" -Level 'WARNING'
        }
    }
    catch {
        Write-Log -Message "Error processing folder ${Path}: $($_.Exception.Message)" -Color "Red" -Level 'ERROR'
    }
    finally {
        Write-ProgressStatus -Activity "Analyzing Folder Permissions" -Status "Processing: $Path" -Current $CurrentCount -Total $TotalCount
    }
}

# Function to process folders recursively
function Invoke-FolderRecursively {
    param (
        [string]$Path,
        [int]$CurrentDepth = 0
    )

    try {
        if ($script:cancellationTokenSource.Token.IsCancellationRequested) {
            Write-Log "Processing cancelled by user" -Level 'WARNING' -Color "Yellow"
            return
        }

        if ((Get-Date) - $script:StartTime -gt $script:processingTimeout) {
            $script:cancellationTokenSource.Cancel()
            Write-Log "Processing timeout reached ($TimeoutMinutes minutes)" -Level 'WARNING' -Color "Yellow"
            return
        }

        $script:TotalFolders++
        $script:ProcessedFolders++

        Invoke-FolderProcessing -Path $Path -CurrentCount $script:ProcessedFolders -TotalCount $script:TotalFolders

        if ($MaxDepth -eq 0 -or $CurrentDepth -lt $MaxDepth) {
            Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                Invoke-FolderRecursively -Path $_.FullName -CurrentDepth ($CurrentDepth + 1)
            }
        }
    }
    catch {
        Write-Log "Error processing folder $Path : $_" -Level 'ERROR' -Color "Red"
    }
}

# Add function to translate permissions to human-readable format
function Get-HumanReadablePermissions {
    param (
        [System.Security.AccessControl.FileSystemRights]$Rights
    )
    
    switch ($Rights) {
        { $_ -band [System.Security.AccessControl.FileSystemRights]::FullControl } { return "Full Control" }
        { $_ -band [System.Security.AccessControl.FileSystemRights]::Modify } { return "Modify" }
        { $_ -band [System.Security.AccessControl.FileSystemRights]::ReadAndExecute } { return "Read & Execute" }
        { $_ -band [System.Security.AccessControl.FileSystemRights]::Read } { return "Read Only" }
        { $_ -band [System.Security.AccessControl.FileSystemRights]::Write } { return "Write Only" }
        default { return $Rights.ToString() }
    }
}

# Main script execution
try {
    # Register ctrl+c handler
    $null = [Console]::TreatControlCAsInput = $true
    Register-ObjectEvent -InputObject ([Console]) -EventName CancelKeyPress -Action {
        $script:cancellationTokenSource.Cancel()
        Write-Log "Cancellation requested by user (Ctrl+C)" -Level 'WARNING' -Color "Yellow"
    } | Out-Null

    Initialize-WellKnownSIDs
    
    # Output initial messages in specified order - only once
    Write-Log -Message "Debug information will be written to: $script:DebugLogFile" -Level 'DEBUG'
    Write-Log ""
    Write-Log -Message "The Administrator SID is: $script:AdminSID" -Color "White" -Level 'INFO'
    Write-Log ""
    Write-Log -Message "Starting folder permission analysis for $FolderPath" -Color "Cyan" -Level 'INFO'
    
    # Process folders with timeout tracking
    Invoke-FolderRecursively -Path $FolderPath

    if ($script:cancellationTokenSource.Token.IsCancellationRequested) {
        Write-Log "`nProcessing terminated before completion" -Level 'WARNING' -Color "Yellow"
    }

    # Calculate elapsed time
    $script:EndTime = Get-Date
    $script:ElapsedTime = $script:EndTime - $script:StartTime

    # Display summary
    Write-Log -Message "`nAnalysis Complete" -Color "Green" -Level 'SUCCESS'
    Write-Log -Message "Total folders processed: $($script:ProcessedFolders)" -Color "Cyan" -Level 'INFO'
    Write-Log -Message "Unique permission sets: $($script:UniquePermissions.Count)" -Color "Cyan" -Level 'INFO'
    Write-Log -Message "Elapsed time: $($script:ElapsedTime.ToString())" -Color "Cyan" -Level 'INFO'

    # Display results in hierarchy mode
    Write-Log -Message "`nFolder Access Permissions:" -Color "Yellow" -Level 'INFO'
    $sortedFolders = $script:FolderPermissions.Keys | Sort-Object

    foreach ($folder in $sortedFolders) {
        $data = $script:FolderPermissions[$folder]

        # Skip folders with identical permissions as parent
        if ($data.MatchesParent) {
            continue
        }

        $inheritanceStatus = if ($data.IsInherited) { "(Inherits parent permissions)" } else { "(Custom permissions)" }
        Write-Log -Message "$folder $inheritanceStatus" -Color "Cyan" -Level 'INFO'

        # Find all descendants with identical permissions
        $identicalDescendants = @($script:PermissionGroups[$data.PermissionHash].Folders | 
            Where-Object { 
                $_ -and 
                $_ -ne $folder -and 
                $script:FolderPermissions[$_].MatchesParent 
            })

        if ($identicalDescendants -and $identicalDescendants.Count -gt 0) {
            Write-Log -Message "Subfolders with same permissions:" -Color "DarkGray" -Level 'INFO'
            foreach ($descendant in ($identicalDescendants | Sort-Object)) {
                if ($descendant -and $folder) {
                    $relativePath = $descendant.Substring($folder.Length + 1)
                    Write-Log -Message "  - $relativePath" -Color "DarkGray" -Level 'INFO'
                }
            }
        }

        Write-Log -Message "Access Rights:" -Color "White" -Level 'INFO'
        # Remove duplicate entries by using a hashtable to track unique permissions
        $uniqueAccessRules = @{}
        
        $data.Access | ForEach-Object {
            $identity = $_.IdentityReference.Value
            if ($identity -match '^S-1-') {
                $identity = Convert-SidToName -Sid $identity
            }
            
            $key = "$identity|$($_.FileSystemRights)|$($_.AccessControlType)"
            if (-not $uniqueAccessRules.ContainsKey($key)) {
                $uniqueAccessRules[$key] = $_
                
                $permission = Get-HumanReadablePermissions -Rights $_.FileSystemRights
                $accessType = if ($_.AccessControlType -eq 'Allow') { '+' } else { '-' }
                Write-Log -Message "  $accessType $identity - $permission" -Color "Gray" -Level 'INFO'
            }
        }
        
        Write-Log -Message "-" * 80 -Color "DarkGray" -Level 'INFO'
    }

    # Display SID resolution errors if any
    if ($EnableSIDDiagnostics -and $script:ADResolutionErrors.Count -gt 0) {
        Write-Log -Message "`nSID Resolution Errors:" -Color "Yellow" -Level 'WARNING'
        $script:ADResolutionErrors.GetEnumerator() | ForEach-Object {
            Write-Log -Message "SID: $($_.Key)" -Color "White" -Level 'INFO'
            Write-Log -Message "Error: $($_.Value)" -Color "Red" -Level 'ERROR'
        }
    }
}
catch [System.Exception] {
    Write-Error "An error occurred: $_"
    Write-Error $_.ScriptStackTrace
}
finally {
    Write-Progress -Activity "Analyzing Folder Permissions" -Completed
    
    try {
        Write-Host "Script execution completed. See $script:DebugLogFile for full details." -ForegroundColor Green
        
        # Clean up transcript only if we started one
        if ($script:TranscriptStarted) {
            Stop-Transcript -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-Log -Message "Error during cleanup: $_" -Level 'ERROR'
    }
    finally {
        # Properly clear important variables
        Remove-Variable -Name SidCache -Scope Script -ErrorAction SilentlyContinue
        Remove-Variable -Name FailedSids -Scope Script -ErrorAction SilentlyContinue
        Remove-Variable -Name SuppressedSids -Scope Script -ErrorAction SilentlyContinue
    }
    # Cleanup cancellation token
    if ($script:cancellationTokenSource) {
        $script:cancellationTokenSource.Dispose()
    }
}