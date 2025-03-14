# =============================================================================
# Script: Get-NTFSFolderPermissions.ps1
# Created: 2025-02-07 21:21:53 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-14 17:27:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.3.10
# Additional Info: Fixed script hanging issue with transcript and console output
# =============================================================================

# First all using statements
using namespace System.Security.AccessControl
using namespace System.IO
using namespace System.Security.Principal

# Then CmdletBinding
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
    [ValidateSet("Hierarchy", "Group")]
    [string]$ViewMode = "Hierarchy"
)

# Enable strict mode and error handling
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Import required modules
Import-Module ActiveDirectory -ErrorAction SilentlyContinue

# Initialize global variables
$script:TotalFolders = 0
$script:ProcessedFolders = 0
$script:StartTime = Get-Date
$script:EndTime = $null
$script:ElapsedTime = $null
$script:FolderPermissions = @{}
$script:UniquePermissions = @{}
$script:PermissionGroups = @{}
$script:ADObjectCache = @{}
$script:SIDCache = @{}
$script:ADResolutionErrors = @{}
$script:PermissionGroups = @{}
$script:InheritanceStatus = @{}
$script:ParentPermissions = @{}  # Add new script-level variables for tracking parent permissions
$script:SidCache = @{}  # Add SID translation cache
$script:DomainController = $null
$script:SidTranslationAttempts = @{}  # Add script-level variables for tracking SID translation attempts

# Add function for sanitizing path for filename
function Get-SafeFilename {
    param([string]$Path)
    # Replace invalid filename characters and common separators
    $safe = $Path -replace '[\\\/\:\*\?\"\<\>\|]', '_'
    # Replace multiple underscores with single underscore
    $safe = $safe -replace '_{2,}', '_'
    # Trim underscores from ends
    $safe = $safe.Trim('_')
    # Limit length to prevent extremely long filenames
    if ($safe.Length -gt 50) {
        $safe = $safe.Substring(0, 47) + '...'
    }
    return $safe
}

# Simplified log file creation - Fix the swapped log file names
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$systemName = $env:COMPUTERNAME
$safeFolderPath = Get-SafeFilename -Path $FolderPath

# Define log files with correct naming
$logBase = Join-Path $PSScriptRoot "NTFSPermissions_${systemName}_${safeFolderPath}_${timestamp}"
$script:DetailedLogFile = "${logBase}_detailed.log"
$script:ConsoleLogFile = "${logBase}_console.log"

# Start transcript to capture everything in the detailed log
Start-Transcript -Path $script:DetailedLogFile -Force

# Initialize console output collection
$script:ConsoleOutputCollection = [System.Collections.ArrayList]@()
[void]$script:ConsoleOutputCollection.Add("Starting folder permission analysis for $FolderPath")

# Output initial messages
Write-Host "Starting folder permission analysis for $FolderPath" -ForegroundColor Cyan
Write-Host "Detailed analysis will be written to: $script:DetailedLogFile" -ForegroundColor Yellow
Write-Host "Console summary will be written to: $script:ConsoleLogFile" -ForegroundColor Yellow

# Updated Write-Log function with proper output separation
function Write-Log {
    param (
        [string]$Message,
        [string]$Color = "White",
        [switch]$NoConsole,
        [switch]$DetailedOnly
    )
    
    # Always create detailed log entry
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $callingFunction = (Get-PSCallStack)[1].FunctionName
    if ($callingFunction -eq "<ScriptBlock>") { $callingFunction = "MainScript" }
    
    # Create detailed diagnostic information without Join-String
    $variables = $(Get-Variable -Scope 1 | 
        Where-Object { $_.Name -notlike "*Preference" } | 
        ForEach-Object { "$($_.Name)=$($_.Value)" }) -join "; "
    
    $callStack = $(Get-PSCallStack | 
        Select-Object -Skip 1 | 
        ForEach-Object { $_.Command }) -join " -> "
    
    $detailedEntry = @"
[TIMESTAMP: $timestamp UTC]
[FUNCTION: $callingFunction]
[THREAD_ID: $([System.Threading.Thread]::CurrentThread.ManagedThreadId)]
[MEMORY_USAGE: $([System.GC]::GetTotalMemory($false)) bytes]
[ACTION: $Message]
[VARIABLES: $variables]
[CALL_STACK: $callStack]
----------------------------------------
"@
    
    # Write to transcript (detailed log) using Write-Verbose
    Write-Verbose $detailedEntry
    
    # Handle console and console log output
    if (-not $DetailedOnly) {
        if (-not $NoConsole) {
            # Write to console with color
            Write-Host $Message -ForegroundColor $Color
        }
        
        # Add to console collection (for console log)
        if ($Message.Trim() -ne "") {
            [void]$script:ConsoleOutputCollection.Add($Message)
        }
    }
}

# Function to display progress bar
function Write-ProgressBar {
    param (
        [int]$Current,
        [int]$Total,
        [string]$Activity,
        [string]$Status
    )
    $percentComplete = [math]::Round(($Current / $Total) * 100, 2)
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $percentComplete
}

# Modify Get-FolderPermissions to check parent permissions
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
        $permissionHash = ($acl.Access | ForEach-Object { 
            "$($_.IdentityReference)|$($_.FileSystemRights)|$($_.AccessControlType)|$($_.IsInherited)" 
        }) -join ';'
        
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
        Write-Log -Message "Error getting permissions for ${Folder}: $($_.Exception.Message)" -Color "Red"
        return $null
    }
}

# Function to resolve SID to name
function Resolve-SIDToName {
    param (
        [string]$SID
    )
    if ($script:SIDCache.ContainsKey($SID)) {
        return $script:SIDCache[$SID]
    }
    try {
        $objSID = New-Object System.Security.Principal.SecurityIdentifier($SID)
        $objName = $objSID.Translate([System.Security.Principal.NTAccount])
        $name = $objName.Value
        $script:SIDCache[$SID] = $name
        return $name
    }
    catch {
        if ($EnableSIDDiagnostics) {
            $script:ADResolutionErrors[$SID] = $_.Exception.Message
        }
        return $SID
    }
}

# Add function to translate SIDs using AD lookup
$script:SidTranslationAttempts = @{}
$script:MaxRetries = 3
$script:RetryDelay = 2

function Convert-SidToName {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Sid
    )
    
    # Check cache first
    if ($script:SidCache.ContainsKey($Sid)) {
        Write-Log -Message "Using cached SID translation: $Sid -> $($script:SidCache[$Sid])" -Color "DarkGray" -NoConsole
        return $script:SidCache[$Sid]
    }
    
    # Check retry count
    if (-not $script:SidTranslationAttempts.ContainsKey($Sid)) {
        $script:SidTranslationAttempts[$Sid] = 0
    }
    elseif ($script:SidTranslationAttempts[$Sid] -ge $script:MaxRetries) {
        if (-not ($script:SuppressedSids -contains $Sid)) {
            Write-Log -Message "Maximum retry attempts ($script:MaxRetries) reached for SID: $Sid" -Color "Yellow"
        }
        $script:FailedSids.Add($Sid)
        return $Sid
    }

    try {
        $script:SidTranslationAttempts[$Sid]++
        $attempt = $script:SidTranslationAttempts[$Sid]
        
        Write-Log -Message "Attempting SID translation (attempt $attempt/$script:MaxRetries): $Sid" -Color "DarkGray" -NoConsole
        
        $user = Get-ADUser -Identity $Sid -Properties SamAccountName -ErrorAction Stop
        if ($user) {
            $name = $user.SamAccountName
            $script:SidCache[$Sid] = $name
            Write-Log -Message "Successfully resolved SID on attempt ${attempt}: ${Sid} -> ${name}" -Color "Green" -NoConsole
            return $name
        }
    }
    catch {
        $errorDetails = @{
            SID = $Sid
            Attempt = $attempt
            ErrorType = $_.Exception.GetType().Name
            ErrorMessage = $_.Exception.Message
            InnerError = if ($_.Exception.InnerException) { $_.Exception.InnerException.Message } else { "None" }
        }

        $logMessage = @"
SID Translation Error:
  SID: $($errorDetails.SID)
  Attempt: $($errorDetails.Attempt) of $script:MaxRetries
  Error Type: $($errorDetails.ErrorType)
  Message: $($errorDetails.ErrorMessage)
  Inner Error: $($errorDetails.InnerError)
"@

        Write-Log -Message $logMessage -Color "Yellow"

        if ($script:SidTranslationAttempts[$Sid] -lt $script:MaxRetries) {
            Write-Log -Message "Retrying in $script:RetryDelay seconds..." -Color "DarkGray"
            Start-Sleep -Seconds $script:RetryDelay
            return Convert-SidToName -Sid $Sid
        }
    }

    $script:FailedSids.Add($Sid)
    return $Sid
}

# Initialize variables
$script:SidCache = @{}
$script:FailedSids = [System.Collections.Generic.HashSet[string]]::new()
$script:SuppressedSids = [System.Collections.Generic.List[string]]::new()

# Add suppressed SIDs individually
$script:SuppressedSids.Add('S-1-5-21-3715258189-2875184700-594828381-500')
$script:SuppressedSids.Add('S-1-5-21-1787995930-3758959370-1315816792-13767')
$script:SuppressedSids.Add('S-1-5-21-1787995930-3758959370-1315816792-13821')
$script:SuppressedSids.Add('S-1-5-21-1787995930-3758959370-1315816792-17638')

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

            # Process the $permissionData object
            Write-Log -Message "Analyzing: $Path" -NoConsole
            Write-Log -Message "Owner: $($permissionData.Owner)" -NoConsole

            # Iterate through the access rules
            foreach ($accessRule in $permissionData.Access) {
                $identity = $accessRule.IdentityReference.Value
                if ($identity -match '^S-1-') {
                    $identity = Convert-SidToName -Sid $identity
                }
                Write-Log -Message "  Identity: $identity" -NoConsole
                Write-Log -Message "  Rights: $($accessRule.FileSystemRights)" -NoConsole
                Write-Log -Message "  Type: $($accessRule.AccessControlType)" -NoConsole
            }

            # Update unique permissions count
            if (-not $SkipUniquenessCounting) {
                $permissionHash = ($permissionData.Access | ForEach-Object { "$($_.IdentityReference)|$($_.FileSystemRights)|$($_.AccessControlType)" }) -join ';'
                if (-not $script:UniquePermissions.ContainsKey($permissionHash)) {
                    $script:UniquePermissions[$permissionHash] = @($Path)
                } else {
                    $script:UniquePermissions[$permissionHash] += $Path
                }
            }

        } else {
            Write-Log -Message "Could not retrieve permissions for $Path" -Color "Yellow"
        }
    }
    catch {
        Write-Log -Message "Error processing folder ${Path}: $($_.Exception.Message)" -Color "Red"
    }
    finally {
        Write-Progress -Activity "Analyzing Folder Permissions" -Status "Processing: $Path" -PercentComplete (($CurrentCount / $TotalCount) * 100)
    }
}

# Function to process folders recursively
function Invoke-FolderRecursively {
    param (
        [string]$Path,
        [int]$CurrentDepth = 0
    )

    $script:TotalFolders++
    $script:ProcessedFolders++

    Invoke-FolderProcessing -Path $Path -CurrentCount $script:ProcessedFolders -TotalCount $script:TotalFolders

    if ($MaxDepth -eq 0 -or $CurrentDepth -lt $MaxDepth) {
        Get-ChildItem -Path $Path -Directory | ForEach-Object {
            Invoke-FolderRecursively -Path $_.FullName -CurrentDepth ($CurrentDepth + 1)
        }
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
    Write-Log -Message "Starting folder permission analysis for $FolderPath" -Color "Cyan"
    Write-Log -Message "Detailed analysis will be written to: $script:DetailedLogFile" -Color "Yellow"

    # Process folders
    Invoke-FolderRecursively -Path $FolderPath

    # Calculate elapsed time
    $script:EndTime = Get-Date
    $script:ElapsedTime = $script:EndTime - $script:StartTime

    # Write summary to console collection
    Write-Log -Message "`nAnalysis Complete" -Color "Green"
    Write-Log -Message "Total folders processed: $($script:ProcessedFolders)" -Color "Cyan"
    Write-Log -Message "Unique permission sets: $($script:UniquePermissions.Count)" -Color "Cyan"
    Write-Log -Message "Elapsed time: $($script:ElapsedTime.ToString())" -Color "Cyan"

    # First, save console output to file
    $script:ConsoleOutputCollection | Out-File -FilePath $script:ConsoleLogFile -Encoding utf8
    
    # Stop transcript before final console writes
    try {
        Stop-Transcript
    }
    catch {
        Write-Warning "Error stopping transcript: $_"
    }

    # Now write to console directly
    $script:ConsoleOutputCollection | ForEach-Object { Write-Host $_ }
    
    # Clear the collection
    $script:ConsoleOutputCollection.Clear()
}
catch {
    Write-Error "An error occurred: $_"
    Write-Error $_.ScriptStackTrace
}
finally {
    Write-Progress -Activity "Analyzing Folder Permissions" -Completed
    
    # Ensure console output is saved if not already done
    if ($script:ConsoleOutputCollection.Count -gt 0) {
        $script:ConsoleOutputCollection | Out-File -FilePath $script:ConsoleLogFile -Encoding utf8 -Append
        $script:ConsoleOutputCollection.Clear()
    }
    
    # Final cleanup
    Write-Host "Script execution completed. See $script:DetailedLogFile for full details." -ForegroundColor Green
    
    # Stop transcript if still running
    if ($Host.Name -eq 'ConsoleHost') {
        try {
            Stop-Transcript -ErrorAction SilentlyContinue
        }
        catch {
            # Ignore any transcript errors in finally block
        }
    }
}
