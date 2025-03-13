# =============================================================================
# Script: Get-NTFSFolderPermissions.ps1
# Created: 2025-03-06 21:06:43 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-13 01:47:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2.0
# Additional Info: Added retry logic and detailed error reporting for SID translation
# =============================================================================

<#
.SYNOPSIS
    Advanced NTFS permission analyzer with directory grouping and inheritance tracking.

.DESCRIPTION
    Comprehensive NTFS permission analysis tool that provides detailed access control information
    for specified folders and their subfolders, with intelligent grouping of directories that
    share identical permissions and inheritance states.

    Key Features:
    - Groups directories with identical permissions
    - Tracks and displays permission inheritance status
    - Parallel folder processing with configurable thread limits
    - Hierarchical or grouped permission display modes
    - SID to name resolution with caching
    - Active Directory integration for accurate identity resolution
    - Performance optimizations for large directory structures

    Dependencies:
    - Windows PowerShell 5.1 or later
    - Active Directory PowerShell module (auto-loaded if available)
    - Read access to target folders
    - .NET Framework 4.5 or later

.PARAMETER FolderPath
    The root folder path to analyze for NTFS permissions.
    Type: String
    Required: True
    Example: "C:\Data" or "\\server\share"

.PARAMETER MaxThreads
    Maximum number of concurrent processing threads.
    Type: Integer
    Default: 10
    Required: False

.PARAMETER MaxDepth
    Maximum subfolder depth to traverse (0 = unlimited).
    Type: Integer
    Default: 0
    Required: False

.PARAMETER SkipUniquenessCounting
    Bypasses permission uniqueness analysis for performance optimization.
    Type: Switch
    Default: False
    Required: False

.PARAMETER SkipADResolution
    Disables Active Directory SID resolution.
    Type: Switch
    Default: False
    Required: False

.PARAMETER EnableSIDDiagnostics
    Enables detailed logging of SID resolution attempts.
    Type: Boolean
    Default: True
    Required: False

.PARAMETER ViewMode
    Determines how permissions are displayed in the output.
    Type: String
    Valid Values: "Hierarchy", "Group"
    Default: "Hierarchy"
    Required: False

.EXAMPLE
    .\Get-NTFSFolderPermissions.ps1 -FolderPath "C:\Data"
    Analyzes permissions for C:\Data and all subfolders using default settings.

.EXAMPLE
    .\Get-NTFSFolderPermissions.ps1 -FolderPath "\\server\share" -MaxThreads 20 -ViewMode "Group"
    Analyzes a network share using 20 threads and groups identical permissions.

.EXAMPLE
    .\Get-NTFSFolderPermissions.ps1 -FolderPath "C:\Users" -MaxDepth 2 -SkipADResolution
    Analyzes permissions up to 2 levels deep without AD resolution.
#>

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
    d function for sanitizing path for filename
    # Replace invalid filename characters and common separators
    $safe = $Path -replace '[\\\/\:\*\?\"\<\>\|]', '_'
    # Replace multiple underscores with single underscore
    $safe = $safe -replace '_{2,}', '_'rs and common separators
    # Trim underscores from ends/\:\*\?\"\<\>\|]', '_'
    $safe = $safe.Trim('_')rscores with single underscore
    # Limit length to prevent extremely long filenames
    if ($safe.Length -gt 50) {ds
        $safe = $safe.Substring(0, 47) + '...'
    } Limit length to prevent extremely long filenames
    return $safength -gt 50) {
}       $safe = $safe.Substring(0, 47) + '...'
    }
# Modify log file path creation
$systemName = [System.Environment]::MachineName
$safeFolderPath = Get-SafeFilename -Path $FolderPath
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$script:LogFile = Join-Path ([System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Path)) `
    "NTFSPermissions_${systemName}_${safeFolderPath}_${timestamp}.log"
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
# Function to write log messagesstem.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Path)) `
function Write-Log {_${systemName}_${safeFolderPath}_${timestamp}.log"
    param (
        [string]$Message,essages
        [string]$Color = "White",
        [switch]$NoConsole
    )   [string]$Message,
        [string]$Color = "White",
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] $Message"
    
    # Always write to log filemat "yyyy-MM-dd HH:mm:ss"
    [System.IO.File]::AppendAllText($script:LogFile, "$logMessage`n")
    
    # Write to console if not suppressed
    if (-not $NoConsole) {ndAllText($script:LogFile, "$logMessage`n")
        Write-Host $Message -ForegroundColor $Color
    } Write to console if not suppressed
}   if (-not $NoConsole) {
        Write-Host $Message -ForegroundColor $Color
# Function to display progress bar
function Write-ProgressBar {
    param (
        [int]$Current,progress bar
        [int]$Total,essBar {
        [string]$Activity,
        [string]$Status
    )   [int]$Total,
    $percentComplete = [math]::Round(($Current / $Total) * 100, 2)
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $percentComplete
}   )
    $percentComplete = [math]::Round(($Current / $Total) * 100, 2)
# Modify Get-FolderPermissions to check parent permissionsrcentComplete $percentComplete
function Get-FolderPermissions {
    param (
        [string]$Folderissions to check parent permissions
    )ion Get-FolderPermissions {
    param (
    try {string]$Folder
        $acl = Get-Acl -Path $Folder
        $owner = $acl.Owner
        $access = $acl.Access
        $isInherited = $acl.AreAccessRulesProtected -eq $false
        $owner = $acl.Owner
        $parentPath = Split-Path -Path $Folder -Parent
        $permissionHash = ($acl.Access | ForEach-Object { alse
            "$($_.IdentityReference)|$($_.FileSystemRights)|$($_.AccessControlType)|$($_.IsInherited)" 
        }) -join ';'= Split-Path -Path $Folder -Parent
        $permissionHash = ($acl.Access | ForEach-Object { 
        # Check if permissions match parentileSystemRights)|$($_.AccessControlType)|$($_.IsInherited)" 
        $matchesParent = $false
        if ($parentPath -and $script:ParentPermissions.ContainsKey($parentPath)) {
            $matchesParent = $script:ParentPermissions[$parentPath] -eq $permissionHash
        }matchesParent = $false
        if ($parentPath -and $script:ParentPermissions.ContainsKey($parentPath)) {
        $permissionData = [PSCustomObject]@{ermissions[$parentPath] -eq $permissionHash
            Folder = $Folder
            Owner  = $owner
            Access = $accessSCustomObject]@{
            PermissionHash = $permissionHash
            IsInherited = $isInherited
            MatchesParent = $matchesParent
        }   PermissionHash = $permissionHash
            IsInherited = $isInherited
        # Store current folder's permissions for child comparison
        $script:ParentPermissions[$Folder] = $permissionHash

        return $permissionData's permissions for child comparison
    }   $script:ParentPermissions[$Folder] = $permissionHash
    catch {
        Write-Log -Message "Error getting permissions for ${Folder}: $($_.Exception.Message)" -Color "Red"
        return $null
    }atch {
}       Write-Log -Message "Error getting permissions for ${Folder}: $($_.Exception.Message)" -Color "Red"
        return $null
# Function to resolve SID to name
function Resolve-SIDToName {
    param (
        [string]$SIDe SID to name
    )ion Resolve-SIDToName {
    param (
    if ($script:SIDCache.ContainsKey($SID)) {
        return $script:SIDCache[$SID]
    }
    if ($script:SIDCache.ContainsKey($SID)) {
    try {eturn $script:SIDCache[$SID]
        $objSID = New-Object System.Security.Principal.SecurityIdentifier($SID)
        $objName = $objSID.Translate([System.Security.Principal.NTAccount])
        $name = $objName.Value
        $script:SIDCache[$SID] = $namecurity.Principal.SecurityIdentifier($SID)
        return $nameobjSID.Translate([System.Security.Principal.NTAccount])
    }   $name = $objName.Value
    catch {ript:SIDCache[$SID] = $name
        if ($EnableSIDDiagnostics) {
            $script:ADResolutionErrors[$SID] = $_.Exception.Message
        } {
        return $SIDSIDDiagnostics) {
    }       $script:ADResolutionErrors[$SID] = $_.Exception.Message
}       }
        return $SID
# Add function to translate SIDs using AD lookup
function Convert-SidToName {
    param(
        [Parameter(Mandatory=$true)]ng AD lookup
        [string]$SidToName {
    )aram(
        [Parameter(Mandatory=$true)]
    # Check cache first
    if ($script:SidCache.ContainsKey($Sid)) {
        return $script:SidCache[$Sid]
    } Check cache first
    if ($script:SidCache.ContainsKey($Sid)) {
    # Check failed SIDs - suppress output for specific SIDs
    if ($script:FailedSids.Contains($Sid)) {
        if (-not ($script:SuppressedSids -contains $Sid)) {
            Write-Host "Skipping previously failed SID: $Sid" -ForegroundColor DarkGray
        }script:SidTranslationAttempts[$Sid] -ge $script:MaxRetries) {
        return $Sidscript:SuppressedSids -contains $Sid)) {
    }       $errorMsg = "Maximum retry attempts ($script:MaxRetries) reached for SID: $Sid"
            Write-Log -Message $errorMsg -Color "Yellow"
    try {
        # Try to get user by SID
        $user = Get-ADUser -Identity $Sid -Properties SamAccountName -ErrorAction Stop
        if ($user) {
            $name = $user.SamAccountNameexists
            $script:SidCache[$Sid] = $names.ContainsKey($Sid)) {
            return $namelationAttempts[$Sid] = 0
        }
    }
    catch {t = 0
        # Add to failed SIDs cache and return original SID
        $script:FailedSids.Add($Sid)
        return $SidTranslationAttempts[$Sid]++
    }
}       try {
            $user = Get-ADUser -Identity $Sid -Properties SamAccountName -ErrorAction Stop
# Initialize variables {
$script:SidCache = @{}= $user.SamAccountName
$script:FailedSids = New-Object System.Collections.Generic.HashSet[string]
$script:SuppressedSids = @(
    'S-1-5-21-3715258189-2875184700-594828381-500',
    'S-1-5-21-1787995930-3758959370-1315816792-13767',esolved SID on attempt $attempt: $Sid -> $name" -Color "Green"
    'S-1-5-21-1787995930-3758959370-1315816792-13821',
    'S-1-5-21-1787995930-3758959370-1315816792-17638'
)               return $name
            }
# Function to process each folder
function Invoke-FolderProcessing {
    param(  $errorDetails = @{
        [string]$Path,$Sid
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

            if ($attempt -lt $script:MaxRetries) {
                Write-Log -Message "Retrying in $script:RetryDelay seconds..." -Color "DarkGray"
                Start-Sleep -Seconds $script:RetryDelay
            }
        }
    }

    # Add to failed SIDs cache after all retries exhausted
    $script:FailedSids.Add($Sid)
    return $Sid
}

# Initialize variables
$script:SidCache = @{}
$script:FailedSids = New-Object System.Collections.Generic.HashSet[string]
$script:SuppressedSids = @(
    'S-1-5-21-3715258189-2875184700-594828381-500',
    'S-1-5-21-1787995930-3758959370-1315816792-13767',
    'S-1-5-21-1787995930-3758959370-1315816792-13821',
    'S-1-5-21-1787995930-3758959370-1315816792-17638'
)

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
    Write-Log -Message "Detailed analysis will be written to: $script:LogFile" -Color "Yellow"

    # Process folders
    Invoke-FolderRecursively -Path $FolderPath

    # Calculate elapsed time
    $script:EndTime = Get-Date
    $script:ElapsedTime = $script:EndTime - $script:StartTime

    # Display summary
    Write-Log -Message "`nAnalysis Complete" -Color "Green"
    Write-Log -Message "Total folders processed: $($script:ProcessedFolders)" -Color "Cyan"
    Write-Log -Message "Unique permission sets: $($script:UniquePermissions.Count)" -Color "Cyan"
    Write-Log -Message "Elapsed time: $($script:ElapsedTime.ToString())" -Color "Cyan"

    # Display results based on view mode
    if ($ViewMode -eq "Hierarchy") {
        Write-Log -Message "`nFolder Access Permissions:" -Color "Yellow"
        $sortedFolders = $script:FolderPermissions.Keys | Sort-Object

        foreach ($folder in $sortedFolders) {
            $data = $script:FolderPermissions[$folder]

            # Skip folders with identical permissions as parent
            if ($data.MatchesParent) {
                continue
            }

            $inheritanceStatus = if ($data.IsInherited) { "(Inherits parent permissions)" } else { "(Custom permissions)" }
        $script:ADResolutionErrors.GetEnumerator() | ForEach-Object {
            Write-Log -Message "SID: $($_.Key)" -Color "White"olor "Cyan"
            Write-Log -Message "Error: $($_.Value)" -Color "Red"
        }   # Find all descendants with identical permissions
    }       $identicalDescendants = $script:PermissionGroups[$data.PermissionHash].Folders | 
}               Where-Object { 
catch {             $_ -ne $folder -and 
    Write-Log -Message "An error occurred: $($_.Exception.Message)" -Color "Red"
}                   $script:FolderPermissions[$_].MatchesParent 
finally {       }
    Write-Progress -Activity "Analyzing Folder Permissions" -Completed
    Write-Log -Message "`nScript execution completed. See $script:LogFile for full details." -Color "Green"
}               Write-Log -Message "Subfolders with same permissions:" -Color "DarkGray"
                foreach ($descendant in ($identicalDescendants | Sort-Object)) {
                    $relativePath = $descendant.Substring($folder.Length + 1)
                    Write-Log -Message "  • $relativePath" -Color "DarkGray"
                }
            }

            Write-Log -Message "Access Rights:" -Color "White"
            $data.Access | ForEach-Object {
                $permission = Get-HumanReadablePermissions -Rights $_.FileSystemRights
                $accessType = if ($_.AccessControlType -eq 'Allow') { '✓' } else { '✗' }
                Write-Log -Message "  $accessType $($_.IdentityReference) - $permission" -Color "Gray"
            }
            Write-Log -Message "-" * 80 -Color "DarkGray"
        }
    }
    elseif ($ViewMode -eq "Group") {
        Write-Log -Message "`nPermission Groups:" -Color "Yellow"
        $script:UniquePermissions.GetEnumerator() | Sort-Object { $_.Value.Count } -Descending | ForEach-Object {
            $permissionSet = $_.Key -split ';' | ForEach-Object { $_ -replace '\|', ' : ' }
            Write-Log -Message "Group with $($_.Value.Count) folder(s):" -Color "White"
            $permissionSet | ForEach-Object { Write-Log -Message "  $_" -Color "Gray" }
            Write-Log -Message "Folders:" -Color "White"
            $_.Value | ForEach-Object { Write-Log -Message "  $_" -Color "Gray" }
            Write-Log -Message "" -Color "White"
        }
    }

    # Display SID resolution errors if any
    if ($EnableSIDDiagnostics -and $script:ADResolutionErrors.Count -gt 0) {
        Write-Log -Message "`nSID Resolution Errors:" -Color "Yellow"
        $script:ADResolutionErrors.GetEnumerator() | ForEach-Object {
            Write-Log -Message "SID: $($_.Key)" -Color "White"
            Write-Log -Message "Error: $($_.Value)" -Color "Red"
        }
    }
}
catch {
    Write-Log -Message "An error occurred: $($_.Exception.Message)" -Color "Red"
}
finally {
    Write-Progress -Activity "Analyzing Folder Permissions" -Completed
    Write-Log -Message "`nScript execution completed. See $script:LogFile for full details." -Color "Green"
}
