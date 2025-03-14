# =============================================================================
# Script: Get-NTFSFolderPermissions.ps1
# Created: 2025-02-07 21:21:53 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-14 22:30:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.9.0
# Additional Info: Removed .log functionality, keeping only transcript and debug.log
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
    [bool]$EnableSIDDiagnostics = $true
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

# Add function for sanitizing path for filename
function Get-SafeFilename {
    <#
    .SYNOPSIS
        Sanitizes a file path for safe file name creation.
    .DESCRIPTION
        Removes invalid characters and converts spaces to underscores.
        Truncates names longer than 50 characters.
        Returns sanitized path suitable for file naming.
    .PARAMETER Path
        The file path to sanitize.
    .EXAMPLE
        Get-SafeFilename -Path "C:\Program Files\My App"
        Returns: C_Program_Files_My_App
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )
    
    try {
        # Remove invalid characters
        $safeName = $Path -replace '[\\/:*?"<>|]', '_'
        
        # Replace multiple spaces/special chars with single underscore
        $safeName = $safeName -replace '[\s\p{P}]+', '_'
        
        # Remove leading/trailing underscores
        $safeName = $safeName.Trim('_')
        
        # Ensure length is reasonable
        if ($safeName.Length -gt 50) {
            $safeName = $safeName.Substring(0, 47) + '...'
        }
        
        return $safeName
    }
    catch {
        throw $_
    }
}

# Consolidated log initialization - single approach
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$computerName = $env:COMPUTERNAME
$safePath = Get-SafeFilename -Path $FolderPath
$logBase = Join-Path $PSScriptRoot "NTFSPermissions_${computerName}_${safePath}_${timestamp}"
$script:DebugLogFile = "${logBase}_debug.log"

# Function to create a standardized log header
function New-LogHeader {
    return @"
# =============================================================================
# NTFS Permissions Debug Log
# Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC")
# System: $computerName
# Analysis Path: $FolderPath
# Version: 1.9.0
# =============================================================================

"@
}

# Initialize debug log with proper header
Set-Content -Path $script:DebugLogFile -Value (New-LogHeader)

# Define Write-Log function before any usage
function Write-Log {
    param (
        [string]$Message,
        [string]$Color = "White",
        [switch]$NoConsole,
        [switch]$Debug
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff UTC"
    $callingFunction = (Get-PSCallStack)[1].FunctionName
    if ($callingFunction -eq "<ScriptBlock>") { $callingFunction = "MainScript" }
    
    $logEntry = "[TIMESTAMP: $timestamp]`n[FUNCTION: $callingFunction]`n[MESSAGE: $Message]`n----------------------------------------`n"
    
    # Write to debug log
    Add-Content -Path $script:DebugLogFile -Value $logEntry
    
    if (-not $NoConsole) {
        Write-Host $Message -ForegroundColor $Color
    }
}

# Import required modules
Import-Module ActiveDirectory -ErrorAction SilentlyContinue

# Initialize transcript - done only once
if ($Host.Name -eq 'ConsoleHost' -and -not $script:TranscriptStarted) {
    $transcriptPath = Join-Path $PSScriptRoot "NTFSPermissions_${timestamp}.transcript"
    try {
        Start-Transcript -Path $transcriptPath -Force
        $script:TranscriptStarted = $true
        Write-Log "Transcript started at $transcriptPath" -Debug
    }
    catch {
        Write-Warning "Could not start transcript: $_"
    }
}

# Initialize well-known SIDs
function Initialize-WellKnownSIDs {
    # Get Administrator SID using WMI
    $AdminSID = (Get-WmiObject Win32_UserAccount -Filter "Name='Administrator'" -ErrorAction SilentlyContinue).SID
    Write-Log -Message "The Administrator SID is: $AdminSID" -Color "White"

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
        "Administrator" = if ($AdminSID) { $AdminSID } else { "S-1-5-21-domain-500" }
        "Administrators" = "S-1-5-32-544"
        "Users" = "S-1-5-32-545"
        "Guests" = "S-1-5-32-546"
    }
    Write-Log -Message "Initialized well-known SIDs collection" -Color "DarkGray" -NoConsole
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
        Write-Log -Message "Skipping suppressed SID: $Sid" -Color "DarkGray" -NoConsole
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
        
        Write-Log -Message "Attempting SID translation (attempt $attempt of $script:MaxRetries): $Sid" -Color "DarkGray" -NoConsole -Debug
        
        # Try to resolve using .NET first
        try {
            $objSID = New-Object System.Security.Principal.SecurityIdentifier($Sid)
            $objName = $objSID.Translate([System.Security.Principal.NTAccount])
            $name = $objName.Value
            $script:SidCache[$Sid] = $name
            Write-Log -Message "Successfully resolved SID on attempt ${attempt}: ${Sid} -> ${name}" -Color "Green" -NoConsole -Debug
            return $name
        }
        catch {
            # Fall back to AD lookup if .NET translation fails
            Write-Log -Message ".NET translation failed on attempt ${attempt}, trying AD lookup for SID: $Sid" -Color "DarkGray" -NoConsole -Debug
            if (-not $SkipADResolution) {
                try {
                    $user = Get-ADUser -Identity $Sid -Properties SamAccountName -ErrorAction Stop
                    if ($user) {
                        $name = $user.SamAccountName
                        $script:SidCache[$Sid] = $name
                        Write-Log -Message "Successfully resolved SID via AD on attempt ${attempt}: ${Sid} -> ${name}" -Color "Green" -NoConsole -Debug
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
                            Write-Log -Message "Successfully resolved SID via AD (group) on attempt ${attempt}: ${Sid} -> ${name}" -Color "Green" -NoConsole -Debug
                            return $name
                        }
                    }
                    catch {
                        Write-Log -Message "AD resolution failed for SID: $Sid" -Color "Yellow" -NoConsole -Debug
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
        Write-Log -Message "SID translation failed on attempt ${attempt}: ${Sid}" -Color "Yellow" -NoConsole -Debug
        if ($script:SidTranslationAttempts[$Sid] -lt $script:MaxRetries) {
            Write-Log -Message "Retrying in $script:RetryDelay seconds (attempt ${attempt}/${script:MaxRetries})..." -Color "DarkGray" -NoConsole -Debug
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
        $permissionHash = Get-PermissionHash -AccessRules $acl.Access -IncludeInheritance
        
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
            Write-Log -Message "Could not retrieve permissions for $Path" -Color "Yellow"
        }
    }
    catch {
        Write-Log -Message "Error processing folder ${Path}: $($_.Exception.Message)" -Color "Red"
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
    Initialize-WellKnownSIDs
    
    # Output initial messages
    Write-Log -Message "Starting folder permission analysis for $FolderPath" -Color "Cyan"
    Write-Log -Message "Debug information will be written to: $script:DebugLogFile" -Color "Yellow"
    
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

    # Display results in hierarchy mode
    Write-Log -Message "`nFolder Access Permissions:" -Color "Yellow"
    $sortedFolders = $script:FolderPermissions.Keys | Sort-Object

    foreach ($folder in $sortedFolders) {
        $data = $script:FolderPermissions[$folder]

        # Skip folders with identical permissions as parent
        if ($data.MatchesParent) {
            continue
        }

        $inheritanceStatus = if ($data.IsInherited) { "(Inherits parent permissions)" } else { "(Custom permissions)" }
        Write-Log -Message "$folder $inheritanceStatus" -Color "Cyan"

        # Find all descendants with identical permissions
        $identicalDescendants = @($script:PermissionGroups[$data.PermissionHash].Folders | 
            Where-Object { 
                $_ -and 
                $_ -ne $folder -and 
                $script:FolderPermissions[$_].MatchesParent 
            })

        if ($identicalDescendants -and $identicalDescendants.Count -gt 0) {
            Write-Log -Message "Subfolders with same permissions:" -Color "DarkGray"
            foreach ($descendant in ($identicalDescendants | Sort-Object)) {
                if ($descendant -and $folder) {
                    $relativePath = $descendant.Substring($folder.Length + 1)
                    Write-Log -Message "  • $relativePath" -Color "DarkGray"
                }
            }
        }

        Write-Log -Message "Access Rights:" -Color "White"
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
                Write-Log -Message "  $accessType $identity - $permission" -Color "Gray"
            }
        }
        
        Write-Log -Message "-" * 80 -Color "DarkGray"
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
        Write-Warning "Error during cleanup: $_"
    }
    finally {
        # Properly clear important variables
        Remove-Variable -Name SidCache -Scope Script -ErrorAction SilentlyContinue
        Remove-Variable -Name FailedSids -Scope Script -ErrorAction SilentlyContinue
        Remove-Variable -Name SuppressedSids -Scope Script -ErrorAction SilentlyContinue
    }
}