# =============================================================================
# Script: Get-NTFSFolderPermissions.ps1
# Created: 2025-02-07 21:21:53 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-14 20:29:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.7.0
# Additional Info: Consolidated logging to two files for cleaner output
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
    [ValidateSet("Hierarchy", "Group")]
    [string]$ViewMode = "Hierarchy"
)

# Enable strict mode and error handling
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Script-level variables
$script:DetailedLogBuffer = $null
$script:TranscriptStarted = $false

function Initialize-LogBuffers {
    <#
    .SYNOPSIS
        Initializes logging buffers for the script.
    .DESCRIPTION
        Creates and configures StringBuilder objects for detailed logging.
        Sets up initial capacity to optimize memory usage.
    #>
    [CmdletBinding()]
    param()
    
    try {
        if ($null -eq $script:DetailedLogBuffer) {
            $script:DetailedLogBuffer = [System.Text.StringBuilder]::new(360192)  # 352KB
            Write-Log "Log buffers initialized successfully" -DetailedOnly
            return $true
        }
        return $false
    }
    catch {
        Write-Log "Failed to initialize log buffers: $_" -Color Red
        throw
    }
}

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

# Initialize well-known SIDs
$script:WellKnownSIDs = @{}

# Initialize script-level variables
$script:DetailedLogBuffer = [System.Text.StringBuilder]::new(360192)  # 352KB initial capacity
$script:TranscriptStarted = $false

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
    
    # Write to appropriate log file
    if ($Debug) {
        Add-Content -Path $script:DebugLogFile -Value $logEntry
    } else {
        Add-Content -Path $script:LogFile -Value $logEntry
    }
    
    if (-not $NoConsole) {
        Write-Host $Message -ForegroundColor $Color
    }
}

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
        Write-Log "Sanitizing path: $Path" -Debug
        
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
        
        Write-Log "Sanitized path result: $safeName" -Debug
        return $safeName
    }
    catch {
        Write-Log "Error in Get-SafeFilename: $_" -Color Red -Debug
        throw
    }
}

# Initialize logging with consolidated files
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$computerName = $env:COMPUTERNAME
$safePath = Get-SafeFilename -Path $FolderPath
$logBase = Join-Path $PSScriptRoot "NTFSPermissions_${computerName}_${safePath}_${timestamp}"
$script:LogFile = "${logBase}.log" 
$script:DebugLogFile = "${logBase}_debug.log"

# Remove old logging variables
Remove-Variable -Name TranscriptFile -ErrorAction SilentlyContinue
Remove-Variable -Name DetailedLogFile -ErrorAction SilentlyContinue
Remove-Variable -Name ConsoleLogFile -ErrorAction SilentlyContinue

# Initialize log files with proper headers
@"
# =============================================================================
# NTFS Permissions Analysis Log
# Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC")
# System: $computerName
# Analysis Path: $FolderPath
# Version: 1.7.0
# =============================================================================

"@ | Set-Content $script:LogFile

@"
# =============================================================================
# NTFS Permissions Debug Log
# Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC")
# System: $computerName
# Analysis Path: $FolderPath
# Version: 1.7.0
# =============================================================================

"@ | Set-Content $script:DebugLogFile

# Remove transcript-related code and variables
Remove-Variable -Name TranscriptStarted -Scope Script -ErrorAction SilentlyContinue

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

# Simplified log file creation - Fix the swapped log file names
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$systemName = $env:COMPUTERNAME
$safeFolderPath = Get-SafeFilename -Path $FolderPath

# Define log files with correct naming
$logBase = Join-Path $PSScriptRoot "NTFSPermissions_${systemName}_${safeFolderPath}_${timestamp}"
$script:DetailedLogFile = "${logBase}_detailed.log"
$script:ConsoleLogFile = "${logBase}_console.log"

# Initialize log files and create them with headers
@"
# =============================================================================
# Detailed Log File
# Created: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC")
# System: $systemName
# Analysis Path: $FolderPath
# =============================================================================

"@ | Out-File -FilePath $script:DetailedLogFile -Force

@"
# =============================================================================
# Console Log File
# Created: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC")
# System: $systemName
# Analysis Path: $FolderPath
# =============================================================================

"@ | Out-File -FilePath $script:ConsoleLogFile -Force

# Initialize console output collection for detailed log (swap from console)
$script:ConsoleOutputCollection = [System.Collections.ArrayList]@()

# Output initial messages (remove duplicate Write-Host calls)
Write-Log -Message "Starting folder permission analysis for $FolderPath" -Color "Cyan"
Write-Log -Message "Detailed analysis will be written to: $script:DetailedLogFile" -Color "Yellow"
Write-Log -Message "Console summary will be written to: $script:ConsoleLogFile" -Color "Yellow"

# Add transcript state tracking
$script:TranscriptStarted = $false

# Modified transcript start handling
try {
    if ($Host.Name -eq 'ConsoleHost') {
        $transcriptPath = Join-Path $PSScriptRoot "NTFSPermissions_$(Get-Date -Format 'yyyyMMdd_HHmmss').transcript"
        Start-Transcript -Path $transcriptPath -Force -ErrorAction Stop
        $script:TranscriptStarted = $true
        Write-Log "Transcript started at $transcriptPath" -Debug
    }
}
catch {
    Write-Warning "Could not start transcript: $_"
    # Continue execution even if transcript fails
}

# Initialize buffers
Initialize-LogBuffers

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

# Modify the Convert-SidToName function to properly handle suppressed SIDs
function Convert-SidToName {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Sid
    )
    
    # Check if SID is in suppressed list first
    if ($script:SuppressedSids -contains $Sid) {
        Write-Log -Message "Skipping suppressed SID: $Sid" -Color "DarkGray" -NoConsole
        return $Sid
    }
    
    # Check cache first
    if ($script:SidCache.ContainsKey($Sid)) {
        return $script:SidCache[$Sid]
    }
    
    # Check if it's a well-known SID
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
        $script:SidTranslationAttempts[$Sid]++
        $attempt = $script:SidTranslationAttempts[$Sid]
        
        Write-Log -Message "Attempting SID translation (attempt $attempt of $script:MaxRetries): $Sid" -Color "DarkGray" -NoConsole
        
        # Try to resolve using .NET first
        try {
            $objSID = New-Object System.Security.Principal.SecurityIdentifier($Sid)
            $objName = $objSID.Translate([System.Security.Principal.NTAccount])
            $name = $objName.Value
            $script:SidCache[$Sid] = $name
            Write-Log -Message "Successfully resolved SID on attempt ${attempt}: ${Sid} -> ${name}" -Color "Green" -NoConsole
            return $name
        }
        catch {
            # Fall back to AD lookup if .NET translation fails
            Write-Log -Message "NET translation failed on attempt ${attempt}, trying AD lookup for SID: $Sid" -Color "DarkGray" -NoConsole
            $user = Get-ADUser -Identity $Sid -Properties SamAccountName -ErrorAction Stop
            if ($user) {
                $name = $user.SamAccountName
                $script:SidCache[$Sid] = $name
                Write-Log -Message "Successfully resolved SID via AD on attempt ${attempt}: ${Sid} -> ${name}" -Color "Green" -NoConsole
                return $name
            }
        }
    }
    catch {
        Write-Log -Message "SID translation failed on attempt ${attempt}: ${Sid}" -Color "Yellow" -NoConsole
        if ($script:SidTranslationAttempts[$Sid] -lt $script:MaxRetries) {
            Write-Log -Message "Retrying in $script:RetryDelay seconds (attempt ${attempt}/${script:MaxRetries})..." -Color "DarkGray" -NoConsole
            Start-Sleep -Seconds $script:RetryDelay
            return Convert-SidToName -Sid $Sid
        }
        $script:FailedSids.Add($Sid)
    }
    
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
    Initialize-WellKnownSIDs
    # Remove duplicate message here since it's already handled above
    
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
            $data.Access | ForEach-Object {
                $permission = Get-HumanReadablePermissions -Rights $_.FileSystemRights
                $accessType = if ($_.AccessControlType -eq 'Allow') { '+' } else { '-' }
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

    # Generate console output and save to separate log
    $consoleOutput = @()
    $consoleOutput += "Analysis Complete"
    $consoleOutput += "Total folders processed: $($script:ProcessedFolders)"
    $consoleOutput += "Unique permission sets: $($script:UniquePermissions.Count)"
    $consoleOutput += "Elapsed time: $($script:ElapsedTime.ToString())"
    $consoleOutput += ""
    $consoleOutput += "Folder Access Permissions:"

    foreach ($group in $script:PermissionGroups.GetEnumerator()) {
        $firstFolder = $group.Value.Folders[0]
        $inherits = (Get-Acl $firstFolder).AreAccessRulesProtected -eq $false
        $consoleOutput += "$firstFolder $(if ($inherits) {'(Inherits parent permissions)'})"
        if ($group.Value.Owner) {
            $consoleOutput += "Owner: $($group.Value.Owner)"
        }
        
        if ($group.Value.Folders -and $group.Value.Folders.Count -gt 1) {
            $consoleOutput += "Subfolders with same permissions:"
            $group.Value.Folders | 
                Where-Object { $_ } | 
                Select-Object -Skip 1 | 
                ForEach-Object {
                    if ($_) {
                        $consoleOutput += "  - $_"
                    }
                }
        }
        
        $consoleOutput += "Access Rights:"
        $group.Value.Permissions | ForEach-Object {
            $consoleOutput += "  + $($_.IdentityReference) - $($_.FileSystemRights)"
        }
        $consoleOutput += "-" * 80
    }

    # Output to console and save to file
    $consoleOutput | Out-File -FilePath $script:ConsoleLogFile
    $consoleOutput | ForEach-Object { Write-Host $_ }

    Stop-Transcript
}
catch [System.Exception] {
    Write-Error "An error occurred: $_"
    Write-Error $_.ScriptStackTrace
}
# In the finally block, update to:
finally {
    Write-Progress -Activity "Analyzing Folder Permissions" -Completed
    
    try {
        # Final flush of log buffers
        if ($null -ne $script:DetailedLogBuffer -and $script:DetailedLogBuffer.Length -gt 0) {
            Add-Content -Path $script:DetailedLogFile -Value $script:DetailedLogBuffer.ToString() -ErrorAction Stop
        }
        
        Write-Host "Script execution completed. See $script:DetailedLogFile for full details." -ForegroundColor Green
        
        # Clean up transcript only if we started one
        if ($script:TranscriptStarted) {
            Stop-Transcript -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-Warning "Error during cleanup: $_"
    }
    finally {
        # Properly clear buffer references
        Remove-Variable -Name DetailedLogBuffer -Scope Script -ErrorAction SilentlyContinue
        Remove-Variable -Name ConsoleLogBuffer -Scope Script -ErrorAction SilentlyContinue
        Remove-Variable -Name BuffersInitialized -Scope Script -ErrorAction SilentlyContinue
    }
}
