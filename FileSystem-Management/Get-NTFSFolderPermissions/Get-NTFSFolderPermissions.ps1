# =============================================================================
# Script: Get-NTFSFolderPermissions.ps1
# Created: 2025-02-07 21:21:53 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-17 00:55:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.12.0
# Additional Info: Added owner display and included owner in permission comparisons
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
$script:TranscriptFile = "${logBase}_transcript.log"

# Function to create a standardized log header with enhanced metadata
function New-LogHeader {
    # Get execution context information
    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")
    $os = Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object -ExpandProperty Caption
    
    @"
# =============================================================================
# NTFS Permissions Debug Log
# Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC")
# System: $computerName
# OS Version: $os
# PowerShell Version: $($PSVersionTable.PSVersion)
# Executed By: $currentUser
# Admin Privileges: $isAdmin
# Analysis Path: $FolderPath
# Max Threads: $MaxThreads
# Max Depth: $MaxDepth
# Skip AD Resolution: $SkipADResolution
# Skip Uniqueness Counting: $SkipUniquenessCounting
# Enable SID Diagnostics: $EnableSIDDiagnostics
# Script Version: 1.12.0
# =============================================================================

"@
}

# Initialize debug log with proper header
Set-Content -Path $script:DebugLogFile -Value (New-LogHeader)

# Define Write-Log function with standardized PowerShell format and enhanced metrics
# Define Write-Log function with standardized PowerShell format and enhanced metrics
function Write-Log {
    param (
        [string]$Message,
        [string]$Color = "White",
        [switch]$NoConsole,
        [switch]$Debug,
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'DEBUG', 'SUCCESS', 'METRIC', 'VERBOSE')]
        [string]$Level = $(if ($Debug) { 'DEBUG' } else { 'INFO' }),
        [string]$Category = "",
        [int]$Indent = 0
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $callStack = Get-PSCallStack
    $callingFunction = $callStack[1].FunctionName
    $lineNumber = $callStack[1].ScriptLineNumber
    if ($callingFunction -eq "<ScriptBlock>") { $callingFunction = "MainScript" }
    
    # Calculate memory usage for metrics
    $memoryInfo = ""
    if ($Level -eq 'METRIC') {
        $process = Get-Process -Id $PID
        $memoryMB = [math]::Round($process.WorkingSet / 1MB, 2)
        $memoryInfo = "[Memory:${memoryMB}MB] "
    }
    
    # Add category for better filtering
    $categoryInfo = if ($Category) { "[$Category] " } else { "" }
    
    # Add indentation for hierarchical clarity
    $indentation = if ($Indent -gt 0) { " " * $Indent } else { "" }
    
    # Standard PowerShell log format with enhancements
    $logEntry = "$timestamp [$Level] [Thread:$([Threading.Thread]::CurrentThread.ManagedThreadId)] [$callingFunction`:$lineNumber] ${memoryInfo}${categoryInfo}${indentation}$Message"

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
            'METRIC' = 'Cyan'
            'VERBOSE' = 'DarkGray'
        }
        $messageColor = if ($levelColors.ContainsKey($Level)) { $levelColors[$Level] } else { $Color }
        Write-Host "$indentation$Message" -ForegroundColor $messageColor
    }
}

# Add function to record performance metrics during folder processing
function Write-PerformanceMetric {
    param(
        [string]$Operation,
        [datetime]$StartTime,
        [int]$ItemCount = 0
    )
    
    $endTime = Get-Date
    $duration = ($endTime - $StartTime).TotalMilliseconds
    $itemsPerSec = if ($ItemCount -gt 0 -and $duration -gt 0) { 
        [math]::Round(($ItemCount * 1000) / $duration, 2) 
    } else { 
        0 
    }
    
    $message = "$Operation completed in $([math]::Round($duration, 2))ms"
    if ($ItemCount -gt 0) {
        $message += " ($itemsPerSec items/sec)"
    }
    
    Write-Log -Message $message -Level 'METRIC' -Category 'Performance' -NoConsole
}

# Initialize well-known SIDs
function Initialize-WellKnownSIDs {
    # Get all Administrator SIDs from WMI
    $adminAccounts = @()
    $wmiAdminAccounts = Get-WmiObject Win32_UserAccount -Filter "Name='Administrator'" -ErrorAction SilentlyContinue

    if ($wmiAdminAccounts) {
        # Process each admin account and store structured information
        foreach ($account in $wmiAdminAccounts) {
            $domainType = if ($account.LocalAccount) { "Local" } else { "Domain" }
            $domain = if ($account.Domain) { $account.Domain } else { $env:COMPUTERNAME }

            # Try to get FQDN for domain accounts
            $fqdn = $domain
            if (-not $account.LocalAccount) {
                try {
                    $domainObj = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain(
                        (New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $domain))
                    )
                    $fqdn = $domainObj.Name
                }
                catch {
                    Write-Log -Message "Could not get FQDN for domain $domain" -Color "DarkGray" -NoConsole -Level 'DEBUG'
                }
            }

            $adminAccounts += [PSCustomObject]@{
                SID = $account.SID
                Domain = $domain
                FQDN = $fqdn
                DomainType = $domainType
                DisplayName = "$($account.SID) [${domainType}: $fqdn]"
            }
        }

        # Set the first SID for lookups
        if ($adminAccounts.Count -gt 0) {
            $script:AdminSID = $adminAccounts[0].SID
        }
    }
    else {
        # Fallback to default Administrator SID pattern
        $script:AdminSID = "S-1-5-21-domain-500"
        Write-Log -Message "Could not retrieve Administrator SIDs, using placeholder" -Color "Yellow" -Level 'WARNING'
    }

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
        "Administrator" = $script:AdminSID  # Use first SID for lookups
        "Administrators" = "S-1-5-32-544"
        "Users" = "S-1-5-32-545"
        "Guests" = "S-1-5-32-546"
    }
    Write-Log -Message "Initialized well-known SIDs collection with primary Administrator SID: $script:AdminSID" -Color "DarkGray" -NoConsole -Level 'DEBUG'

    # Return the administrator accounts collection
    return $adminAccounts
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

# Enhanced SID translation function
function Get-SIDTranslation {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SID
    )
    
    try {
        if ($script:SidCache.ContainsKey($SID)) {
            return $script:SidCache[$SID]
        }
        
        # Get domain SID prefix
        $domainSid = $SID.Split('-')[0..4] -join '-'
        Write-Log "Attempting to translate SID: $SID (Domain prefix: $domainSid)" -Level 'DEBUG'
        
        $objSID = New-Object System.Security.Principal.SecurityIdentifier($SID)
        $objUser = $objSID.Translate([System.Security.Principal.NTAccount])
        
        $script:SidCache[$SID] = $objUser.Value
        Write-Log "Successfully translated $SID to $($objUser.Value)" -Level 'DEBUG'
        return $objUser.Value
    }
    catch {
        Write-Log "Failed to translate SID $SID : $_" -Level 'WARNING'
        if (-not $script:FailedSids.ContainsKey($SID)) {
            $script:FailedSids[$SID] = $_
        }
        return $SID
    }
}

# Add this helper function near the top with other functions
function Test-AdministratorSID {
    param([string]$SID)
    
    if ([string]::IsNullOrEmpty($SID)) { return $false }
    
    # Check if it's a domain SID ending in -500 (Administrator)
    return $SID -match '^S-1-5-21-\d+-\d+-\d+-500$'
}

# Modify the existing ConvertTo-NTAccountOrSID function to handle Administrator SIDs
function ConvertTo-NTAccountOrSID {
    param (
        [Parameter(Mandatory=$true)]
        [string]$SID
    )
    
    try {
        # Check if it's an Administrator SID from any domain
        if (Test-AdministratorSID -SID $SID) {
            Write-Log -Message "Found Administrator SID from domain: $SID" -Level 'DEBUG'
            return "ADMINISTRATOR (Domain: $(($SID -split '-')[1..3] -join '-'))"
        }
        
        # Try to convert SID to NT account name
        if (-not [string]::IsNullOrEmpty($SID)) {
            $objSID = New-Object System.Security.Principal.SecurityIdentifier($SID)
            $objUser = $objSID.Translate([System.Security.Principal.NTAccount])
            return $objUser.Value
        }
        return $SID
    }
    catch {
        Write-Log -Message "Error translating SID $SID : $_" -Level 'ERROR'
        return $SID
    }
}

# Standardized function to generate permission hashes
function Get-PermissionHash {
    param (
        [Parameter(Mandatory=$true)]
        [object]$AccessRules,

        [Parameter(Mandatory=$true)]
        [string]$Owner,

        [Parameter(Mandatory=$false)]
        [bool]$IncludeInheritance = $true
    )

    $ownerPart = "OWNER:$Owner"
    if ($IncludeInheritance) {
        return "$ownerPart;" + ($AccessRules | ForEach-Object { 
            "$($_.IdentityReference)|$($_.FileSystemRights)|$($_.AccessControlType)|$($_.IsInherited)" 
        }) -join ';'
    }
    else {
        return "$ownerPart;" + ($AccessRules | ForEach-Object { 
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
        $permissionHash = Get-PermissionHash -AccessRules $acl.Access -Owner $owner -IncludeInheritance $true

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
# Modify Invoke-FolderProcessing to track detailed metrics
function Invoke-FolderProcessing {
    param(
        [string]$Path,
        [int]$CurrentCount,
        [int]$TotalCount
    )

    $startTime = Get-Date
    try {
        Write-Log -Message "Processing folder: $Path" -Level 'VERBOSE' -Category 'FolderProcess' -NoConsole

        # Get permissions for the folder
        $permissionData = Get-FolderPermissions -Folder $Path

        if ($permissionData) {
            # Track unique permission sets with owner
            $permissionHash = Get-PermissionHash -AccessRules $permissionData.Access -Owner $permissionData.Owner -IncludeInheritance $true
            
            if (-not $script:PermissionGroups.ContainsKey($permissionHash)) {
                $script:PermissionGroups[$permissionHash] = @{
                    Folders = @()
                    Permissions = $permissionData.Access
                    Owner = $permissionData.Owner
                    IsInherited = $permissionData.IsInherited
                    ParentPaths = @{}
                }
                Write-Log -Message "Found new permission set with owner (hash: $($permissionHash.Substring(0,20))...)" -Level 'DEBUG' -Category 'Permissions' -NoConsole
            }

            # Group by parent path for better organization
            $currentParent = Split-Path -Path $Path -Parent
            if (-not [string]::IsNullOrEmpty($currentParent)) {
                if (-not $script:PermissionGroups[$permissionHash].ParentPaths.ContainsKey($currentParent)) {
                    $script:PermissionGroups[$permissionHash].ParentPaths[$currentParent] = @()
                }
                $script:PermissionGroups[$permissionHash].ParentPaths[$currentParent] += $Path
            }

            $script:PermissionGroups[$permissionHash].Folders += $Path
            $script:FolderPermissions[$Path] = $permissionData

            # Log non-inherited permissions for security analysis
            if (-not $permissionData.IsInherited) {
                Write-Log -Message "Found explicit permissions on: $Path" -Level 'DEBUG' -Category 'Security' -NoConsole
            }

            # Update unique permissions count - only write to debug log
            if (-not $SkipUniquenessCounting) {
                $permissionHash = Get-PermissionHash -AccessRules $permissionData.Access -Owner $permissionData.Owner -IncludeInheritance:$false

                if (-not $script:UniquePermissions.ContainsKey($permissionHash)) {
                    $script:UniquePermissions[$permissionHash] = @($Path)
                } else {
                    $script:UniquePermissions[$permissionHash] += $Path
                }
            }
        } else {
            Write-Log -Message "Could not retrieve permissions for $Path" -Color "Yellow" -Level 'WARNING'
            Write-Log -Message "Could not retrieve permissions for $Path" -Color "Yellow" -Level 'WARNING' -Category 'AccessDenied'
        }
    }
    catch {
        Write-Log -Message "Error processing folder ${Path}: $($_.Exception.Message)" -Color "Red" -Level 'ERROR'
        Write-Log -Message "Error processing folder ${Path}: $($_.Exception.Message)" -Color "Red" -Level 'ERROR' -Category 'Exception'
        Write-Log -Message "Stack trace: $($_.ScriptStackTrace)" -Level 'DEBUG' -Category 'Exception' -NoConsole
    }
    finally {
        Write-PerformanceMetric -Operation "Folder processing for $Path" -StartTime $startTime
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
    # Start transcript first thing
    try {
        # Suppress default transcript message by redirecting to null
        Start-Transcript -Path $script:TranscriptFile -Force | Out-Null
        $script:TranscriptStarted = $true
        Write-Host "Initializing transcript at: $script:TranscriptFile" -ForegroundColor Cyan
    }
    catch {
        Write-Warning "Failed to start transcript: $_"
        $script:TranscriptStarted = $false
    }

    # Register ctrl+c handler
    $null = [Console]::TreatControlCAsInput = $true
    Register-ObjectEvent -InputObject ([Console]) -EventName CancelKeyPress -Action {
        $script:cancellationTokenSource.Cancel()
        Write-Log "Cancellation requested by user (Ctrl+C)" -Level 'WARNING' -Color "Yellow"
    } | Out-Null

    # Output initial messages in correct order
    Write-Log -Message "Debug information will be written to: $script:DebugLogFile" -Level 'DEBUG'
    Write-Log ""

    # Initialize SIDs and get Administrator accounts
    $adminAccounts = Initialize-WellKnownSIDs

    # Always ensure $adminAccounts is an array for consistent behavior
    if ($null -eq $adminAccounts) {
        $adminAccounts = @()
    }
    elseif ($adminAccounts -isnot [Array] -and $adminAccounts -isnot [System.Collections.ICollection]) {
        # Convert single item to array
        $adminAccounts = @($adminAccounts)
    }

    # Display warning about multiple accounts if needed
    if ($adminAccounts.Count -gt 1) {
        Write-Log -Message "Multiple Administrator accounts found ($($adminAccounts.Count))" -Color "Yellow" -Level 'WARNING'
    }

    # Display each Administrator SID individually
    if ($adminAccounts.Count -gt 0) {
        foreach ($admin in $adminAccounts) {
            $domainName = if ($admin.DomainType -eq "Local") { "LOCAL" } else { $admin.FQDN }
            Write-Log -Message "The Administrator SID for $domainName is $($admin.SID)" -Color "White" -Level 'INFO'
        }
    } else {
        Write-Log -Message "No Administrator accounts found. Using default SID patterns." -Color "Yellow" -Level 'WARNING'
    }

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

        # Skip folders with identical permissions as parent (including owner)
        if ($data.MatchesParent) {
            continue
        }

        $inheritanceStatus = if ($data.IsInherited) { "(Inherits parent permissions)" } else { "(Custom permissions)" }
        Write-Log -Message "$folder $inheritanceStatus" -Color "Cyan" -Level 'INFO'
        
        # Display owner information
        Write-Log -Message "Owner: $($data.Owner)" -Color "White" -Level 'INFO'

        # Find all descendants with identical permissions (including owner)
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
                    if ($descendant.Length -gt $folder.Length) {
                        $relativePath = $descendant.Substring($folder.Length + 1)
                        Write-Log -Message "  - $relativePath" -Color "DarkGray" -Level 'INFO'
                    } else {
                        # Handle case where descendant path is the same or shorter than folder path
                        Write-Log -Message "  - <root folder>" -Color "DarkGray" -Level 'INFO'
                    }
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
} # End of main try block
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
