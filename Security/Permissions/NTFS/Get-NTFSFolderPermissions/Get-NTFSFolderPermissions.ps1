# =============================================================================
# Script: Get-NTFSFolderPermissions.ps1
# Created: 2025-03-15 18:30:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-28 19:38:00 UTC
# Updated By: jdyer-nuvodia
# Version: 3.3.4
# Additional Info: Added property validation and null checks to Compare-PermissionSets
# =============================================================================

<#
.SYNOPSIS
Gets NTFS folder permissions for specified path.

.DESCRIPTION
Analyzes and reports NTFS permissions for specified folder path and its subfolders.
Consolidates output into two log files:
- Main log for permission details
- Debug log for troubleshooting information
Subfolders with identical permissions and owners as their parent are grouped together.

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

.EXAMPLE
.\Get-NTFSFolderPermissions.ps1 -StartPath "C:\Temp"
Analyzes permissions on C:\Temp and outputs to logs
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
    [int]$TimeoutMinutes = 120
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

# Function to get domain controllers and domain information
function Get-DomainControllers {
    try {
        # Try to get domain information using .NET first
        $domainInfo = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
        return @($domainInfo.DomainControllers | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                Domain = $domainInfo.Name
                Forest = $domainInfo.Forest.Name
                IsGlobalCatalog = $_.IsGlobalCatalog
            }
        })
    }
    catch {
        Write-Log -Message "Failed to get domain controllers using .NET: $_" -Level 'WARNING' -Color "Yellow"
        try {
            # Fallback to using AD cmdlets if available
            if (Get-Command Get-ADDomainController -ErrorAction SilentlyContinue) {
                return @(Get-ADDomainController -Filter * | ForEach-Object {
                    [PSCustomObject]@{
                        Name = $_.HostName
                        Domain = $_.Domain
                        Forest = $_.Forest
                        IsGlobalCatalog = $_.IsGlobalCatalog
                    }
                })
            }
        }
        catch {
            Write-Log -Message "Failed to get domain controllers using AD cmdlets: $_" -Level 'WARNING' -Color "Yellow"
        }
        
        # If both methods fail, return computer domain info
        try {
            $computerDomain = (Get-WmiObject Win32_ComputerSystem).Domain
            if ($computerDomain) {
                return @([PSCustomObject]@{
                    Name = $env:COMPUTERNAME
                    Domain = $computerDomain
                    Forest = $computerDomain
                    IsGlobalCatalog = $false
                })
            }
        }
        catch {
            Write-Log -Message "Failed to get computer domain info: $_" -Level 'WARNING' -Color "Yellow"
        }
    }
    
    # Return empty array if all methods fail
    return @()
}

# Function to count total folders recursively
function Get-TotalFolderCount {
    param (
        [Parameter(Mandatory=$true)]
        [string]$StartPath
    )
    
    try {
        $folderCount = 0
        $folders = @(Get-ChildItem -Path $StartPath -Directory -Force -ErrorAction Stop)
        $folderCount += $folders.Count
        
        foreach ($folder in $folders) {
            try {
                $folderCount += Get-TotalFolderCount -StartPath $folder.FullName
            }
            catch {
                Write-Log -Message "Error counting subfolders in $($folder.FullName): $_" -Level 'WARNING' -Color "Yellow"
            }
        }
        
        return $folderCount
    }
    catch {
        Write-Log -Message "Error counting folders in $StartPath : $_" -Level 'ERROR' -Color "Red"
        return 0
    }
}

# Function to process folders recursively
function Invoke-FolderRecursively {
    param (
        [Parameter(Mandatory=$true)]
        [string]$StartPath,
        
        [Parameter(Mandatory=$false)]
        [int]$CurrentDepth = 0
    )
    
    try {
        # Check for timeout or cancellation
        if ($script:cancellationTokenSource.Token.IsCancellationRequested -or 
            ((Get-Date) - $script:StartTime) -gt $script:processingTimeout) {
            return
        }

        # Get folder permissions
        $acl = Get-Acl -Path $StartPath -ErrorAction Stop
        $permissionHash = Get-PermissionHash -AccessRules $acl.Access -Owner $acl.Owner
        
        # Store permissions
        $script:FolderPermissions[$StartPath] = @{
            Owner = $acl.Owner
            Access = $acl.Access
            IsInherited = $true
            UniqueHash = $permissionHash
        }
        
        # Track unique permissions
        if (-not $SkipUniquenessCounting) {
            if (-not $script:UniquePermissions.ContainsKey($permissionHash)) {
                $script:UniquePermissions[$permissionHash] = @()
            }
            $script:UniquePermissions[$permissionHash] += $StartPath
        }
        
        # Update progress
        $script:ProcessedFolders++
        Write-ProgressStatus -Activity "Analyzing Folder Permissions" -Status $StartPath -Current $script:ProcessedFolders -Total $script:TotalFolders
        
        # Check depth limit
        if ($MaxDepth -gt 0 -and $CurrentDepth -ge $MaxDepth) {
            Write-Log -Message "Reached maximum depth ($MaxDepth) at: $StartPath" -Level 'DEBUG' -NoConsole
            return
        }
        
        # Process subfolders
        $folders = Get-ChildItem -Path $StartPath -Directory -Force -ErrorAction Stop
        foreach ($folder in $folders) {
            try {
                Invoke-FolderRecursively -StartPath $folder.FullName -CurrentDepth ($CurrentDepth + 1)
            }
            catch {
                Write-Log -Message "Error processing subfolder $($folder.FullName): $_" -Level 'WARNING' -Color "Yellow"
            }
        }
    }
    catch {
        Write-Log -Message "Error processing folder $StartPath : $_" -Level 'ERROR' -Color "Red"
    }
}

# Add function for sanitizing path for filename
function Get-SafeFilename {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]$StartPath
    )

    try {
        # Initial null/empty check
        if ([string]::IsNullOrEmpty($StartPath)) {
            throw "Path cannot be null or empty"
        }

        # Convert path separators to underscores
        $safeName = $StartPath.Replace('\', '_').Replace('/', '_')

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
        if (-not [string]::IsNullOrEmpty($StartPath)) {
            # Take last part of path if available
            $lastPart = $StartPath.Split('\')[-1]
            if (-not [string]::IsNullOrEmpty($lastPart)) {
                $defaultName = "Log_" + ($lastPart -replace '[^\w\-]', '_')
            }
        }
        return $defaultName
    }
}

# Format folder hierarchy for better readability
Function Format-FolderHierarchy {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FolderPath,
        
        [Parameter(Mandatory = $false)]
        [int]$IndentLevel = 0
    )
    
    $folderName = Split-Path -Leaf $FolderPath
    $indent = "  " * $IndentLevel
    
    return "$indent$folderName"
}

# Format hierarchical output for folder structure
function Format-Hierarchy {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$FolderPermissions
    )
    
    $results = New-Object System.Collections.Generic.List[PSObject]
    
    foreach ($path in ($FolderPermissions.Keys | Sort-Object)) {
        $permissions = $FolderPermissions[$path]
        $parentPath = Split-Path -Path $path -Parent
        
        $item = [PSCustomObject]@{
            Path = $path
            ParentPath = $parentPath
            Owner = $permissions.Owner
            AccessRules = $permissions.Access
            MatchingSubfolders = $(
                if ($permissions.ContainsKey('MatchingSubfolders')) {
                    $permissions.MatchingSubfolders
                } else {
                    @()
                }
            )
        }
        
        $results.Add($item)
    }
    
    return $results | Sort-Object Path
}

# Function to write hierarchical output
function Write-HierarchicalOutput {
    param (
        [Parameter(Mandatory=$true)]
        [hashtable]$Hierarchy,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$Permissions,
        
        [Parameter(Mandatory=$false)]
        [int]$Level = 0,
        
        [Parameter(Mandatory=$false)]
        [string]$ParentPath = ""
    )
    
    foreach ($key in ($Hierarchy.Keys | Sort-Object)) {
        $item = $Hierarchy[$key]
        $path = $item.Path
        $indent = "  " * $Level
        
        # Output folder name with indentation to show hierarchy
        Write-Log -Message "$indent├─ $key" -Color "Cyan" -Level 'INFO'
        
        # Check if we have permissions for this path
        if ($Permissions.ContainsKey($path)) {
            $perm = $Permissions[$path]
            $owner = $perm.Owner
            
            # Output permissions with increased indentation
            Write-Log -Message "$indent│  Owner: $owner" -Color "White" -Level 'INFO'
            
            # Show first 3 permissions (to avoid excessive output)
            $accessCount = $perm.Access.Count
            $showCount = [Math]::Min(3, $accessCount)
            
            for ($i = 0; $i -lt $showCount; $i++) {
                $access = $perm.Access[$i]
                $inherited = if ($access.IsInherited) { "(Inherited)" } else { "(Direct)" }
                Write-Log -Message "$indent│  $($access.IdentityReference) - $($access.FileSystemRights) $inherited" -Color "White" -Level 'INFO'
            }
            
            # Show count if there are more
            if ($accessCount -gt $showCount) {
                Write-Log -Message "$indent│  ... and $($accessCount - $showCount) more permissions" -Color "DarkGray" -Level 'INFO'
            }
        }
        
        # Process children recursively
        if ($item.Children.Count -gt 0) {
            Write-HierarchicalOutput -Hierarchy $item.Children -Permissions $Permissions -Level ($Level + 1) -ParentPath $path
        }
    }
}

# Consolidated log initialization with enhanced configuration
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$computerName = $env:COMPUTERNAME
$safePath = Get-SafeFilename -StartPath $StartPath
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
# Analysis Path: $StartPath
# Max Threads: $MaxThreads
# Max Depth: $MaxDepth
# Skip AD Resolution: $SkipADResolution
# Skip Uniqueness Counting: $SkipUniquenessCounting
# Enable SID Diagnostics: $EnableSIDDiagnostics
# Script Version: 1.14.0
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
    [array]$adminAccounts = @()  # Explicitly declare as array
    [array]$domains = @(Get-DomainControllers)  # Force array

    if ($domains.Count -eq 0) {
        Write-Log -Message "No domains found. Only checking local Administrator accounts." -Level 'WARNING' -Color "Yellow"
        
        # Get local Administrator account
        $wmiAdminAccounts = @(Get-WmiObject Win32_UserAccount -Filter "Name='Administrator' AND LocalAccount='True'" -ErrorAction SilentlyContinue)
    }
    else {
        Write-Log -Message "Checking Administrator accounts across $($domains.Count) domain(s)..." -Level 'INFO' -Color "Cyan"
        
        # Get all Administrator accounts (both local and domain)
        $wmiAdminAccounts = @(Get-WmiObject Win32_UserAccount -Filter "Name='Administrator'" -ErrorAction SilentlyContinue)
    }

    if ($wmiAdminAccounts -and $wmiAdminAccounts.Count -gt 0) {
        foreach ($account in $wmiAdminAccounts) {
            $domainType = if ($account.LocalAccount) { "Local" } else { "Domain" }
            $domain = if ($account.Domain) { $account.Domain } else { $env:COMPUTERNAME }

            # Get FQDN for domain accounts
            $fqdn = $domain
            if (-not $account.LocalAccount) {
                try {
                    $domainObj = $domains | Where-Object { $_.Name -like "*$domain*" } | Select-Object -First 1
                    if ($domainObj) {
                        $fqdn = $domainObj.Name
                    }
                }
                catch {
                    Write-Log -Message "Could not get FQDN for domain $domain" -Level 'DEBUG' -NoConsole
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
    }

    if ($adminAccounts.Count -eq 0) {
        Write-Log -Message "No Administrator accounts found. Using default SID pattern." -Level 'WARNING' -Color "Yellow"
        $script:AdminSID = "S-1-5-21-domain-500"
    }
    else {
        $script:AdminSID = $adminAccounts[0].SID
    }

    # Initialize well-known SIDs
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

# Update Write-ProgressStatus to prevent duplicate progress messages
function Write-ProgressStatus {
    param (
        [string]$Activity,
        [string]$Status,
        [int]$Current,
        [int]$Total
    )

    # Use script-level variables to track last progress update
    if (-not $script:lastProgressUpdate -or 
        (Get-Date) - $script:lastProgressUpdate -gt [TimeSpan]::FromSeconds(1)) {
        
        $percentComplete = if ($Total -gt 0) { [math]::Round(($Current / $Total) * 100, 2) } else { 0 }
        $currentFile = Split-Path $Status -Leaf
        Write-Progress -Activity "Analyzing Folder Permissions ($Current of $Total folders found)" -Status "Scanning: $currentFile" -PercentComplete $percentComplete
        $script:lastProgressUpdate = Get-Date
    }
}

# Main script execution section
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
    Write-Log -Message "Starting folder permission analysis for $StartPath" -Color "Cyan" -Level 'INFO'

    # Count total folders first (near the beginning of the try block)
    Write-Log -Message "Counting total folders in $StartPath..." -Color "Cyan" -Level 'INFO'
    $script:TotalFolders = Get-TotalFolderCount -StartPath $StartPath
    Write-Log -Message "Found $script:TotalFolders folders to process" -Color "Cyan" -Level 'INFO'
    $script:ProcessedFolders = 0
    $script:lastProgressUpdate = $null  # Initialize progress tracking

    # Process folders with timeout tracking
    Write-Log -Message "Starting folder permission analysis..." -Color "Cyan" -Level 'INFO'
    Invoke-FolderRecursively -StartPath $StartPath

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
    
    if ($script:FolderPermissions.Count -gt 0) {
        $hierarchyItems = Format-Hierarchy -FolderPermissions $script:FolderPermissions
        Write-HierarchicalOutput -Hierarchy $script:FolderPermissions -Permissions $script:FolderPermissions
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

function Compare-PermissionSets {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Parent,
        
        [Parameter(Mandatory = $true)]
        [object]$Child
    )
    
    # Ensure both objects have required properties
    if (-not ($Parent.PSObject.Properties.Name -contains "Owner") -or 
        -not ($Parent.PSObject.Properties.Name -contains "AccessRules") -or
        -not ($Child.PSObject.Properties.Name -contains "Owner") -or
        -not ($Child.PSObject.Properties.Name -contains "AccessRules")) {
        Write-Log -Message "Invalid permission object structure detected" -Level 'WARNING' -Color "Yellow"
        return $false
    }
    
    # Compare owners
    if ($Parent.Owner -ne $Child.Owner) {
        return $false
    }
    
    # Check if AccessRules is null or empty
    if ($null -eq $Parent.AccessRules -or $null -eq $Child.AccessRules) {
        Write-Log -Message "Null AccessRules detected" -Level 'WARNING' -Color "Yellow"
        return $false
    }
    
    # Check if they have same number of permissions
    if ($Parent.AccessRules.Count -ne $Child.AccessRules.Count) {
        return $false
    }
    
    # Create hashtables for easier comparison
    $parentRules = @{}
    foreach ($rule in $Parent.AccessRules) {
        if ($null -ne $rule.IdentityReference -and $null -ne $rule.FileSystemRights) {
            $key = "$($rule.IdentityReference)|$($rule.FileSystemRights)|$($rule.IsInherited)"
            $parentRules[$key] = $true
        }
    }
    
    # Check if all child rules exist in parent
    foreach ($rule in $Child.AccessRules) {
        if ($null -ne $rule.IdentityReference -and $null -ne $rule.FileSystemRights) {
            $key = "$($rule.IdentityReference)|$($rule.FileSystemRights)|$($rule.IsInherited)"
            if (-not $parentRules.ContainsKey($key)) {
                return $false
            }
        }
    }
    
    return $true
}

function Get-FolderPermissions {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FolderPath,
        
        [Parameter()]
        [int]$MaxDepth = [int]::MaxValue,
        
        [Parameter()]
        [int]$CurrentDepth = 0,
        
        [Parameter()]
        [string]$ParentPath = "",

        [Parameter()]
        [System.Collections.Generic.List[PSObject]]$Results = $null
    )
    
    if ($null -eq $Results) {
        $Results = New-Object System.Collections.Generic.List[PSObject]
    }
    
    if ($CurrentDepth -gt $MaxDepth) {
        return $Results
    }
    
    try {
        $folder = Get-Item -Path $FolderPath -ErrorAction Stop
        $acl = Get-Acl -Path $FolderPath -ErrorAction Stop
        
        $folderInfo = [PSCustomObject]@{
            Path = $folder.FullName
            ParentPath = $ParentPath
            Owner = $acl.Owner
            AccessRules = @($acl.Access | Select-Object IdentityReference, FileSystemRights, IsInherited)
            MatchingSubfolders = @()
        }
        
        $Results.Add($folderInfo)
        
        if ($CurrentDepth -lt $MaxDepth) {
            $subfolders = Get-ChildItem -Path $FolderPath -Directory -ErrorAction SilentlyContinue
            
            foreach ($subfolder in $subfolders) {
                try {
                    $subAcl = Get-Acl -Path $subfolder.FullName -ErrorAction Stop
                    
                    $subfolderId = [PSCustomObject]@{
                        Path = $subfolder.FullName
                        Owner = $subAcl.Owner
                        AccessRules = @($subAcl.Access | Select-Object IdentityReference, FileSystemRights, IsInherited)
                    }
                    
                    # Check if this subfolder has identical permissions to its parent
                    if (Compare-PermissionSets -Parent $folderInfo -Child $subfolderId) {
                        # Add to matching subfolders list instead of processing independently
                        $folderInfo.MatchingSubfolders += $subfolder.FullName
                    } else {
                        # Process this subfolder recursively
                        Get-FolderPermissions -FolderPath $subfolder.FullName -MaxDepth $MaxDepth `
                                             -CurrentDepth ($CurrentDepth + 1) -ParentPath $folder.FullName -Results $Results
                    }
                } catch {
                    Write-Warning "Unable to access permissions for $($subfolder.FullName): $_"
                }
            }
        }
    } catch {
        Write-Error "An error occurred: $_"
    }
    
    return $Results
}

