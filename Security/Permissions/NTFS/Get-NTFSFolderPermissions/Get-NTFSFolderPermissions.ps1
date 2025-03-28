# =============================================================================
# Script: Get-NTFSFolderPermissions.ps1
# Created: 2025-03-15 18:30:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-28 19:12:22 UTC
# Updated By: jdyer-nuvodia
# Version: 3.3.1
# Additional Info: Fixed ParentPath property error in Format-Hierarchy function
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
    [int]$TimeoutMinutes = 120,
    
    [Parameter(Mandatory = $false)]
    [switch]$GroupIdenticalSubfolders = $true
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
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [System.Collections.Generic.List[PSObject]]$Items,
        
        [Parameter()]
        [string]$ParentPath = "",
        
        [Parameter()]
        [int]$CurrentDepth = 0,
        
        [Parameter()]
        [bool]$IsLast = $true,
        
        [Parameter()]
        [string]$Prefix = "",
        
        [Parameter()]
        [int]$MaxDepth = [int]::MaxValue
    )
    
    if ($CurrentDepth -gt $MaxDepth) {
        return
    }
    
    # Use a different approach to filter based on parent path
    $currentItems = $Items | Where-Object { 
        if ($_.PSObject.Properties.Match('ParentPath').Count -gt 0) {
            $_.ParentPath -eq $ParentPath
        } else {
            $false # Skip items without ParentPath property
        }
    }
    
    $count = $currentItems.Count
    $i = 0
    
    foreach ($item in $currentItems) {
        $i++
        $isLastItem = ($i -eq $count)
        
        $itemName = Split-Path -Leaf $item.Path
        $marker = if ($isLastItem) { "└─ " } else { "├─ " }
        $childPrefix = if ($isLastItem) { "   " } else { "│  " }
        
        Write-Host "$Prefix$marker$itemName" -ForegroundColor Cyan
        
        # Display owner and permissions
        Write-Host "$Prefix$childPrefix`Owner: $($item.Owner)" -ForegroundColor White
        
        foreach ($ace in $item.AccessRules) {
            $inheritedText = if ($ace.IsInherited) { " (Inherited)" } else { "" }
            Write-Host "$Prefix$childPrefix$($ace.IdentityReference) - $($ace.FileSystemRights)$inheritedText" -ForegroundColor White
        }
        
        # Display matching subfolders if any
        if ($item.MatchingSubfolders -and $item.MatchingSubfolders.Count -gt 0) {
            Write-Host "$Prefix$childPrefix`Matching subfolders with identical permissions:" -ForegroundColor Yellow
            foreach ($subfolder in $item.MatchingSubfolders) {
                $subfolderName = Split-Path -Leaf $subfolder
                Write-Host "$Prefix$childPrefix  - $subfolderName" -ForegroundColor DarkGray
            }
        }
        
        # Process children for this item - check if Path property exists first
        if ($item.PSObject.Properties.Match('Path').Count -gt 0) {
            Format-Hierarchy -Items $Items -ParentPath $item.Path -CurrentDepth ($CurrentDepth + 1) `
                            -IsLast $isLastItem -Prefix "$Prefix$childPrefix" -MaxDepth $MaxDepth
        }
    }
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
        
        # Check if Path property exists before using it
        $path = if ($item.PSObject.Properties.Match('Path').Count -gt 0) {
            $item.Path
        } else {
            $key
        }
        
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
        
        # Process children recursively - check if Children property exists
        if ($item.PSObject.Properties.Match('Children').Count -gt 0 -and $item.Children.Count -gt 0) {
            Write-HierarchicalOutput -Hierarchy $item.Children -Permissions $Permissions -Level ($Level + 1) -ParentPath $path
        }
    }
}

