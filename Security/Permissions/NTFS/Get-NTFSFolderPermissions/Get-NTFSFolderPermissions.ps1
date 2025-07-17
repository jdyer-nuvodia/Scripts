# =============================================================================
# Script: Get-NTFSFolderPermissions.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-16 18:45:00 UTC
# Updated By: jdyer-nuvodia
# Version: 3.4.2
# Additional Info: Fixed syntax error causing PSScriptAnalyzer null reference exception:
#   - Added proper line break between cancellationTokenSource and processingTimeout variable assignments
#   - Separated comment from variable assignment to prevent parsing errors
#   - Fixed script structure to ensure proper PSScriptAnalyzer compliance checking
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
    [ValidateScript({ Test-Path $_ -PathType Container })]
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

process {
    # Enable strict mode and error handling
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    # Store parameters in script scope for access by all functions
    $script:MaxThreads = $MaxThreads
    $script:MaxDepth = $MaxDepth
    $script:SkipUniquenessCounting = $SkipUniquenessCounting
    $script:SkipADResolution = $SkipADResolution
    $script:EnableSIDDiagnostics = $EnableSIDDiagnostics

    # Script-level variables - consolidated to avoid duplication
    $script:TranscriptStarted = $false
    # Single standardized cache
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

    # Define script-level retry values for consistency
    $script:MaxRetries = 3
    $script:RetryDelay = 2

    # Add script-level cancellation token
    $script:cancellationTokenSource = New-Object System.Threading.CancellationTokenSource
    $script:processingTimeout = New-TimeSpan -Minutes $TimeoutMinutes

    # Function to get domain controllers and domain information
    function Get-DomainController {
        try {
            # Try to get domain information using .NET first
            $domainInfo = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            return @($domainInfo.DomainControllers | ForEach-Object {
                    [PSCustomObject]@{
                        Name            = $_.Name
                        Domain          = $domainInfo.Name
                        Forest          = $domainInfo.Forest.Name
                        IsGlobalCatalog = $_.IsGlobalCatalog
                    }
                })
        } catch {
            Write-NTFSLog -Message "Failed to get domain controllers using .NET: $_" -Level 'WARNING' -Color "Yellow"
            try {
                # Fallback to using AD cmdlets if available
                if (Get-Command Get-ADDomainController -ErrorAction SilentlyContinue) {
                    return @(Get-ADDomainController -Filter * | ForEach-Object {
                            [PSCustomObject]@{
                                Name            = $_.HostName
                                Domain          = $_.Domain
                                Forest          = $_.Forest
                                IsGlobalCatalog = $_.IsGlobalCatalog
                            }
                        })
                }
            } catch {
                Write-NTFSLog -Message "Failed to get domain controllers using AD cmdlets: $_" -Level 'WARNING' -Color "Yellow"
            }

            # If both methods fail, return computer domain info
            try {
                $computerDomain = (Get-CimInstance Win32_ComputerSystem).Domain
                if ($computerDomain) {
                    return @([PSCustomObject]@{
                            Name            = $env:COMPUTERNAME
                            Domain          = $computerDomain
                            Forest          = $computerDomain
                            IsGlobalCatalog = $false
                        })
                }
            } catch {
                Write-NTFSLog -Message "Failed to get computer domain info: $_" -Level 'WARNING' -Color "Yellow"
            }
        }

        # Return empty array if all methods fail
        return @()
    }

    # Function to count total folders recursively
    function Get-TotalFolderCount {
        param (
            [Parameter(Mandatory = $true)]
            [string]$StartPath
        )

        try {
            $folderCount = 0
            $processedPaths = @{}

            $folders = @(Get-ChildItem -Path $StartPath -Directory -Force -ErrorAction Stop)
            foreach ($folder in $folders) {
                if (-not $processedPaths.ContainsKey($folder.FullName)) {
                    $processedPaths[$folder.FullName] = $true
                    $folderCount++
                    try {
                        $folderCount += Get-TotalFolderCount -StartPath $folder.FullName
                    } catch {
                        Write-NTFSLog -Message "Error counting subfolders in $($folder.FullName): $_" -Level 'WARNING' -Color "Yellow"
                    }
                }
            }

            return $folderCount
        } catch {
            Write-NTFSLog -Message "Error counting folders in $StartPath : $_" -Level 'ERROR' -Color "Red"
            return 0
        }

    }
    # Add a script-level hashtable to track processed folders
    $script:ProcessedFolderPaths = @{}
    function Compare-PermissionSet {
        param (
            [Parameter(Mandatory = $true)]
            [object]$Parent,

            [Parameter(Mandatory = $true)]
            [object]$Child
        )

        # First check if any parameters are null
        if ($null -eq $Parent -or $null -eq $Child) {
            Write-NTFSLog -Message "Compare-PermissionSets: Null object detected" -Level 'DEBUG' -Color "Yellow" -NoConsole
            return $false
        }

        # Compare owners first - handle both direct owner and object property
        $parentOwner = if ($Parent.Owner) { $Parent.Owner } elseif ($Parent.PSObject.Properties['Owner']) { $Parent.Owner } else { $null }
        $childOwner = if ($Child.Owner) { $Child.Owner } elseif ($Child.PSObject.Properties['Owner']) { $Child.Owner } else { $null }

        if ([string]::IsNullOrEmpty($parentOwner) -or [string]::IsNullOrEmpty($childOwner) -or $parentOwner -ne $childOwner) {
            Write-NTFSLog -Message "Compare-PermissionSets: Owner mismatch or missing" -Level 'DEBUG' -Color "Yellow" -NoConsole
            return $false
        }

        # Get the access rules, handling different property names and types
        $parentRules = if ($Parent.Access) {
            $Parent.Access
        } elseif ($Parent.AccessRules) {
            $Parent.AccessRules
        } elseif ($Parent.PSObject.Properties['Access']) {
            $Parent.Access
        } else {
            $null
        }

        $childRules = if ($Child.Access) {
            $Child.Access
        } elseif ($Child.AccessRules) {
            $Child.AccessRules
        } elseif ($Child.PSObject.Properties['Access']) {
            $Child.Access
        } else {
            $null
        }

        # Validate access rules
        if ($null -eq $parentRules -or $null -eq $childRules) {
            Write-NTFSLog -Message "Compare-PermissionSets: Missing access rules" -Level 'DEBUG' -Color "Yellow" -NoConsole
            return $false
        }

        # Ensure we're working with arrays
        $parentRules = @($parentRules)
        $childRules = @($childRules)

        # Check if they have the same number of permissions
        if ($parentRules.Count -ne $childRules.Count) {
            Write-NTFSLog -Message "Compare-PermissionSets: Rule count mismatch Parent:$($parentRules.Count) Child:$($childRules.Count)" -Level 'DEBUG' -Color "Yellow" -NoConsole
            return $false
        }

        # Create normalized hashtables for comparison
        $parentHash = @{}
        $childHash = @{}

        # Convert parent rules to comparable format
        foreach ($rule in $parentRules) {
            try {
                $rights = $rule.FileSystemRights
                # Convert numerical rights to string representation
                if ($rights -match '^\-?\d+$') {
                    $rights = [System.Security.AccessControl.FileSystemRights]$rights
                }
                $key = "$($rule.IdentityReference)|$rights|$($rule.AccessControlType)|$($rule.IsInherited)"
                $parentHash[$key] = $true
            } catch {
                Write-NTFSLog -Message "Compare-PermissionSets: Error processing parent rule - $_" -Level 'DEBUG' -Color "Yellow" -NoConsole
                return $false
            }
        }

        # Compare child rules against parent
        foreach ($rule in $childRules) {
            try {
                $rights = $rule.FileSystemRights
                # Convert numerical rights to string representation
                if ($rights -match '^\-?\d+$') {
                    $rights = [System.Security.AccessControl.FileSystemRights]$rights
                }
                $key = "$($rule.IdentityReference)|$rights|$($rule.AccessControlType)|$($rule.IsInherited)"
                if (-not $parentHash.ContainsKey($key)) {
                    Write-NTFSLog -Message "Compare-PermissionSets: Child rule not found in parent" -Level 'DEBUG' -Color "Yellow" -NoConsole
                    return $false
                }
                $childHash[$key] = $true
            } catch {
                Write-NTFSLog -Message "Compare-PermissionSets: Error processing child rule - $_" -Level 'DEBUG' -Color "Yellow" -NoConsole
                return $false
            }
        }

        # Final validation that all rules were matched
        if ($parentHash.Count -ne $childHash.Count) {
            Write-NTFSLog -Message "Compare-PermissionSets: Hash count mismatch after comparison" -Level 'DEBUG' -Color "Yellow" -NoConsole
            return $false
        }

        return $true
    }

    # Modify the Invoke-FolderRecursively function
    function Invoke-FolderRecursively {
        param (
            [Parameter(Mandatory = $true)]
            [string]$StartPath,

            [Parameter(Mandatory = $false)]
            [int]$CurrentDepth = 0,

            [Parameter(Mandatory = $false)]
            [bool]$IsLeafNode = $false
        )

        try {
            # Check if we've already processed this path
            if ($script:ProcessedFolderPaths.ContainsKey($StartPath)) {
                Write-NTFSLog -Message "Skipping already processed path: $StartPath" -Level 'DEBUG' -NoConsole
                return
            }

            # Mark this path as processed
            $script:ProcessedFolderPaths[$StartPath] = $true

            # Get ACL and generate permission hash
            $acl = Get-Acl -Path $StartPath -ErrorAction Stop
            $permissionHash = Get-PermissionHash -AccessRules $acl.Access -Owner $acl.Owner

            # Only track unique permissions if not skipped for performance
            if (-not $script:SkipUniquenessCounting) {
                if (-not $script:UniquePermissions.ContainsKey($permissionHash)) {
                    $script:UniquePermissions[$permissionHash] = @{
                        Paths  = @($StartPath)
                        Owner  = $acl.Owner
                        Access = $acl.Access
                    }
                } else {
                    $script:UniquePermissions[$permissionHash].Paths += $StartPath
                }
            }

            # Get subfolders
            $folders = @(Get-ChildItem -Path $StartPath -Directory -Force -ErrorAction Stop)
            $IsLeafNode = ($folders.Count -eq 0)

            # Create folder permissions entry with MatchingSubfolders property
            $script:FolderPermissions[$StartPath] = @{
                Owner              = $acl.Owner
                Access             = $acl.Access
                IsInherited        = $true
                UniqueHash         = $permissionHash
                IsLeafNode         = $IsLeafNode
                ChildCount         = $folders.Count
                # Initialize empty list
                MatchingSubfolders = [System.Collections.Generic.List[string]]::new()
            }

            # Update progress
            $script:ProcessedFolders++
            Write-ProgressStatus -Activity "Analyzing Folder Permissions" -Status $StartPath -Current $script:ProcessedFolders -Total $script:TotalFolders

            if (-not $IsLeafNode) {
                foreach ($folder in $folders) {
                    if (-not $script:ProcessedFolderPaths.ContainsKey($folder.FullName)) {
                        $subAcl = Get-Acl -Path $folder.FullName -ErrorAction Stop

                        # Compare permissions with parent
                        if (Compare-PermissionSet -Parent $script:FolderPermissions[$StartPath] -Child @{
                                Owner  = $subAcl.Owner
                                Access = $subAcl.Access
                            }) {
                            # Add to matching subfolders if permissions are identical
                            $script:FolderPermissions[$StartPath].MatchingSubfolders.Add($folder.FullName)
                        } else {
                            # Process subfolder recursively if permissions differ
                            Invoke-FolderRecursively -StartPath $folder.FullName -CurrentDepth ($CurrentDepth + 1)
                        }
                    }
                }
            }
        } catch {
            Write-NTFSLog -Message "Error processing folder $StartPath : $_" -Level 'ERROR' -Color "Red"
        }
    }

    # Add function for sanitizing path for filename
    function Get-SafeFilename {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true, Position = 0)]
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
        } catch {
            Write-NTFSLog -Message "Error in Get-SafeFilename: $_" -Level 'ERROR' -Color "Red"
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
    function Format-FolderHierarchy {
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
                Path               = $path
                ParentPath         = $parentPath
                Owner              = $permissions.Owner
                AccessRules        = $permissions.Access
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

    # Format hierarchical output for folder structure
    function Write-HierarchicalOutput {
        param (
            [Parameter(Mandatory = $true)]
            [array]$Hierarchy,

            [Parameter(Mandatory = $true)]
            [hashtable]$Permissions,

            [Parameter(Mandatory = $false)]
            [int]$Level = 0,

            [Parameter(Mandatory = $false)]
            [string]$ParentPath = "",

            [Parameter(Mandatory = $false)]
            [System.Collections.Generic.HashSet[string]]$ProcessedPaths = $null
        )

        # Initialize processed paths tracking on first call
        if ($null -eq $ProcessedPaths) {
            $ProcessedPaths = [System.Collections.Generic.HashSet[string]]::new()
        }

        # Calculate indent based on level
        $indent = "    " * $Level

        # Get items at current level that haven't been processed
        $items = @($Hierarchy | Where-Object {
                $_.ParentPath -eq $ParentPath -and
                -not $ProcessedPaths.Contains($_.Path)
            } | Sort-Object Path)

        foreach ($item in $items) {
            $path = $item.Path

            # Skip if already processed
            if ($ProcessedPaths.Contains($path)) {
                continue
            }

            # Mark path as processed
            $ProcessedPaths.Add($path) | Out-Null

            $folderName = if ($Level -eq 0) { $path } else { Split-Path -Leaf $path }

            # Output folder name with proper indentation
            if ($Level -eq 0) {
                Write-NTFSLog -Message "$folderName" -Color "Cyan" -Level "INFO"
            } else {
                # Check if permissions are same as parent
                $parentPerms = $Permissions[$ParentPath]
                $currentPerms = $Permissions[$path]
                $hasSamePermissions = $false

                if ($parentPerms -and $currentPerms) {
                    $hasSamePermissions = Compare-PermissionSets -Parent $parentPerms -Child $currentPerms
                }

                if ($hasSamePermissions) {
                    Write-NTFSLog -Message "$indent|---+ $folderName (Same owner and permissions as parent)" -Color "DarkGray" -Level "INFO"
                } else {
                    Write-NTFSLog -Message "$indent|---+ $folderName" -Color "Cyan" -Level "INFO"
                }
            }

            # Get permissions for current folder
            $currentPerms = $Permissions[$path]
            if ($currentPerms) {
                # Only display owner and permissions details if not same as parent or level 0
                $parentPerms = if ($ParentPath) { $Permissions[$ParentPath] } else { $null }
                $hasSamePermissions = $false

                if ($parentPerms -and $currentPerms -and $Level -gt 0) {
                    $hasSamePermissions = Compare-PermissionSets -Parent $parentPerms -Child $currentPerms
                }

                # Only show permissions if this is the root level or has different permissions from parent
                if ($Level -eq 0 -or -not $hasSamePermissions) {
                    # Output Owner with correct indentation
                    Write-NTFSLog -Message "$indent|   Owner: $($currentPerms.Owner)" -Color "White" -Level "INFO"
                    Write-NTFSLog -Message "$indent|" -Level "INFO"

                    # Output Permissions with correct indentation
                    Write-NTFSLog -Message "$indent|   Permissions:" -Color "White" -Level "INFO"
                    foreach ($access in @($currentPerms.Access)) {
                        $inherited = if ($access.IsInherited) { "(Inherited)" } else { "(Direct)" }
                        Write-NTFSLog -Message "$indent|       $($access.IdentityReference) - $($access.FileSystemRights) $inherited" -Color "White" -Level "INFO"
                    }
                    Write-NTFSLog -Message "$indent|" -Level "INFO"
                }

                # Handle matching subfolders
                if ($currentPerms.MatchingSubfolders -and $currentPerms.MatchingSubfolders.Count -gt 0) {
                    Write-NTFSLog -Message "$indent|   Subfolders with identical permissions ($($currentPerms.MatchingSubfolders.Count)):" -Color "DarkGray" -Level "INFO"
                    Write-NTFSLog -Message "$indent|" -Level "INFO"
                    foreach ($subfolder in ($currentPerms.MatchingSubfolders | Sort-Object)) {
                        $subName = Split-Path -Leaf $subfolder
                        Write-NTFSLog -Message "$indent|---+ $subName (Same owner and permissions as parent)" -Color "DarkGray" -Level "INFO"
                        $ProcessedPaths.Add($subfolder) | Out-Null
                    }
                    Write-NTFSLog -Message "$indent|" -Level "INFO"
                }

                # Process different permission subfolders
                $children = @($Hierarchy | Where-Object {
                        $_.ParentPath -eq $path -and
                        -not $ProcessedPaths.Contains($_.Path) -and
                        $currentPerms.MatchingSubfolders -notcontains $_.Path
                    } | Sort-Object Path)

                if ($children.Count -gt 0) {
                    Write-NTFSLog -Message "$indent|   Subfolders with different permissions ($($children.Count)):" -Color "DarkGray" -Level "INFO"
                    Write-NTFSLog -Message "$indent|" -Level "INFO"

                    # Process each child folder
                    foreach ($child in $children) {
                        $childPath = $child.Path
                        $childName = Split-Path -Leaf $childPath

                        # Check if child permissions are same as current folder
                        $childPerms = $Permissions[$childPath]
                        $childHasSamePermissions = $false

                        if ($childPerms) {
                            $childHasSamePermissions = Compare-PermissionSets -Parent $currentPerms -Child $childPerms
                        }

                        # Write the child folder name with correct indent
                        if ($childHasSamePermissions) {
                            Write-NTFSLog -Message "$indent|---+ $childName (Same owner and permissions as parent)" -Color "DarkGray" -Level "INFO"
                        } else {
                            Write-NTFSLog -Message "$indent|---+ $childName" -Color "Cyan" -Level "INFO"
                        }

                        # Mark child path as processed
                        $ProcessedPaths.Add($childPath) | Out-Null

                        # Only display permissions if different from parent
                        if (-not $childHasSamePermissions) {
                            # Calculate the next level indent
                            $childIndent = "    " * ($Level + 1)

                            # Output Owner for child
                            Write-NTFSLog -Message "$childIndent|   Owner: $($childPerms.Owner)" -Color "White" -Level "INFO"
                            Write-NTFSLog -Message "$childIndent|" -Level "INFO"

                            # Output Permissions for child
                            Write-NTFSLog -Message "$childIndent|   Permissions:" -Color "White" -Level "INFO"
                            foreach ($access in @($childPerms.Access)) {
                                $inherited = if ($access.IsInherited) { "(Inherited)" } else { "(Direct)" }
                                Write-NTFSLog -Message "$childIndent|       $($access.IdentityReference) - $($access.FileSystemRights) $inherited" -Color "White" -Level "INFO"
                            }
                            Write-NTFSLog -Message "$childIndent|" -Level "INFO"
                        }

                        # Handle matching subfolders for child
                        if ($childPerms.MatchingSubfolders -and $childPerms.MatchingSubfolders.Count -gt 0) {
                            Write-NTFSLog -Message "$childIndent|   Subfolders with identical permissions ($($childPerms.MatchingSubfolders.Count)):" -Color "DarkGray" -Level "INFO"
                            Write-NTFSLog -Message "$childIndent|" -Level "INFO"
                            foreach ($subfolder in ($childPerms.MatchingSubfolders | Sort-Object)) {
                                $subName = Split-Path -Leaf $subfolder
                                Write-NTFSLog -Message "$childIndent|---+ $subName (Same owner and permissions as parent)" -Color "DarkGray" -Level "INFO"
                                $ProcessedPaths.Add($subfolder) | Out-Null
                            }
                            Write-NTFSLog -Message "$childIndent|" -Level "INFO"
                        }

                        # Find grand-children with different permissions
                        $grandChildren = @($Hierarchy | Where-Object {
                                $_.ParentPath -eq $childPath -and
                                -not $ProcessedPaths.Contains($_.Path) -and
                                $childPerms.MatchingSubfolders -notcontains $_.Path
                            } | Sort-Object Path)

                        if ($grandChildren.Count -gt 0) {
                            Write-NTFSLog -Message "$childIndent|   Subfolders with different permissions ($($grandChildren.Count)):" -Color "DarkGray" -Level "INFO"
                            Write-NTFSLog -Message "$childIndent|" -Level "INFO"

                            # Process each grand-child directly instead of using recursion
                            foreach ($grandChild in $grandChildren) {
                                $grandChildPath = $grandChild.Path
                                $grandChildName = Split-Path -Leaf $grandChildPath

                                # Check if grand-child permissions are same as child folder
                                $grandChildPerms = $Permissions[$grandChildPath]
                                $grandChildHasSamePermissions = $false

                                if ($grandChildPerms) {
                                    $grandChildHasSamePermissions = Compare-PermissionSets -Parent $childPerms -Child $grandChildPerms
                                }

                                # Write grand-child with proper indentation
                                if ($grandChildHasSamePermissions) {
                                    Write-NTFSLog -Message "$childIndent|---+ $grandChildName (Same owner and permissions as parent)" -Color "DarkGray" -Level "INFO"
                                } else {
                                    Write-NTFSLog -Message "$childIndent|---+ $grandChildName" -Color "Cyan" -Level "INFO"
                                }

                                # Mark as processed
                                $ProcessedPaths.Add($grandChildPath) | Out-Null

                                # Only display permissions if different from parent
                                if (-not $grandChildHasSamePermissions) {
                                    # Get and display grand-child permissions
                                    $grandChildIndent = "    " * ($Level + 2)

                                    # Output Owner
                                    Write-NTFSLog -Message "$grandChildIndent|   Owner: $($grandChildPerms.Owner)" -Color "White" -Level "INFO"
                                    Write-NTFSLog -Message "$grandChildIndent|" -Level "INFO"

                                    # Output Permissions
                                    Write-NTFSLog -Message "$grandChildIndent|   Permissions:" -Color "White" -Level "INFO"
                                    foreach ($access in @($grandChildPerms.Access)) {
                                        $inherited = if ($access.IsInherited) { "(Inherited)" } else { "(Direct)" }
                                        Write-NTFSLog -Message "$grandChildIndent|       $($access.IdentityReference) - $($access.FileSystemRights) $inherited" -Color "White" -Level "INFO"
                                    }
                                    Write-NTFSLog -Message "$grandChildIndent|" -Level "INFO"
                                }
                            }
                        }
                    }
                }
            }
        }

        # Add spacing between root-level items
        if ($Level -eq 0) {
            Write-NTFSLog -Message "" -Level "INFO"
        }
    }

    # Consolidated log initialization with enhanced configuration
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $computerName = $env:COMPUTERNAME
    $safePath = Get-SafeFilename -StartPath $StartPath
    $script:DebugLogFile = Join-Path $PSScriptRoot "NTFSPermissions_${computerName}_${safePath}_${timestamp}_debug.log"
    $script:TranscriptFile = Join-Path $PSScriptRoot "NTFSPermissions_${computerName}_${safePath}_${timestamp}_transcript.log"

    # Function to create a standardized log header with enhanced metadata
    function New-LogHeader {
        [CmdletBinding(SupportsShouldProcess)]
        [OutputType([string])]
        param()

        if ($PSCmdlet.ShouldProcess("Log file", "Create log header")) {
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
# Script Version: 3.4.0
# =============================================================================

"@
        }
    }

    # Initialize debug log with proper header
    Set-Content -Path $script:DebugLogFile -Value (New-LogHeader)

    # Define Write-NTFSLog function with standardized PowerShell format and enhanced metrics
    function Write-NTFSLog {
        param (
            [string]$Message,
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
        } catch {
            Write-Warning "Failed to write to debug log: $_"
        }

        if (-not $NoConsole) {
            Write-Output "$indentation$Message"
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

        Write-NTFSLog -Message $message -Level 'METRIC' -Category 'Performance' -NoConsole
    }

    # Initialize well-known SIDs
    function Initialize-WellKnownSID {
        # Explicitly declare as array
        [array]$adminAccounts = @()
        # Force array
        [array]$domains = @(Get-DomainController)

        if ($domains.Count -eq 0) {
            Write-NTFSLog -Message "No domains found. Only checking local Administrator accounts." -Level 'WARNING' -Color "Yellow"

            # Get local Administrator account
            $wmiAdminAccounts = @(Get-CimInstance Win32_UserAccount -Filter "Name = 'Administrator' AND LocalAccount = 'True'" -ErrorAction SilentlyContinue)
        } else {
            Write-NTFSLog -Message "Checking Administrator accounts across $($domains.Count) domain(s)..." -Level 'INFO' -Color "Cyan"

            # Get all Administrator accounts (both local and domain)
            $wmiAdminAccounts = @(Get-CimInstance Win32_UserAccount -Filter "Name = 'Administrator'" -ErrorAction SilentlyContinue)
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
                    } catch {
                        Write-NTFSLog -Message "Could not get FQDN for domain $domain" -Level 'DEBUG' -NoConsole
                    }
                }

                $adminAccounts += [PSCustomObject]@{
                    SID         = $account.SID
                    Domain      = $domain
                    FQDN        = $fqdn
                    DomainType  = $domainType
                    DisplayName = "$($account.SID) [${ domainType}: $fqdn]"
                }
            }
        }

        if ($adminAccounts.Count -eq 0) {
            Write-NTFSLog -Message "No Administrator accounts found. Using default SID pattern." -Level 'WARNING' -Color "Yellow"
            $script:AdminSID = "S-1-5-21-domain-500"
        } else {
            $script:AdminSID = $adminAccounts[0].SID
        }

        # Initialize well-known SIDs
        $script:WellKnownSIDs = @{
            "Nobody"             = "S-1-0-0"
            "Everyone"           = "S-1-1-0"
            "Local"              = "S-1-2-0"
            "CreatorOwner"       = "S-1-3-0"
            "CreatorGroup"       = "S-1-3-1"
            "Network"            = "S-1-5-2"
            "Interactive"        = "S-1-5-4"
            "AuthenticatedUsers" = "S-1-5-11"
            "LocalSystem"        = "S-1-5-18"
            "LocalService"       = "S-1-5-19"
            "NetworkService"     = "S-1-5-20"
            # Use first SID for lookups
            "Administrator"      = $script:AdminSID
            "Administrators"     = "S-1-5-32-544"
            "Users"              = "S-1-5-32-545"
            "Guests"             = "S-1-5-32-546"
        }
        Write-NTFSLog -Message "Initialized well-known SIDs collection with primary Administrator SID: $script:AdminSID" -Color "DarkGray" -NoConsole -Level 'DEBUG'

        # Return the administrator accounts collection
        return $adminAccounts
    }

    function Test-WellKnownSID {
        param(
            [Parameter(Mandatory = $true)]
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
            [Parameter(Mandatory = $true)]
            [string]$Sid
        )

        if ($script:SuppressedSids -contains $Sid) {
            Write-NTFSLog -Message "Skipping suppressed SID: $Sid" -Color "DarkGray" -NoConsole -Level 'DEBUG'
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

            Write-NTFSLog -Message "Attempting SID translation (attempt $attempt of $script:MaxRetries): $Sid" -Color "DarkGray" -NoConsole -Level 'DEBUG'

            # Try to resolve using .NET first
            try {
                $objSID = New-Object System.Security.Principal.SecurityIdentifier($Sid)
                $objName = $objSID.Translate([System.Security.Principal.NTAccount])
                $name = $objName.Value
                $script:SidCache[$Sid] = $name
                Write-NTFSLog -Message "Successfully resolved SID on attempt ${ attempt}: ${ Sid} -> ${ name}" -Color "Green" -NoConsole -Level 'SUCCESS'
                return $name
            } catch {
                # Fall back to AD lookup if .NET translation fails
                Write-NTFSLog -Message ".NET translation failed on attempt ${ attempt}, trying AD lookup for SID: $Sid" -Color "DarkGray" -NoConsole -Level 'DEBUG'
                if (-not $script:SkipADResolution) {
                    try {
                        $user = Get-ADUser -Identity $Sid -Properties SamAccountName -ErrorAction Stop
                        if ($user) {
                            $name = $user.SamAccountName
                            $script:SidCache[$Sid] = $name
                            Write-NTFSLog -Message "Successfully resolved SID via AD on attempt ${ attempt}: ${ Sid} -> ${ name}" -Color "Green" -NoConsole -Level 'SUCCESS'
                            return $name
                        }
                    } catch {
                        # Try as a group if user lookup failed
                        try {
                            $group = Get-ADGroup -Identity $Sid -Properties SamAccountName -ErrorAction Stop
                            if ($group) {
                                $name = $group.SamAccountName
                                $script:SidCache[$Sid] = $name
                                Write-NTFSLog -Message "Successfully resolved SID via AD (group) on attempt ${ attempt}: ${ Sid} -> ${ name}" -Color "Green" -NoConsole -Level 'SUCCESS'
                                return $name
                            }
                        } catch {
                            if ($script:EnableSIDDiagnostics) {
                                Write-NTFSLog -Message "AD resolution failed for SID: $Sid - $_" -Color "Yellow" -NoConsole -Level 'WARNING'
                            }
                            throw "Failed to resolve SID via AD: $_"
                        }
                    }
                } else {
                    throw "AD resolution skipped by user"
                }
            }
        } catch {
            Write-NTFSLog -Message "SID translation failed on attempt ${ attempt}: ${ Sid}" -Color "Yellow" -NoConsole -Level 'WARNING'
            if ($script:SidTranslationAttempts[$Sid] -lt $script:MaxRetries) {
                Write-NTFSLog -Message "Retrying in $script:RetryDelay seconds (attempt ${ attempt}/${ script:MaxRetries})..." -Color "DarkGray" -NoConsole -Level 'DEBUG'
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
            [Parameter(Mandatory = $true)]
            [string]$SID
        )

        try {
            if ($script:SidCache.ContainsKey($SID)) {
                return $script:SidCache[$SID]
            }

            # Get domain SID prefix
            $domainSid = $SID.Split('-')[0..4] -join '-'
            Write-NTFSLog "Attempting to translate SID: $SID (Domain prefix: $domainSid)" -Level 'DEBUG'

            $objSID = New-Object System.Security.Principal.SecurityIdentifier($SID)
            $objUser = $objSID.Translate([System.Security.Principal.NTAccount])

            $script:SidCache[$SID] = $objUser.Value
            Write-NTFSLog "Successfully translated $SID to $($objUser.Value)" -Level 'DEBUG'
            return $objUser.Value
        } catch {
            Write-NTFSLog "Failed to translate SID $SID : $_" -Level 'WARNING'
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
            [Parameter(Mandatory = $true)]
            [string]$SID
        )

        try {
            # Check if it's an Administrator SID from any domain
            if (Test-AdministratorSID -SID $SID) {
                Write-NTFSLog -Message "Found Administrator SID from domain: $SID" -Level 'DEBUG'
                return "ADMINISTRATOR (Domain: $(($SID -split '-')[1..3] -join '-'))"
            }

            # Try to convert SID to NT account name
            if (-not [string]::IsNullOrEmpty($SID)) {
                $objSID = New-Object System.Security.Principal.SecurityIdentifier($SID)
                $objUser = $objSID.Translate([System.Security.Principal.NTAccount])
                return $objUser.Value
            }
            return $SID
        } catch {
            Write-NTFSLog -Message "Error translating SID $SID : $_" -Level 'ERROR'
            return $SID
        }
    }

    # Standardized function to generate permission hashes
    function Get-PermissionHash {
        param (
            [Parameter(Mandatory = $true)]
            [object]$AccessRules,

            [Parameter(Mandatory = $true)]
            [string]$Owner,

            [Parameter(Mandatory = $false)]
            [bool]$IncludeInheritance = $true
        )

        $ownerPart = "OWNER:$Owner"
        if ($IncludeInheritance) {
            return "$ownerPart;" + ($AccessRules | ForEach-Object {
                    "$($_.IdentityReference)|$($_.FileSystemRights)|$($_.AccessControlType)|$($_.IsInherited)"
                }) -join ';'
        } else {
            return "$ownerPart;" + ($AccessRules | ForEach-Object {
                    "$($_.IdentityReference)|$($_.FileSystemRights)|$($_.AccessControlType)"
                }) -join ';'
        }

    }
    # Update Write-ProgressStatus to prevent duplicate progress messages
    function Write-ProgressStatus {
        param (
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
            Write-Output "Initializing transcript at: $script:TranscriptFile"
        } catch {
            Write-Warning "Failed to start transcript: $_"
            $script:TranscriptStarted = $false
        }

        # Register ctrl+c handler
        $null = [Console]::TreatControlCAsInput = $true
        Register-ObjectEvent -InputObject ([Console]) -EventName CancelKeyPress -Action {
            $script:cancellationTokenSource.Cancel()
            Write-NTFSLog "Cancellation requested by user (Ctrl+C)" -Level 'WARNING' -Color "Yellow"
        } | Out-Null

        # Output initial messages in correct order
        Write-NTFSLog -Message "Debug information will be written to: $script:DebugLogFile" -Level 'DEBUG'
        Write-NTFSLog ""

        # Initialize SIDs and get Administrator accounts
        $adminAccounts = Initialize-WellKnownSID

        # Always ensure $adminAccounts is an array for consistent behavior
        if ($null -eq $adminAccounts) {
            $adminAccounts = @()
        } elseif ($adminAccounts -isnot [Array] -and $adminAccounts -isnot [System.Collections.ICollection]) {
            # Convert single item to array
            $adminAccounts = @($adminAccounts)
        }

        # Display warning about multiple accounts if needed
        if ($adminAccounts.Count -gt 1) {
            Write-NTFSLog -Message "Multiple Administrator accounts found ($($adminAccounts.Count))" -Color "Yellow" -Level 'WARNING'
        }

        # Display each Administrator SID individually
        if ($adminAccounts.Count -gt 0) {
            foreach ($admin in $adminAccounts) {
                $domainName = if ($admin.DomainType -eq "Local") { "LOCAL" } else { $admin.FQDN }
                Write-NTFSLog -Message "The Administrator SID for $domainName is $($admin.SID)" -Color "White" -Level 'INFO'
            }
        } else {
            Write-NTFSLog -Message "No Administrator accounts found. Using default SID patterns." -Color "Yellow" -Level 'WARNING'
        }

        Write-NTFSLog ""
        Write-NTFSLog -Message "Starting folder permission analysis for $StartPath" -Color "Cyan" -Level 'INFO'

        # Count total folders first (near the beginning of the try block)
        Write-NTFSLog -Message "Counting total folders in $StartPath..." -Color "Cyan" -Level 'INFO'
        $script:TotalFolders = Get-TotalFolderCount -StartPath $StartPath
        Write-NTFSLog -Message "Found $script:TotalFolders folders to process" -Color "Cyan" -Level 'INFO'
        $script:ProcessedFolders = 0
        # Initialize progress tracking
        $script:lastProgressUpdate = $null

        # Process folders with timeout tracking
        Write-NTFSLog -Message "Starting folder permission analysis..." -Color "Cyan" -Level 'INFO'

        # Use MaxDepth parameter if provided (0 means no limit, convert to max value)
        $maxDepthToUse = if ($MaxDepth -eq 0) { [int]::MaxValue } else { $MaxDepth }

        # Call recursive processing with the user-specified parameters
        $results = Get-FolderPermission -FolderPath $StartPath -MaxDepth $maxDepthToUse -CurrentDepth 0 -ParentPath "" -Results @{}

        if ($script:cancellationTokenSource.Token.IsCancellationRequested) {
            Write-NTFSLog "`nProcessing terminated before completion" -Level 'WARNING' -Color "Yellow"
        }

        # Calculate elapsed time
        $script:EndTime = Get-Date
        $script:ElapsedTime = $script:EndTime - $script:StartTime

        # Display summary
        Write-NTFSLog -Message "`nAnalysis Complete" -Color "Green" -Level 'SUCCESS'
        Write-NTFSLog -Message "Total folders processed: $($script:ProcessedFolders)" -Color "Cyan" -Level 'INFO'

        # Only display unique permissions count if not skipped
        if (-not $script:SkipUniquenessCounting) {
            Write-NTFSLog -Message "Unique permission sets: $($script:UniquePermissions.Count)" -Color "Cyan" -Level 'INFO'
        } else {
            Write-NTFSLog -Message "Uniqueness counting was skipped for performance" -Color "DarkGray" -Level 'INFO'
        }

        Write-NTFSLog -Message "Elapsed time: $($script:ElapsedTime.ToString())" -Color "Cyan" -Level 'INFO'
        Write-NTFSLog -Message "" -Level 'INFO'

        # Create hierarchy structure
        $hierarchy = @()
        $hierarchy += [PSCustomObject]@{
            Path       = $StartPath
            ParentPath = ""
        }

        # Add all processed folders to hierarchy
        foreach ($path in $script:FolderPermissions.Keys | Where-Object { $_ -ne $StartPath }) {
            $hierarchy += [PSCustomObject]@{
                Path       = $path
                ParentPath = Split-Path -Parent $path
            }
        }

        # Sort hierarchy by path for consistent output
        $hierarchy = $hierarchy | Sort-Object Path

        # Write the hierarchical output
        if ($script:FolderPermissions.Count -gt 0) {
            Write-HierarchicalOutput -Hierarchy $hierarchy -Permissions $script:FolderPermissions
        } else {
            Write-NTFSLog -Message "No permissions data collected!" -Level 'WARNING' -Color "Yellow"
        }
    } catch [System.Exception] {
        Write-Error "An error occurred: $_"
        Write-Error $_.ScriptStackTrace
    } finally {
        Write-Progress -Activity "Analyzing Folder Permissions" -Completed

        try {
            Write-Output "Script execution completed. See $script:DebugLogFile for full details."

            # Clean up transcript only if we started one
            if ($script:TranscriptStarted) {
                Stop-Transcript -ErrorAction SilentlyContinue
            }
        } catch {
            Write-NTFSLog -Message "Error during cleanup: $_" -Level 'ERROR'
        } finally {
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

    function Get-FolderPermission {
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
                Path               = $folder.FullName
                ParentPath         = $ParentPath
                Owner              = $acl.Owner
                AccessRules        = @($acl.Access | Select-Object IdentityReference, FileSystemRights, IsInherited)
                MatchingSubfolders = @()
            }

            $Results.Add($folderInfo)

            if ($CurrentDepth -lt $MaxDepth) {
                $subfolders = Get-ChildItem -Path $FolderPath -Directory -ErrorAction SilentlyContinue

                # Use parallel processing if MaxThreads > 1 and we have multiple subfolders
                if ($script:MaxThreads -gt 1 -and $subfolders.Count -gt 1) {
                    Write-NTFSLog -Message "Processing $($subfolders.Count) subfolders with up to $script:MaxThreads threads" -Level 'DEBUG' -NoConsole

                    # Process subfolders in parallel batches
                    $batchSize = [Math]::Max(1, [Math]::Ceiling($subfolders.Count / $script:MaxThreads))
                    $batches = @()
                    for ($i = 0; $i -lt $subfolders.Count; $i += $batchSize) {
                        $batches += , @($subfolders[$i..([Math]::Min($i + $batchSize - 1, $subfolders.Count - 1))])
                    }

                    foreach ($batch in $batches) {
                        foreach ($subfolder in $batch) {
                            try {
                                $subAcl = Get-Acl -Path $subfolder.FullName -ErrorAction Stop

                                $subfolderId = [PSCustomObject]@{
                                    Path        = $subfolder.FullName
                                    Owner       = $subAcl.Owner
                                    AccessRules = @($subAcl.Access | Select-Object IdentityReference, FileSystemRights, IsInherited)
                                }

                                # Check if this subfolder has identical permissions to its parent
                                if (Compare-PermissionSet -Parent $folderInfo -Child $subfolderId) {
                                    # Add to matching subfolders list instead of processing independently
                                    $folderInfo.MatchingSubfolders += $subfolder.FullName
                                } else {
                                    # Process this subfolder recursively
                                    Get-FolderPermission -FolderPath $subfolder.FullName -MaxDepth $MaxDepth `
                                        -CurrentDepth ($CurrentDepth + 1) -ParentPath $folder.FullName -Results $Results
                                }
                            } catch {
                                Write-Warning "Unable to access permissions for $($subfolder.FullName): $_"
                            }
                        }
                    }
                } else {
                    # Single-threaded processing for small folder counts or MaxThreads = 1
                    foreach ($subfolder in $subfolders) {
                        try {
                            $subAcl = Get-Acl -Path $subfolder.FullName -ErrorAction Stop

                            $subfolderId = [PSCustomObject]@{
                                Path        = $subfolder.FullName
                                Owner       = $subAcl.Owner
                                AccessRules = @($subAcl.Access | Select-Object IdentityReference, FileSystemRights, IsInherited)
                            }

                            # Check if this subfolder has identical permissions to its parent
                            if (Compare-PermissionSet -Parent $folderInfo -Child $subfolderId) {
                                # Add to matching subfolders list instead of processing independently
                                $folderInfo.MatchingSubfolders += $subfolder.FullName
                            } else {
                                # Process this subfolder recursively
                                Get-FolderPermission -FolderPath $subfolder.FullName -MaxDepth $MaxDepth `
                                    -CurrentDepth ($CurrentDepth + 1) -ParentPath $folder.FullName -Results $Results
                            }
                        } catch {
                            Write-Warning "Unable to access permissions for $($subfolder.FullName): $_"
                        }
                    }
                }
            }
        } catch {
            Write-Error "An error occurred: $_"
        }

        return $Results
    }

}
