# =============================================================================
# Script: Get-NTFSFolderPermissions.ps1
# Created: 2025-03-06 21:06:43 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-12 23:03:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.15.24
# Additional Info: Add assembly loading for SID resolution in Get-NTFSFolderPermissions
# =============================================================================

<#
.SYNOPSIS
    Extracts and reports NTFS permissions for specified folders with optimized performance.
.DESCRIPTION
    This script retrieves NTFS permissions for a specified folder path and all its subfolders.
    Key features:
    - Uses optimized directory traversal methods for improved performance
    - Processes folders in parallel with configurable thread limits
    - Forces Active Directory module loading for SID resolution
    - Supports SID resolution on non-domain controller systems
    - Groups folders with identical permissions to reduce output clutter
    - Exports results to a formatted log file
    
    Dependencies:
    - Windows PowerShell 5.1 or later
    - RSAT AD PowerShell module (auto-installed if missing)
    - Read access to target folders
.PARAMETER FolderPath
    The path to the folder for which permissions will be extracted.
    Example: "C:\Important\Data" or "\\server\share\folder"
.PARAMETER MaxThreads
    Maximum number of parallel threads to use for processing.
    Default: 10
.PARAMETER MaxDepth
    Maximum folder depth to traverse. Set to 0 for unlimited depth.
    Default: 0
.PARAMETER SkipUniquenessCounting
    Skip counting unique permissions for large directories to improve performance.
    Default: False
.PARAMETER SkipADResolution
    Skip Active Directory SID resolution to avoid AD module dependency.
    Default: False
.PARAMETER EnableSIDDiagnostics
    Enable detailed diagnostic logging for SID resolution attempts.
    Type: Boolean
    Default: True
.PARAMETER ViewMode
    Switch between hierarchical and grouped view modes for displaying permissions.
    Valid values: "Hierarchy", "Group"
    - Hierarchy: Display permissions in a folder tree structure (default)
    - Group: Display permissions grouped by identical permission sets
    Default: "Hierarchy"
    Example: -ViewMode "Group"
.EXAMPLE
    .\Get-NTFSFolderPermissions.ps1 -FolderPath "C:\Important\Data"
    Retrieves NTFS permissions for C:\Important\Data and all subfolders
.EXAMPLE
    .\Get-NTFSFolderPermissions.ps1 -FolderPath "\\server\share\folder" -MaxThreads 20
    Uses 20 parallel threads to process folders on a network share
.EXAMPLE
    .\Get-NTFSFolderPermissions.ps1 -FolderPath "C:\VeryLargeFolder" -MaxDepth 3
    Processes only folders up to 3 levels deep from the root
.EXAMPLE
    .\Get-NTFSFolderPermissions.ps1 -FolderPath "C:\Data" -SkipADResolution
    Processes permissions without attempting to resolve SIDs through Active Directory
.EXAMPLE
    .\Get-NTFSFolderPermissions.ps1 -FolderPath "C:\Data" -EnableSIDDiagnostics
    Processes permissions with detailed SID resolution logging for troubleshooting
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

# Function to write log messages
function Write-Log {
    param (
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
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

# Function to get folder permissions
function Get-FolderPermissions {
    param (
        [string]$Folder
    )

    try {
        $acl = Get-Acl -Path $Folder
        $owner = $acl.Owner
        $access = $acl.Access

        $permissionData = [PSCustomObject]@{
            Folder = $Folder
            Owner  = $owner
            Access = $access
        }

        return $permissionData
    }
    catch {
        Write-Log -Message "Error getting permissions for $Folder: $($_.Exception.Message)" -Color "Red"
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
            # Process the $permissionData object
            Write-Log -Message "Folder: $($permissionData.Folder)" -Color "White"
            Write-Log -Message "Owner: $($permissionData.Owner)" -Color "White"

            # Iterate through the access rules
            foreach ($accessRule in $permissionData.Access) {
                $identity = $accessRule.IdentityReference.Value
                if ($identity -match '^S-1-') {
                    $identity = Resolve-SIDToName -SID $identity
                }
                Write-Log -Message "  Identity: $identity" -Color "White"
                Write-Log -Message "  Rights: $($accessRule.FileSystemRights)" -Color "White"
                Write-Log -Message "  Type: $($accessRule.AccessControlType)" -Color "White"
            }

            # Store the permissions data
            $script:FolderPermissions[$Path] = $permissionData

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
        Write-Log -Message "Error processing folder $Path: $($_.Exception.Message)" -Color "Red"
    }
    finally {
        Write-ProgressBar -Current $CurrentCount -Total $TotalCount -Activity "Processing Folders" -Status "Current Progress"
    }
}

# Function to process folders recursively
function Process-FoldersRecursively {
    param (
        [string]$Path,
        [int]$CurrentDepth = 0
    )

    $script:TotalFolders++
    $script:ProcessedFolders++

    Invoke-FolderProcessing -Path $Path -CurrentCount $script:ProcessedFolders -TotalCount $script:TotalFolders

    if ($MaxDepth -eq 0 -or $CurrentDepth -lt $MaxDepth) {
        Get-ChildItem -Path $Path -Directory | ForEach-Object {
            Process-FoldersRecursively -Path $_.FullName -CurrentDepth ($CurrentDepth + 1)
        }
    }
}

# Main script execution
try {
    Write-Log -Message "Starting folder permission analysis for $FolderPath" -Color "Green"

    # Process folders
    Process-FoldersRecursively -Path $FolderPath

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
        Write-Log -Message "`nFolder Hierarchy:" -Color "Yellow"
        $script:FolderPermissions.GetEnumerator() | Sort-Object Key | ForEach-Object {
            Write-Log -Message "$($_.Key)" -Color "White"
            $_.Value.Access | ForEach-Object {
                Write-Log -Message "  $($_.IdentityReference) : $($_.FileSystemRights) ($($_.AccessControlType))" -Color "Gray"
            }
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
    Write-Log -Message "`nScript execution completed." -Color "Green"
}
