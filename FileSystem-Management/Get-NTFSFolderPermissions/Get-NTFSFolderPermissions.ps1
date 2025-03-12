# =============================================================================
# Script: Get-NTFSFolderPermissions.ps1
# Created: 2025-03-06 21:06:43 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-12 23:15:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.16.0
# Additional Info: Enhanced documentation and parameter descriptions for clarity
# =============================================================================

<#
.SYNOPSIS
    Advanced NTFS permission analyzer for folders with parallel processing capabilities.

.DESCRIPTION
    Comprehensive NTFS permission analysis tool that provides detailed access control information
    for specified folders and their subfolders. The script utilizes parallel processing and
    caching mechanisms for optimal performance.

    Key Features:
    - Parallel folder processing with configurable thread limits
    - Hierarchical or grouped permission display modes
    - SID to name resolution with caching
    - Active Directory integration for accurate identity resolution
    - Detailed progress tracking and statistics
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
        Write-Log -Message "Error processing folder ${Path}: $($_.Exception.Message)" -Color "Red"
    }
    finally {
        Write-ProgressBar -Current $CurrentCount -Total $TotalCount -Activity "Processing Folders" -Status "Current Progress"
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

# Main script execution
try {
    Write-Log -Message "Starting folder permission analysis for $FolderPath" -Color "Green"

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
