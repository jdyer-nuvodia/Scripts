# =============================================================================
# Script: Get-NTFSPermissions.ps1
# Created: 2025-02-25 23:15:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-25 23:15:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1
# Additional Info: Added parameters and documentation
# =============================================================================

<#
.SYNOPSIS
    Gets NTFS permissions for a specified user in a directory structure.
.DESCRIPTION
    This script recursively checks NTFS permissions for a specified user
    across all folders under a given root directory.
.PARAMETER User
    The username to check permissions for. Must include domain name if on a domain (e.g., "DOMAIN\username")
.PARAMETER RootFolder
    The starting folder path to begin the recursive permission check
.EXAMPLE
    .\Get-NTFSPermissions.ps1 -User "DOMAIN\jsmith" -RootFolder "D:\"
    Checks permissions for user DOMAIN\jsmith starting from D:\ drive
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$User,
    
    [Parameter(Mandatory=$true)]
    [string]$RootFolder
)

Get-ChildItem -Directory -Path $RootFolder -Recurse -Force | ForEach-Object {
    $folder = $_.FullName
    $acl = Get-Acl $folder
    $userAccess = $acl.Access | Where-Object { $_.IdentityReference -eq $User }
    
    if ($userAccess) {
        [PSCustomObject]@{
            Folder = $folder
            User = $User
            Permissions = $userAccess.FileSystemRights
            IsInherited = $userAccess.IsInherited
        }
    }
} | Format-Table -AutoSize
