# =============================================================================
# Script: Add-FoldersToPath.ps1
# Created: 2025-02-05 22:15:38 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-20 17:15:00 UTC
# Updated By: jdyer-nuvodia
# Version: 2.1
# Additional Info: Updated header format to match standardization requirements
# =============================================================================

<#
.SYNOPSIS
    Adds specified folder and optionally its subfolders to the system or user PATH.
.DESCRIPTION
    This script adds a specified folder and optionally its subfolders to either the
    system (Machine) or user PATH environment variable. It includes validation,
    duplicate checking, and supports WhatIf operations.
    
    Key actions:
     - Validates input paths and permissions
     - Checks for duplicates in PATH
     - Supports recursive folder addition
     - Provides verbose logging
    
    Dependencies:
     - PowerShell 5.1 or higher
     - Administrator rights (for Machine scope only)
    
    Security considerations:
     - Requires admin rights for Machine scope
     - Validates all paths before addition
     - Uses secure environment variable methods
    
    Performance impact:
     - Minimal for single folders
     - May take longer with recursive operations on deep folder structures
.PARAMETER RootPath
    The root directory to add to PATH. Must be a valid directory path.
.PARAMETER NoRecurse
    If specified, only adds the root folder without subfolders.
.PARAMETER Scope
    Whether to modify Machine (system) or User PATH. Default is User.
.EXAMPLE
    .\Add-FoldersToPath.ps1 -RootPath "C:\Scripts"
    Adds C:\Scripts to the user PATH
.NOTES
    Security Level: Medium
    Required Permissions: Local Admin (for Machine scope), User (for User scope)
    Validation Requirements: 
     - Verify path exists
     - Check for admin rights when using Machine scope
     - Validate no duplicate entries
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(Mandatory=$true,
               Position=0,
               ValueFromPipeline=$true,
               HelpMessage="Root directory to add to PATH")]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$RootPath,

    [Parameter(Mandatory=$false)]
    [switch]$NoRecurse,

    [Parameter(Mandatory=$false)]
    [ValidateSet('Machine', 'User')]
    [string]$Scope = 'User'
)

begin {
    # Verify running as administrator for Machine scope
    if ($Scope -eq 'Machine' -and -not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "Administrator privileges required for Machine scope. Please run as administrator or use User scope."
    }

    # Get current PATH based on scope
    try {
        $currentPath = [System.Environment]::GetEnvironmentVariable("Path", [System.EnvironmentVariableTarget]::$Scope)
        $currentPathArray = $currentPath -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        Write-Verbose "Current PATH contains $($currentPathArray.Count) entries"
    }
    catch {
        throw "Failed to get current PATH: $_"
    }
}

process {
    # Function to sanitize and validate path
    function Get-SanitizedPath {
        param([string]$StartPath)
        
        try {
            return (Resolve-Path $StartPath).Path.TrimEnd('\')
        }
        catch {
            Write-Warning "Failed to resolve path: $StartPath"
            return $null
        }
    }

    # Function to add a path if it doesn't exist
    function Add-UniquePathItem {
        param([string]$StartPath)
        
        $sanitizedPath = Get-SanitizedPath $StartPath
        if ($null -eq $sanitizedPath) { return $null }
        
        if ($currentPathArray -notcontains $sanitizedPath) {
            Write-Verbose "Adding new path: $sanitizedPath"
            return $sanitizedPath
        }
        else {
            Write-Verbose "Path already exists: $sanitizedPath"
            return $null
        }
    }

    try {
        # Get all directories to process
        $directories = @()
        $directories += Get-SanitizedPath $RootPath
        
        if (-not $NoRecurse) {
            Write-Verbose "Getting subdirectories for $RootPath"
            $subDirs = Get-ChildItem -Path $RootPath -Recurse -Directory -ErrorAction Stop
            $directories += $subDirs.FullName
        }

        # Add unique paths
        $newPaths = @()
        foreach ($dir in $directories) {
            $newPath = Add-UniquePathItem $dir
            if ($null -ne $newPath) {
                $newPaths += $newPath
            }
        }

        # Update PATH if we have new entries
        if ($newPaths.Count -gt 0) {
            $newPathString = ($currentPathArray + $newPaths) -join ";"
            
            if ($PSCmdlet.ShouldProcess("PATH Environment Variable", "Add $($newPaths.Count) new directories")) {
                [System.Environment]::SetEnvironmentVariable("Path", $newPathString, [System.EnvironmentVariableTarget]::$Scope)
                
                Write-Host "`nSuccessfully added $($newPaths.Count) directories to PATH ($Scope scope):" -ForegroundColor Green
                $newPaths | ForEach-Object { Write-Host "  + $_" -ForegroundColor Cyan }
            }
        }
        else {
            Write-Host "No new directories needed to be added to PATH." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Error "Failed to process directories: $_"
        return
    }
}

end {
    Write-Verbose "Script completed"
}
