# Script: Add-FolderToPath.ps1
# Version: 2.0
# Description: Adds folders and subfolders to system PATH environment variable
# Author: jdyer-nuvodia
# Last Modified: 2025-02-05 22:15:38
#
# .SYNOPSIS
#   Adds specified folder and optionally its subfolders to the system or user PATH.
#
# .DESCRIPTION
#   This script adds a specified folder and optionally its subfolders to either the
#   system (Machine) or user PATH environment variable. It includes validation,
#   duplicate checking, and supports WhatIf operations.
#
# .PARAMETER RootPath
#   The root directory to add to PATH
#
# .PARAMETER NoRecurse
#   If specified, only adds the root folder without subfolders
#
# .PARAMETER Scope
#   Whether to modify Machine (system) or User PATH. Default is User
#
# .EXAMPLE
#   # Add single folder to user PATH
#   .\Add-FolderToPath.ps1 -RootPath "C:\Scripts"
#
# .EXAMPLE
#   # Add folder and subfolders to system PATH (requires admin)
#   .\Add-FolderToPath.ps1 -RootPath "C:\Scripts" -Scope Machine
#
# .EXAMPLE
#   # Test what would happen without making changes
#   .\Add-FolderToPath.ps1 -RootPath "C:\Scripts" -WhatIf
#
# .EXAMPLE
#   # Add folder without subfolders
#   .\Add-FolderToPath.ps1 -RootPath "C:\Scripts" -NoRecurse
#
# .EXAMPLE
#   # See detailed operation information
#   .\Add-FolderToPath.ps1 -RootPath "C:\Scripts" -Verbose
#

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
        param([string]$Path)
        
        try {
            return (Resolve-Path $Path).Path.TrimEnd('\')
        }
        catch {
            Write-Warning "Failed to resolve path: $Path"
            return $null
        }
    }

    # Function to add a path if it doesn't exist
    function Add-UniquePathItem {
        param([string]$Path)
        
        $sanitizedPath = Get-SanitizedPath $Path
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