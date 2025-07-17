# =============================================================================
# Script: Add-FoldersToPath.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 2.5.1
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
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
     - Skips hidden folders (those starting with a dot)

    Dependencies:
     - PowerShell 5.1 or higher
     - Administrator rights (for Machine scope only)

    Security considerations:
     - Requires admin rights for Machine scope
     - Validates all paths before addition
     - Uses secure environment variable methods
     - Excludes hidden directories for security

    Performance impact:
     - Minimal for single folders
     - May take longer with recursive operations on deep folder structures
.PARAMETER StartPath
    The root directory to add to PATH. Must be a valid directory path.
.PARAMETER NoRecurse
    If specified, only adds the root folder without subfolders.
.PARAMETER Scope
    Whether to modify Machine (system) or User PATH. Default is User.
.EXAMPLE
    .\Add-FoldersToPath.ps1 -StartPath "C:\Scripts"
    Adds C:\Scripts to the user PATH
.NOTES
    Security Level: Medium
    Required Permissions: Local Admin (for Machine scope), User (for User scope)
    Validation Requirements:
     - Verify path exists
     - Check for admin rights when using Machine scope
     - Validate no duplicate entries
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true,
        Position = 0,
        ValueFromPipeline = $true,
        HelpMessage = "Root directory to add to PATH")]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$StartPath,

    [Parameter(Mandatory = $false)]
    [switch]$NoRecurse,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Machine', 'User')]
    [string]$Scope = 'User'
)

begin {
    # Set up logging
    $ScriptName = $MyInvocation.MyCommand.Name
    $Host.UI.RawUI.WindowTitle = $ScriptName
    $ComputerName = $env:COMPUTERNAME
    $Timestamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd_HHmmss")
    $LogFile = Join-Path -Path $PSScriptRoot -ChildPath "$ComputerName`_$ScriptName`_$Timestamp.log"

    function Write-LogEntry {
        param (
            [Parameter(Mandatory = $true)]
            [string]$Message,
            [ValidateSet('INFO', 'WARNING', 'ERROR', 'DEBUG', 'SUCCESS')]
            [string]$Level = 'INFO'
        )
        $TimeStamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss UTC")
        $LogEntry = "$TimeStamp [$Level] $Message"
        Add-Content -Path $LogFile -Value $LogEntry

        if ($Level -eq 'DEBUG' -and -not $VerbosePreference) {
            # Skip debug messages unless -Verbose is specified
            return
        }

        # Use appropriate output methods
        switch ($Level) {
            'ERROR' { Write-Error $LogEntry }
            'WARNING' { Write-Warning $Message }
            'DEBUG' { Write-Verbose $Message }
            'INFO' { Write-Information $Message -InformationAction Continue }
            'SUCCESS' { Write-Information "$Message" -InformationAction Continue }
            default { Write-Output $LogEntry }
        }
    }

    Write-LogEntry "Starting script execution" -Level INFO
    Write-LogEntry "Script version: 2.5.0" -Level INFO

    # Verify running as administrator for Machine scope
    if ($Scope -eq 'Machine' -and -not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-LogEntry "Administrator privileges required for Machine scope" -Level ERROR
        throw "Administrator privileges required for Machine scope. Please run as administrator or use User scope."
    }

    Write-LogEntry "Processing PATH changes for scope: $Scope" -Level INFO

    # Get current PATH based on scope
    try {
        $currentPath = [System.Environment]::GetEnvironmentVariable("Path", [System.EnvironmentVariableTarget]::$Scope)
        $currentPathArray = $currentPath -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        Write-LogEntry "Current PATH contains $($currentPathArray.Count) entries" -Level INFO
    } catch {
        Write-LogEntry "Failed to get current PATH: $_" -Level ERROR
        throw "Failed to get current PATH: $_"
    }
}

process {
    # Function to sanitize and validate path
    function Get-SanitizedPath {
        param([string]$StartPath)

        try {
            return (Resolve-Path $StartPath).Path.TrimEnd('\')
        } catch {
            Write-LogEntry "Failed to resolve path: $StartPath" -Level WARNING
            return $null
        }
    }

    # Function to add a path if it doesn't exist
    function Add-UniquePathItem {
        param([string]$StartPath)

        $sanitizedPath = Get-SanitizedPath $StartPath
        if ($null -eq $sanitizedPath) { return $null }

        if ($currentPathArray -notcontains $sanitizedPath) {
            Write-LogEntry "Adding new path: $sanitizedPath" -Level DEBUG
            return $sanitizedPath
        } else {
            Write-LogEntry "Path already exists: $sanitizedPath" -Level DEBUG
            return $null
        }
    }

    try {
        # Get all directories to process
        $directories = @()
        $rootPath = Get-SanitizedPath $StartPath

        if ($null -eq $rootPath) {
            Write-LogEntry "Root path could not be resolved: $StartPath" -Level ERROR
            throw "Root path could not be resolved."
        }

        $directories += $rootPath
        Write-LogEntry "Added root path to process list: $rootPath" -Level DEBUG

        # Using a more efficient method to get subdirectories
        if (-not $NoRecurse) {
            Write-LogEntry "Getting subdirectories for $StartPath (skipping hidden folders)" -Level INFO
            # Filter out hidden directories and any subdirectories within hidden directories
            $startTime = Get-Date
            $subDirs = Get-ChildItem -Path $StartPath -Directory -Recurse -ErrorAction Stop |
                Where-Object {
                    # Exclude directories that start with a dot or are within a hidden directory path
                    -not ($_.Name.StartsWith('.')) -and
                    -not ($_.FullName -match '\\\.[\w\-_]+\\')
                }
            $endTime = Get-Date
            $processingTime = ($endTime - $startTime).TotalSeconds
            Write-LogEntry "Found $($subDirs.Count) subdirectories (excluding hidden) in $processingTime seconds" -Level INFO
            $directories += $subDirs.FullName
        }

        Write-LogEntry "Total directories to process: $($directories.Count)" -Level INFO

        # Add unique paths more efficiently
        $startTime = Get-Date
        $newPaths = @()

        # Process in batches to improve performance
        $batchSize = 100
        for ($i = 0; $i -lt $directories.Count; $i += $batchSize) {
            $batch = $directories[$i..([Math]::Min($i + $batchSize - 1, $directories.Count - 1))]

            foreach ($dir in $batch) {
                $newPath = Add-UniquePathItem $dir
                if ($null -ne $newPath) {
                    $newPaths += $newPath
                }
            }

            # Progress update for large directories
            if ($directories.Count -gt $batchSize -and ($i % ($batchSize * 5) -eq 0)) {
                $percentComplete = [Math]::Min(100, [Math]::Floor(($i / $directories.Count) * 100))
                Write-LogEntry "Processing: $percentComplete% complete ($i of $($directories.Count))" -Level INFO
            }
        }

        $endTime = Get-Date
        $processingTime = ($endTime - $startTime).TotalSeconds
        Write-LogEntry "Path processing completed in $processingTime seconds" -Level INFO

        # Update PATH if we have new entries
        if ($newPaths.Count -gt 0) {
            $newPathString = ($currentPathArray + $newPaths) -join ";"

            if ($PSCmdlet.ShouldProcess("PATH Environment Variable", "Add $($newPaths.Count) new directories")) {
                Write-LogEntry "Updating PATH variable with $($newPaths.Count) new entries" -Level INFO
                [System.Environment]::SetEnvironmentVariable("Path", $newPathString, [System.EnvironmentVariableTarget]::$Scope)
                Write-LogEntry "Successfully added $($newPaths.Count) directories to PATH ($Scope scope)" -Level SUCCESS
                $newPaths | ForEach-Object {
                    Write-LogEntry "  + $_" -Level SUCCESS
                }
            } else {
                Write-LogEntry "WhatIf: Would add $($newPaths.Count) directories to PATH" -Level INFO
            }
        } else {
            Write-LogEntry "No new directories needed to be added to PATH" -Level WARNING
        }
    } catch {
        Write-LogEntry "Failed to process directories: $_" -Level ERROR
        Write-Error "Failed to process directories: $_"
        # Ensure we don't hang by explicitly returning
        return
    }
}

end {
    # Clean up and ensure PowerShell doesn't hang
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Write-LogEntry "Script execution completed successfully" -Level SUCCESS
    Write-LogEntry "Log file saved to: $LogFile" -Level INFO

    # Clear any variables that might be causing a hang
    Remove-Variable -Name currentPath, currentPathArray, directories, newPaths -ErrorAction SilentlyContinue
}
