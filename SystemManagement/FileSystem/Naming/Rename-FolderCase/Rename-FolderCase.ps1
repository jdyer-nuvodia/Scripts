# =============================================================================
# Script: Rename-FolderCase.ps1
# Created: 2025-02-05 22:33:29 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-26 23:13:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.4
# Additional Info: Script for renaming folders to proper PowerShell case convention
# =============================================================================

<#
.SYNOPSIS
    Renames folders to follow PowerShell case conventions (PascalCase with hyphens).
.DESCRIPTION
    This script renames folders to follow proper PowerShell naming conventions,
    converting names to PascalCase with hyphens (e.g., "test-folder" becomes "Test-Folder").

    Key actions:
    - Converts folder names to PascalCase with hyphens
    - Supports single folder or recursive processing
    - Includes safety checks and WhatIf support

    Dependencies:
    - Windows PowerShell 5.1 or later
    - Appropriate permissions to rename folders

    Security considerations:
    - Requires appropriate folder permissions
    - No elevated privileges required unless targeting system folders

    Performance impact:
    - Minimal CPU usage
    - Processing time depends on number of folders
.PARAMETER Path
    The path to the folder or directory to process
.PARAMETER Recursive
    If specified, processes all subfolders in the specified path
.EXAMPLE
    .\Rename-FolderCase.ps1 -Path "C:\Scripts\test-mailboxexistence"
    Renames a single folder to proper case convention
.NOTES
    Security Level: Low
    Required Permissions: File system read/write access
    Validation Requirements:
    - Test path existence
    - Verify write permissions
    - Validate successful rename operations
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(Mandatory=$true,
               Position=0,
               ValueFromPipeline=$true,
               ValueFromPipelineByPropertyName=$true)]
    [string]$StartPath,

    [Parameter(Mandatory=$false)]
    [switch]$Recursive
)

# Enable verbose output
$VerbosePreference = "Continue"

Write-Host "Script started - Processing path: $StartPath" -ForegroundColor Cyan
Write-Verbose "Recursive mode: $Recursive"

function Convert-ToPascalCase {
    param([string]$text)

    Write-Verbose "Converting to PascalCase: $text"

    # Split by common delimiters
    $words = $text -split '[-_\s]'

    # Convert each word to proper case
    $words = $words | ForEach-Object {
        if ($_.Length -gt 0) {
            $_.Substring(0,1).ToUpper() + $_.Substring(1).ToLower()
        }
    }

    # Rejoin with hyphens for PowerShell convention
    $result = $words -join '-'
    Write-Verbose "Converted to: $result"
    return $result
}

function Rename-FolderWithCase {
    param(
        [string]$StartPath
    )

    try {
        $folder = Get-Item -LiteralPath $StartPath
        $parentPath = Split-Path -Path $StartPath -Parent
        $currentName = Split-Path -Path $StartPath -Leaf

        # Skip if it's a file
        if (!$folder.PSIsContainer) {
            return
        }

        # Convert name to proper case
        $newName = Convert-ToPascalCase -text $currentName

        # Compare with exact case - Using case-sensitive comparison
        if ($currentName -cne $newName) {
            Write-Host "Renaming '$currentName' to '$newName'" -ForegroundColor Cyan
            $newPath = Join-Path -Path $parentPath -ChildPath $newName

            # Handle case where only case is different (needs temp rename)
            if ($newPath.ToLower() -eq $StartPath.ToLower()) {
                $tempName = "_temp_" + [Guid]::NewGuid().ToString().Substring(0,8)
                $tempPath = Join-Path -Path $parentPath -ChildPath $tempName

                if ($PSCmdlet.ShouldProcess($StartPath, "Rename to temp folder '$tempPath'")) {
                    Write-Verbose "Temporary rename: '$StartPath' -> '$tempPath'"
                    Rename-Item -LiteralPath $StartPath -NewName $tempName -ErrorAction Stop

                    Write-Verbose "Final rename: '$tempPath' -> '$newPath'"
                    Rename-Item -LiteralPath $tempPath -NewName $newName -ErrorAction Stop
                    Write-Host "Successfully renamed: '$currentName' -> '$newName'" -ForegroundColor Green
                }
            }
            else {
                if ($PSCmdlet.ShouldProcess($StartPath, "Rename to '$newPath'")) {
                    Rename-Item -LiteralPath $StartPath -NewName $newName -ErrorAction Stop
                    Write-Host "Successfully renamed: '$currentName' -> '$newName'" -ForegroundColor Green
                }
            }
        }
        else {
            Write-Verbose "Skipping '$currentName' - already in correct case"
        }
    }
    catch {
        Write-Error "Error processing folder '$StartPath': $_"
        Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    }
}

try {
    Write-Host "`nStarting folder case correction process..." -ForegroundColor Cyan

    # Verify path exists
    if (!(Test-Path -Path $StartPath)) {
        throw "Path '$StartPath' does not exist."
    }

    Write-Host "Getting list of folders to process..." -ForegroundColor Cyan
    # Get folders to process
    $folders = @()
    if ($Recursive) {
        Write-Verbose "Getting all subfolders recursively..."
        $folders = Get-ChildItem -Path $StartPath -Directory -Recurse
        Write-Host "Found $($folders.Count) subfolders" -ForegroundColor Cyan
    }
    else {
        Write-Verbose "Getting immediate subfolders only..."
        $folders = Get-ChildItem -Path $StartPath -Directory
        Write-Host "Found $($folders.Count) folders" -ForegroundColor Cyan
    }

    # Process folders in reverse order (deepest first) to handle nested folders
    Write-Host "Processing folders from deepest to shallowest..." -ForegroundColor Cyan
    $folders = $folders | Sort-Object -Property FullName -Descending

    # Process each folder
    foreach ($folder in $folders) {
        Write-Host "`nProcessing folder: $($folder.FullName)" -ForegroundColor Cyan
        Rename-FolderWithCase -StartPath $folder.FullName
    }

    # Process the root folder if it's a directory
    if ((Get-Item -Path $StartPath).PSIsContainer) {
        Write-Host "`nProcessing root folder: $StartPath" -ForegroundColor Cyan
        Rename-FolderWithCase -StartPath $StartPath
    }

    Write-Host "`nFolder case correction process completed successfully." -ForegroundColor Green
}
catch {
    Write-Error "Script error: $_"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    exit 1
}
