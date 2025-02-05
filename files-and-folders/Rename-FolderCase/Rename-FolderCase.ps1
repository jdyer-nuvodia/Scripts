# Script: Rename-FolderCase.ps1
# Version: 1.1
# Description: Renames folders to proper PowerShell case convention (Verb-Noun)
# Author: jdyer-nuvodia
# Last Modified: 2025-02-05 22:25:10
#
# .SYNOPSIS
#   Renames folders to follow PowerShell case conventions (PascalCase with hyphens).
#
# .DESCRIPTION
#   This script renames folders to follow proper PowerShell naming conventions,
#   converting names to PascalCase with hyphens (e.g., "test-folder" becomes "Test-Folder").
#   It can process single folders or operate recursively, includes safety checks,
#   and supports WhatIf operations for testing before making changes.
#
# .PARAMETER Path
#   The path to the folder or directory to process
#
# .PARAMETER Recursive
#   If specified, processes all subfolders in the specified path
#
# .EXAMPLE
#   # Rename a single folder
#   .\Rename-FolderCase.ps1 -Path "C:\Scripts\test-mailboxexistence"
#
# .EXAMPLE
#   # Rename folder and all subfolders
#   .\Rename-FolderCase.ps1 -Path "C:\Scripts" -Recursive
#
# .EXAMPLE
#   # Test what would happen without making changes
#   .\Rename-FolderCase.ps1 -Path "C:\Scripts" -Recursive -WhatIf
#
# .EXAMPLE
#   # Show detailed progress with verbose output
#   .\Rename-FolderCase.ps1 -Path "C:\Scripts" -Recursive -Verbose
#
# .NOTES
#   - Requires appropriate permissions to rename folders
#   - Uses temporary renaming for case-only changes
#   - Processes folders from deepest to shallowest when using recursive option
#

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(Mandatory=$true,
               Position=0,
               ValueFromPipeline=$true,
               ValueFromPipelineByPropertyName=$true)]
    [string]$Path,

    [Parameter(Mandatory=$false)]
    [switch]$Recursive
)

function Convert-ToPascalCase {
    param([string]$text)
    
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
    return $result
}

function Rename-FolderWithCase {
    param(
        [string]$folderPath
    )
    
    try {
        $folder = Get-Item -LiteralPath $folderPath
        $parentPath = Split-Path -Path $folderPath -Parent
        $currentName = Split-Path -Path $folderPath -Leaf
        
        # Skip if it's a file
        if (!$folder.PSIsContainer) {
            return
        }

        # Convert name to proper case
        $newName = Convert-ToPascalCase -text $currentName
        
        # Skip if name wouldn't change
        if ($newName -eq $currentName) {
            Write-Verbose "Skipping '$currentName' - already in correct case"
            return
        }
        
        $newPath = Join-Path -Path $parentPath -ChildPath $newName
        
        # Handle case where only case is different (needs temp rename)
        if ($newPath.ToLower() -eq $folderPath.ToLower()) {
            $tempPath = Join-Path -Path $parentPath -ChildPath "_temp_$newName"
            
            if ($PSCmdlet.ShouldProcess($folderPath, "Rename to temp folder '$tempPath'")) {
                Rename-Item -LiteralPath $folderPath -NewName "_temp_$newName"
                Write-Verbose "Temporary rename: '$folderPath' -> '$tempPath'"
            }
            
            if ($PSCmdlet.ShouldProcess($tempPath, "Rename to final name '$newPath'")) {
                Rename-Item -LiteralPath $tempPath -NewName $newName
                Write-Host "Renamed: '$currentName' -> '$newName'" -ForegroundColor Green
            }
        }
        else {
            if ($PSCmdlet.ShouldProcess($folderPath, "Rename to '$newPath'")) {
                Rename-Item -LiteralPath $folderPath -NewName $newName
                Write-Host "Renamed: '$currentName' -> '$newName'" -ForegroundColor Green
            }
        }
    }
    catch {
        Write-Error "Error renaming folder '$folderPath': $_"
    }
}

try {
    # Verify path exists
    if (!(Test-Path -Path $Path)) {
        throw "Path '$Path' does not exist."
    }
    
    # Get folders to process
    $folders = @()
    if ($Recursive) {
        $folders = Get-ChildItem -Path $Path -Directory -Recurse
    }
    else {
        $folders = Get-ChildItem -Path $Path -Directory
    }
    
    # Process folders in reverse order (deepest first) to handle nested folders
    $folders = $folders | Sort-Object -Property FullName -Descending
    
    # Process each folder
    foreach ($folder in $folders) {
        Rename-FolderWithCase -folderPath $folder.FullName
    }
    
    # Process the root folder if it's a directory
    if ((Get-Item -Path $Path).PSIsContainer) {
        Rename-FolderWithCase -folderPath $Path
    }
}
catch {
    Write-Error "Script error: $_"
    exit 1
}