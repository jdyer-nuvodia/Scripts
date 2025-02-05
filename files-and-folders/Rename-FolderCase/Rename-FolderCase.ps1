# Script: Rename-FolderCase.ps1
# Version: 1.2
# Description: Renames folders to proper PowerShell case convention (Verb-Noun)
# Author: jdyer-nuvodia
# Last Modified: 2025-02-05 22:27:39
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

# Enable verbose output
$VerbosePreference = "Continue"

Write-Host "Script started - Processing path: $Path" -ForegroundColor Cyan
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
        [string]$folderPath
    )
    
    Write-Verbose "Processing folder: $folderPath"
    
    try {
        if (!(Test-Path -LiteralPath $folderPath)) {
            Write-Warning "Path not found: $folderPath"
            return
        }

        $folder = Get-Item -LiteralPath $folderPath
        if (!$folder.PSIsContainer) {
            Write-Verbose "Skipping non-folder item: $folderPath"
            return
        }

        $parentPath = Split-Path -Path $folderPath -Parent
        $currentName = Split-Path -Path $folderPath -Leaf
        
        Write-Verbose "Current folder name: $currentName"
        
        # Convert name to proper case
        $newName = Convert-ToPascalCase -text $currentName
        
        # Skip if name wouldn't change
        if ($newName -eq $currentName) {
            Write-Host "Skipping '$currentName' - already in correct case" -ForegroundColor Yellow
            return
        }
        
        $newPath = Join-Path -Path $parentPath -ChildPath $newName
        Write-Verbose "New path will be: $newPath"
        
        # Handle case where only case is different (needs temp rename)
        if ($newPath.ToLower() -eq $folderPath.ToLower()) {
            $tempPath = Join-Path -Path $parentPath -ChildPath "_temp_$newName"
            
            if ($PSCmdlet.ShouldProcess($folderPath, "Rename to temp folder '$tempPath'")) {
                Write-Host "Renaming '$folderPath' to temporary name..." -ForegroundColor Gray
                Rename-Item -LiteralPath $folderPath -NewName "_temp_$newName" -ErrorAction Stop
                Write-Verbose "Temporary rename successful"
                
                Write-Host "Renaming to final name..." -ForegroundColor Gray
                Rename-Item -LiteralPath $tempPath -NewName $newName -ErrorAction Stop
                Write-Host "Successfully renamed: '$currentName' -> '$newName'" -ForegroundColor Green
            }
        }
        else {
            if ($PSCmdlet.ShouldProcess($folderPath, "Rename to '$newPath'")) {
                Write-Host "Renaming '$currentName' to '$newName'..." -ForegroundColor Gray
                Rename-Item -LiteralPath $folderPath -NewName $newName -ErrorAction Stop
                Write-Host "Successfully renamed: '$currentName' -> '$newName'" -ForegroundColor Green
            }
        }
    }
    catch {
        Write-Error "Error processing folder '$folderPath': $_"
        Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    }
}

try {
    Write-Host "`nStarting folder case correction process..." -ForegroundColor Cyan
    
    # Verify path exists
    if (!(Test-Path -Path $Path)) {
        throw "Path '$Path' does not exist."
    }
    
    Write-Host "Getting list of folders to process..." -ForegroundColor Cyan
    # Get folders to process
    $folders = @()
    if ($Recursive) {
        Write-Verbose "Getting all subfolders recursively..."
        $folders = Get-ChildItem -Path $Path -Directory -Recurse
        Write-Host "Found $($folders.Count) subfolders" -ForegroundColor Cyan
    }
    else {
        Write-Verbose "Getting immediate subfolders only..."
        $folders = Get-ChildItem -Path $Path -Directory
        Write-Host "Found $($folders.Count) folders" -ForegroundColor Cyan
    }
    
    # Process folders in reverse order (deepest first) to handle nested folders
    Write-Host "Processing folders from deepest to shallowest..." -ForegroundColor Cyan
    $folders = $folders | Sort-Object -Property FullName -Descending
    
    # Process each folder
    foreach ($folder in $folders) {
        Write-Host "`nProcessing folder: $($folder.FullName)" -ForegroundColor Cyan
        Rename-FolderWithCase -folderPath $folder.FullName
    }
    
    # Process the root folder if it's a directory
    if ((Get-Item -Path $Path).PSIsContainer) {
        Write-Host "`nProcessing root folder: $Path" -ForegroundColor Cyan
        Rename-FolderWithCase -folderPath $Path
    }
    
    Write-Host "`nFolder case correction process completed successfully." -ForegroundColor Green
}
catch {
    Write-Error "Script error: $_"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    exit 1
}