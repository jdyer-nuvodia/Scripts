# Script: Rename-FolderCase.ps1
# Version: 1.0
# Description: Renames folders to proper PowerShell case convention (Verb-Noun)
# Author: jdyer-nuvodia
# Created: 2025-02-05 22:06:01

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(Mandatory=$true,
               Position=0,
               ValueFromPipeline=$true,
               ValueFromPipelineByPropertyName=$true)]
    [string]$Path,

    [Parameter(Mandatory=$false)]
    [switch]$Recursive,

    [Parameter(Mandatory=$false)]
    [switch]$WhatIf
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