# Script: Rename-ScriptToFolderName.ps1
# Version: 1.0
# Description: Renames PowerShell scripts to match their parent folder names
# Author: jdyer-nuvodia
# Last Modified: 2025-02-05 22:49:51
#
# .SYNOPSIS
#   Renames PowerShell scripts to match their parent folder names
#
# .DESCRIPTION
#   This script searches through folders and renames any .ps1 files to match
#   their parent folder name. It can process a single folder or recursively
#   process all subfolders. If multiple .ps1 files exist in a folder,
#   it will prompt for confirmation.
#
# .PARAMETER Path
#   The path to the folder to process
#
# .PARAMETER Recursive
#   If specified, processes all subfolders
#
# .EXAMPLE
#   # Process single folder
#   .\Rename-ScriptToFolderName.ps1 -Path "C:\Scripts\test-folder"
#
# .EXAMPLE
#   # Process folder and all subfolders
#   .\Rename-ScriptToFolderName.ps1 -Path "C:\Scripts" -Recursive
#
# .NOTES
#   - Requires appropriate permissions to rename files
#   - Will prompt for confirmation if multiple .ps1 files exist in a folder
#   - Uses WhatIf support for testing before making changes
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

function Rename-ScriptFile {
    param(
        [string]$FolderPath
    )
    
    try {
        # Get folder name
        $folderName = Split-Path -Path $FolderPath -Leaf
        
        # Get all .ps1 files in the folder
        $psFiles = Get-ChildItem -Path $FolderPath -Filter "*.ps1" -File
        
        if ($psFiles.Count -eq 0) {
            Write-Verbose "No PowerShell scripts found in: $FolderPath"
            return
        }
        
        if ($psFiles.Count -eq 1) {
            $psFile = $psFiles[0]
            $newName = "$folderName.ps1"
            
            # Skip if name already matches
            if ($psFile.Name -eq $newName) {
                Write-Verbose "Script '$($psFile.Name)' already matches folder name"
                return
            }
            
            # Rename the file
            $newPath = Join-Path -Path $FolderPath -ChildPath $newName
            if ($PSCmdlet.ShouldProcess($psFile.FullName, "Rename to '$newName'")) {
                Rename-Item -Path $psFile.FullName -NewName $newName
                Write-Host "Renamed '$($psFile.Name)' to '$newName'" -ForegroundColor Green
            }
        }
        else {
            # Multiple .ps1 files found - prompt for action
            Write-Host "`nMultiple PowerShell scripts found in: $FolderPath" -ForegroundColor Yellow
            Write-Host "Files found:" -ForegroundColor Yellow
            $psFiles | ForEach-Object { Write-Host "  - $($_.Name)" }
            
            $choice = Read-Host "`nDo you want to rename one of these files to '$folderName.ps1'? (y/n)"
            if ($choice -eq 'y') {
                for ($i = 0; $i -lt $psFiles.Count; $i++) {
                    Write-Host "[$i] $($psFiles[$i].Name)"
                }
                
                $fileChoice = Read-Host "`nEnter the number of the file to rename"
                if ($fileChoice -match '^\d+$' -and [int]$fileChoice -lt $psFiles.Count) {
                    $selectedFile = $psFiles[[int]$fileChoice]
                    $newName = "$folderName.ps1"
                    
                    if ($PSCmdlet.ShouldProcess($selectedFile.FullName, "Rename to '$newName'")) {
                        Rename-Item -Path $selectedFile.FullName -NewName $newName
                        Write-Host "Renamed '$($selectedFile.Name)' to '$newName'" -ForegroundColor Green
                    }
                }
                else {
                    Write-Warning "Invalid selection, skipping folder"
                }
            }
        }
    }
    catch {
        Write-Error "Error processing folder '$FolderPath': $_"
        Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    }
}

try {
    # Verify path exists
    if (!(Test-Path -Path $Path)) {
        throw "Path '$Path' does not exist."
    }
    
    Write-Host "Starting script rename process..." -ForegroundColor Cyan
    Write-Verbose "Processing path: $Path"
    Write-Verbose "Recursive mode: $Recursive"
    
    # Get folders to process
    $folders = @()
    if ($Recursive) {
        Write-Verbose "Getting all subfolders recursively..."
        $folders = Get-ChildItem -Path $Path -Directory -Recurse
        Write-Host "Found $($folders.Count) subfolders to process" -ForegroundColor Cyan
    }
    
    # Process each folder
    foreach ($folder in $folders) {
        Write-Host "`nProcessing folder: $($folder.FullName)" -ForegroundColor Cyan
        Rename-ScriptFile -FolderPath $folder.FullName
    }
    
    # Process the root folder
    if ((Get-Item -Path $Path).PSIsContainer) {
        Write-Host "`nProcessing root folder: $Path" -ForegroundColor Cyan
        Rename-ScriptFile -FolderPath $Path
    }
    
    Write-Host "`nScript rename process completed successfully." -ForegroundColor Green
}
catch {
    Write-Error "Script error: $_"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    exit 1
}