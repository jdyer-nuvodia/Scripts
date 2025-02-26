# =============================================================================
# Script: Rename-ScriptToFolderName.ps1
# Created: 2025-02-05 23:22:49 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-26 23:14:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.3
# Additional Info: Updated header format to meet new standards
# =============================================================================

<#
.SYNOPSIS
    Renames PowerShell scripts to match their parent folder names and case exactly
.DESCRIPTION
    This script searches through folders and renames any .ps1 files to match
    their parent folder name and case exactly. It can process a single folder or recursively
    process all subfolders. If multiple .ps1 files exist in a folder,
    it will prompt for confirmation.
    
    Key actions:
    - Preserves exact case of folder names
    - Handles single or multiple .ps1 files
    - Supports recursive folder processing
    
    Dependencies:
    - PowerShell 5.1 or higher
    - Write/Modify permissions on target folders
    
    Security considerations:
    - Requires appropriate file system permissions
    - No elevated privileges required
    
    Performance impact:
    - Minimal CPU and memory usage
    - File system operations are sequential
.PARAMETER Path
    The path to the folder to process. Must be a valid filesystem path.
.PARAMETER Recursive
    If specified, processes all subfolders recursively
.EXAMPLE
    .\Rename-ScriptToFolderName.ps1 -Path "C:\Scripts\Test-Folder"
    Process a single folder and rename any .ps1 files to match the folder name
.EXAMPLE
    .\Rename-ScriptToFolderName.ps1 -Path "C:\Scripts" -Recursive
    Process all folders recursively and rename .ps1 files to match their parent folders
.NOTES
    Security Level: Low
    Required Permissions: File system read/write access
    Validation Requirements:
    - Verify path exists
    - Confirm file rename operations
    - Check for file naming conflicts
#>

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
        # Get folder info to preserve exact case
        $folderInfo = Get-Item -Path $FolderPath
        $folderName = $folderInfo.Name  # This preserves the exact case
        
        Write-Host "`nProcessing folder: $FolderPath" -ForegroundColor Cyan
        Write-Host "Folder name (exact case): $folderName" -ForegroundColor Gray
        
        # Get all .ps1 files in the folder
        $psFiles = Get-ChildItem -Path $FolderPath -Filter "*.ps1" -File
        
        if ($psFiles.Count -eq 0) {
            Write-Host "  No PowerShell scripts found in this folder" -ForegroundColor Yellow
            return
        }
        
        Write-Host "  Found $($psFiles.Count) PowerShell script(s)" -ForegroundColor Gray
        
        if ($psFiles.Count -eq 1) {
            $psFile = $psFiles[0]
            $newName = "$folderName.ps1"
            
            # Skip if name already matches exactly (including case)
            if ($psFile.Name -ceq $newName) {
                Write-Host "  Script '$($psFile.Name)' already matches folder name and case" -ForegroundColor Yellow
                return
            }
            
            # Show case difference if only case is different
            if ($psFile.Name -eq $newName) {
                Write-Host "  Case difference detected:" -ForegroundColor Yellow
                Write-Host "    Current: $($psFile.Name)" -ForegroundColor Gray
                Write-Host "    New:     $newName" -ForegroundColor Gray
            }
            
            # Rename the file
            $newPath = Join-Path -Path $FolderPath -ChildPath $newName
            if ($PSCmdlet.ShouldProcess($psFile.FullName, "Rename to '$newName'")) {
                Write-Host "  Renaming '$($psFile.Name)' to '$newName'" -ForegroundColor Gray
                
                # Handle case-only changes
                if ($psFile.Name -eq $newName) {
                    $tempName = "_temp_" + [Guid]::NewGuid().ToString().Substring(0,8) + ".ps1"
                    Rename-Item -Path $psFile.FullName -NewName $tempName -Force
                    Rename-Item -Path (Join-Path -Path $FolderPath -ChildPath $tempName) -NewName $newName -Force
                }
                else {
                    Rename-Item -Path $psFile.FullName -NewName $newName -Force
                }
                
                Write-Host "  Successfully renamed to '$newName'" -ForegroundColor Green
            }
        }
        else {
            # Multiple .ps1 files found - prompt for action
            Write-Host "`n  Multiple PowerShell scripts found:" -ForegroundColor Yellow
            for ($i = 0; $i -lt $psFiles.Count; $i++) {
                Write-Host "    [$i] $($psFiles[$i].Name)"
            }
            
            $choice = Read-Host "`n  Do you want to rename one of these files to '$folderName.ps1'? (y/n)"
            if ($choice -eq 'y') {
                $fileChoice = Read-Host "  Enter the number of the file to rename"
                if ($fileChoice -match '^\d+$' -and [int]$fileChoice -lt $psFiles.Count) {
                    $selectedFile = $psFiles[[int]$fileChoice]
                    $newName = "$folderName.ps1"
                    
                    if ($PSCmdlet.ShouldProcess($selectedFile.FullName, "Rename to '$newName'")) {
                        Write-Host "  Renaming '$($selectedFile.Name)' to '$newName'" -ForegroundColor Gray
                        
                        # Handle case-only changes
                        if ($selectedFile.Name -eq $newName) {
                            $tempName = "_temp_" + [Guid]::NewGuid().ToString().Substring(0,8) + ".ps1"
                            Rename-Item -Path $selectedFile.FullName -NewName $tempName -Force
                            Rename-Item -Path (Join-Path -Path $FolderPath -ChildPath $tempName) -NewName $newName -Force
                        }
                        else {
                            Rename-Item -Path $selectedFile.FullName -NewName $newName -Force
                        }
                        
                        Write-Host "  Successfully renamed to '$newName'" -ForegroundColor Green
                    }
                }
                else {
                    Write-Warning "Invalid file selection '$fileChoice'. Please enter a number between 0 and $($psFiles.Count - 1)"
                    Write-Host "  Skipping folder: $FolderPath" -ForegroundColor Yellow
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
    Write-Host "Processing path: $Path" -ForegroundColor Gray
    Write-Host "Recursive mode: $Recursive" -ForegroundColor Gray
    
    # Initialize folders collection
    $folders = @()
    
    # Get the root folder
    if ((Get-Item -Path $Path).PSIsContainer) {
        $folders += Get-Item -Path $Path
    }
    
    # Add subfolders if recursive
    if ($Recursive) {
        Write-Host "Getting all subfolders recursively..." -ForegroundColor Gray
        $subFolders = Get-ChildItem -Path $Path -Directory -Recurse
        $folders += $subFolders
        Write-Host "Found $($subFolders.Count) subfolder(s)" -ForegroundColor Gray
    }
    
    # Sort folders by depth (deepest first) to handle nested folders properly
    $folders = $folders | Sort-Object { ($_.FullName -split '\\').Count } -Descending
    
    # Process each folder
    foreach ($folder in $folders) {
        Rename-ScriptFile -FolderPath $folder.FullName
    }
    
    Write-Host "`nScript rename process completed successfully." -ForegroundColor Green
}
catch {
    Write-Error "Script error: $_"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    exit 1
}