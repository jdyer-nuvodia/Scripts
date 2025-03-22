# =============================================================================
# Script: Search-ContentRecursively.ps1
# Created: 2025-03-17 21:00:00 UTC
# Author: nunya-nunya
# Last Updated: 2025-03-22 16:42:00 UTC
# Updated By: nunya-nunya
# Version: 1.4.1
# Additional Info: Modified to ignore log files in script directory
# =============================================================================

<#
.SYNOPSIS
Searches through directories and files recursively for a specified keyword.

.DESCRIPTION
This script performs a recursive search through directories and files, looking for
matches of a specified keyword. It searches both file/directory names and file contents.
Results are displayed with color coding for better visibility.

.PARAMETER Keyword
The search term to look for in file names and content.

.PARAMETER StartPath
The root directory path where the search should begin.

.EXAMPLE
.\Search-ContentRecursively.ps1 -Keyword "ConfigMgr" -StartPath "C:\Scripts"
Searches for "ConfigMgr" in the specified directory and all subdirectories

.EXAMPLE
.\Search-ContentRecursively.ps1 -Keyword "password" -StartPath "."
Searches for "password" in the current directory and all subdirectories
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true,
        Position = 0,
        HelpMessage = "Enter the keyword to search for")]
    [string]$Keyword,

    [Parameter(Mandatory = $true,
        Position = 1,
        HelpMessage = "Enter the starting directory path")]
    [string]$StartPath
)

# Define helper function first
function Write-ColorOutput {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $true)]
        [string]$ForegroundColor
    )
    
    Write-Host $Message -ForegroundColor $ForegroundColor
}

# Initialize logging
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $PSScriptRoot "Search-ContentRecursively_${timestamp}.transcript.log"
$null = Start-Transcript -Path $logFile

try {
    Write-ColorOutput "Transcript logging started at $logFile" -ForegroundColor Magenta

    # Validate the start path
    if (-not (Test-Path -Path $StartPath)) {
        Write-ColorOutput "Error: The specified path '$StartPath' does not exist." -ForegroundColor Red
        exit 1
    }

    Write-ColorOutput "Starting search for keyword '$Keyword' in path '$StartPath'..." -ForegroundColor Cyan

    Write-ColorOutput "`nSearching in metadata..." -ForegroundColor White
    try {
        $metadataMatches = Get-ChildItem -Path $StartPath -Recurse | ForEach-Object {
            $item = $_
            $metadata = Get-ItemProperty -Path $item.FullName -ErrorAction SilentlyContinue
            if ($metadata) {
                $props = $metadata.PSObject.Properties | 
                    Where-Object { $_.Value -is [string] -and $_.Value -match $Keyword }
                if ($props) {
                    foreach ($prop in $props) {
                        [PSCustomObject]@{
                            File = $item.FullName
                            Property = $prop.Name
                            Value = $prop.Value
                        }
                    }
                }
            }
        }

        if ($metadataMatches) {
            Write-ColorOutput "Found matches in metadata:" -ForegroundColor Green
            $metadataMatches | ForEach-Object {
                Write-ColorOutput "`nFile: $($_.File)" -ForegroundColor Yellow
                Write-ColorOutput "$($_.Property): $($_.Value)" -ForegroundColor White
            }
        } else {
            Write-ColorOutput "No matches found in metadata." -ForegroundColor DarkGray
        }
    } catch {
        Write-ColorOutput "Error occurred while searching metadata: $_" -ForegroundColor Red
    }

    # Search in file and directory names
    Write-ColorOutput "`nSearching in file and directory names..." -ForegroundColor White
    $nameMatches = Get-ChildItem -Path $StartPath -Recurse | 
        Where-Object { $_.Name -like "*$Keyword*" }

    if ($nameMatches) {
        Write-ColorOutput "Found matches in names:" -ForegroundColor Green
        foreach ($match in $nameMatches) {
            Write-ColorOutput "  $($match.FullName)" -ForegroundColor White
        }
    } else {
        Write-ColorOutput "No matches found in file or directory names." -ForegroundColor DarkGray
    }

    # Search in file contents
    Write-ColorOutput "`nSearching in file contents..." -ForegroundColor White
    try {
        $contentMatches = Get-ChildItem -Path $StartPath -Recurse -File |
            Where-Object { 
                $_.Extension -notmatch '\.(exe|dll|zip|png|jpg|jpeg|gif|pdf|doc|docx|xls|xlsx)$' -and
                -not ($_.DirectoryName -eq $PSScriptRoot -and $_.Extension -eq '.log')
            } |
            ForEach-Object {
                $file = $_
                $lineNumber = 1
                Get-Content $file.FullName -ErrorAction SilentlyContinue | 
                    ForEach-Object {
                        if ($_ -match $Keyword) {
                            [PSCustomObject]@{
                                File = $file.FullName
                                LineNumber = $lineNumber
                                Line = $_
                            }
                        }
                        $lineNumber++
                    }
            }

        if ($contentMatches) {
            Write-ColorOutput "Found matches in content:" -ForegroundColor Green
            $contentMatches | ForEach-Object {
                Write-ColorOutput "`nFile: $($_.File)" -ForegroundColor Yellow
                Write-ColorOutput "Line $($_.LineNumber): $($_.Line)" -ForegroundColor White
            }
        } else {
            Write-ColorOutput "No matches found in file contents." -ForegroundColor DarkGray
        }
    } catch {
        Write-ColorOutput "Error occurred while searching file contents: $_" -ForegroundColor Red
    }

    Write-ColorOutput "`nSearch completed." -ForegroundColor Cyan

    # Offer replacement if matches were found
    if ($contentMatches -or $nameMatches -or $metadataMatches) {
        $confirmation = Read-Host "`nWould you like to replace all instances of '$Keyword'? (Y/N)"
        if ($confirmation -eq 'Y') {
            $replaceWith = Read-Host "Enter the replacement string"
            
            Write-ColorOutput "`nPerforming replacements..." -ForegroundColor Cyan
            
            # Replace in file contents
            if ($contentMatches) {
                $contentMatches | Select-Object -ExpandProperty File -Unique | ForEach-Object {
                    $filePath = $_
                    try {
                        (Get-Content $filePath) | 
                            ForEach-Object { $_ -replace [regex]::Escape($Keyword), $replaceWith } |
                            Set-Content $filePath
                        Write-ColorOutput "Updated content in: $filePath" -ForegroundColor Green
                    }
                    catch {
                        Write-ColorOutput "Error updating $filePath : $_" -ForegroundColor Red
                    }
                }
            }

            # Rename files and folders
            if ($nameMatches) {
                $nameMatches | ForEach-Object {
                    try {
                        $newName = $_.Name -replace [regex]::Escape($Keyword), $replaceWith
                        $newPath = Join-Path (Split-Path $_.FullName -Parent) $newName
                        Rename-Item -Path $_.FullName -NewName $newName -ErrorAction Stop
                        Write-ColorOutput "Renamed: $($_.FullName) to $newPath" -ForegroundColor Green
                    }
                    catch {
                        Write-ColorOutput "Error renaming $($_.FullName): $_" -ForegroundColor Red
                    }
                }
            }

            Write-ColorOutput "`nReplacement operation completed." -ForegroundColor Cyan
        }
        else {
            Write-ColorOutput "`nReplacement operation cancelled." -ForegroundColor Yellow
        }
    }
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
} finally {
    if ($logFile) {
        Stop-Transcript
    }
}
