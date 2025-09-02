# =============================================================================
# Script: Search-ContentRecursively.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-16 21:15:00 UTC
# Updated By: jdyer-nuvodia
# Version: 2.5.0
# Additional Info: Fixed all indentation inconsistencies using 4-space indentation throughout
# =============================================================================

<#
.SYNOPSIS
Searches through directories and files recursively for specified keywords and optionally replaces them.

.DESCRIPTION
This script performs a recursive search through directories and files, looking for
matches of specified keywords. It searches both file/directory names and file contents.
Results are displayed with color coding for better visibility. The script can also
replace found keywords with specified replacement strings.

.PARAMETER SearchReplacePairs
A hashtable containing search terms as keys and their replacement values.
Example: @{ "oldtext"="newtext"; "anotherold"="anothernew"}

.PARAMETER StartPath
The root directory path where the search should begin.

.PARAMETER ExcludePath
The full path to a folder that should be excluded from search and replacement operations.
All files and subfolders within the excluded path will be ignored.

.PARAMETER AutoReplace
When specified, automatically performs replacements without prompting for confirmation.

.EXAMPLE
.\Search-ContentRecursively.ps1 -SearchReplacePairs @{ "ConfigMgr"="SCCM"} -StartPath "C:\Scripts"
Searches for "ConfigMgr" in the specified directory and prompts to replace with "SCCM"

.EXAMPLE
.\Search-ContentRecursively.ps1 -SearchReplacePairs @{ "jdyer-nuvodia"="moo"; "jdyer\OneDrive - Nuvodia\"="moo"} -StartPath "C:\Github" -AutoReplace
Searches for both terms and automatically replaces them with "moo" without prompting

.EXAMPLE
.\Search-ContentRecursively.ps1 -SearchReplacePairs @{ "ConfigMgr"="SCCM"} -StartPath "C:\Scripts" -ExcludePath "C:\Scripts\Archive"
Searches for "ConfigMgr" in the specified directory, excluding the Archive folder and its contents

.EXAMPLE
.\Search-ContentRecursively.ps1 -Keyword "password" -StartPath "."
Searches for "password" in the current directory (backward compatibility)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false,
        HelpMessage = "Hashtable of search terms and their replacements")]
    [hashtable]$SearchReplacePairs,

    [Parameter(Mandatory = $true,
        HelpMessage = "Enter the starting directory path")]
    [string]$StartPath,

    [Parameter(Mandatory = $false,
        HelpMessage = "Enter the keyword to search for (for backward compatibility)")]
    [string]$Keyword,

    [Parameter(Mandatory = $false,
        HelpMessage = "Full path to a folder to exclude from search and replacement operations")]
    [string]$ExcludePath,

    [Parameter(Mandatory = $false,
        HelpMessage = "Automatically replace without prompting")]
    [switch]$AutoReplace
)

# Define helper function first
function Write-ColorOutput {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $true)]
        [string]$ForegroundColor
    )

    # Format message with color-coded prefixes and output to both console and transcript
    $formattedMessage = switch ($ForegroundColor) {
        'Red' { "ERROR: $Message" }
        'Yellow' { "WARNING: $Message" }
        'Green' { "SUCCESS: $Message" }
        'Cyan' { "INFO: $Message" }
        'Magenta' { "DEBUG: $Message" }
        'DarkGray' { "DETAIL: $Message" }
        default { $Message }
    }

    # Write to transcript log (captured by Start-Transcript)
    Write-Output $formattedMessage
}

# Helper function to check if a path should be excluded
function Test-ExcludedPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ItemPath,

        [Parameter(Mandatory = $false)]
        [string]$ExcludePath
    )

    if ([string]::IsNullOrWhiteSpace($ExcludePath)) {
        return $false
    }

    # Normalize paths for comparison
    $normalizedItemPath = [System.IO.Path]::GetFullPath($ItemPath)
    $normalizedExcludePath = [System.IO.Path]::GetFullPath($ExcludePath)

    # Check if the item path starts with the exclude path
    return $normalizedItemPath.StartsWith($normalizedExcludePath, [System.StringComparison]::OrdinalIgnoreCase)
}

# Initialize logging
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$systemName = $env:COMPUTERNAME
$logFile = Join-Path $PSScriptRoot "Search-ContentRecursively_${systemName}_${timestamp}.log"
$null = Start-Transcript -Path $logFile

try {
    Write-ColorOutput "Transcript logging started at $logFile" -ForegroundColor Magenta

    # Validate the start path
    if (-not (Test-Path -Path $StartPath)) {
        Write-ColorOutput "Error: The specified path '$StartPath' does not exist." -ForegroundColor Red
        exit 1
    }

    # Validate the exclude path if provided
    if ($ExcludePath -and -not (Test-Path -Path $ExcludePath)) {
        Write-ColorOutput "Warning: The specified exclude path '$ExcludePath' does not exist. Continuing without exclusion." -ForegroundColor Yellow
        $ExcludePath = $null
    }

    if ($ExcludePath) {
        Write-ColorOutput "Excluding path: '$ExcludePath'" -ForegroundColor Cyan
    }

    # Handle parameters - support both new hashtable and legacy keyword parameter
    if (-not $SearchReplacePairs -and $Keyword) {
        $SearchReplacePairs = @{ $Keyword = "" }
        # Empty replacement will be filled during interactive prompt
    } elseif (-not $SearchReplacePairs -and -not $Keyword) {
        Write-ColorOutput "Error: Either -Keyword or -SearchReplacePairs must be provided." -ForegroundColor Red
        exit 1
    }

    Write-ColorOutput "Starting search in path '$StartPath' for keywords: '$($SearchReplacePairs.Keys -join "', '")'..." -ForegroundColor Cyan
    Write-ColorOutput "`nSearching in metadata..." -ForegroundColor White

    $metadataMatches = @()
    try {
        $metadataMatches = Get-ChildItem -Path $StartPath -Recurse -Force |
            Where-Object { -not (Test-ExcludedPath -ItemPath $_.FullName -ExcludePath $ExcludePath) } |
            ForEach-Object {
                $item = $_
                $metadata = Get-ItemProperty -Path $item.FullName -ErrorAction SilentlyContinue
                if ($metadata) {
                    foreach ($searchTerm in $SearchReplacePairs.Keys) {
                        # Escape the search term for regex pattern matching
                        $escapedSearchTerm = [regex]::Escape($searchTerm)
                        $props = $metadata.PSObject.Properties |
                            Where-Object { $_.Value -is [string] -and $_.Value -imatch $escapedSearchTerm }
                            if ($props) {
                                foreach ($prop in $props) {
                                    [PSCustomObject]@{
                                        File       = $item.FullName
                                        Property   = $prop.Name
                                        Value      = $prop.Value
                                        SearchTerm = $searchTerm
                                    }
                                }
                            }
                        }
                    }
                }
    } catch {
        Write-ColorOutput "Error occurred while searching metadata: $_" -ForegroundColor Red
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

    # Search in file and directory names
    Write-ColorOutput "`nSearching in file and directory names..." -ForegroundColor White
    $nameMatches = @()
    foreach ($searchTerm in $SearchReplacePairs.Keys) {
        $foundItems = Get-ChildItem -Path $StartPath -Recurse -Force |
            Where-Object {
                $_.Name -ilike "*$searchTerm*" -and
                -not (Test-ExcludedPath -ItemPath $_.FullName -ExcludePath $ExcludePath)
            }

        if ($foundItems) {
            foreach ($match in $foundItems) {
                $nameMatches += [PSCustomObject]@{
                    FullName   = $match.FullName
                    Name       = $match.Name
                    Type       = if ($match.PSIsContainer) { "Directory" } else { "File" }
                    SearchTerm = $searchTerm
                }
            }
        }
    }

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
        $contentMatches = Get-ChildItem -Path $StartPath -Recurse -Force -File |
            Where-Object {
                $_.Extension -notmatch '\.(exe|dll|zip|png|jpg|jpeg|gif|pdf|doc|docx|xls|xlsx)$' -and
                -not ($_.DirectoryName -eq $PSScriptRoot -and $_.Extension -eq '.log') -and
                -not (Test-ExcludedPath -ItemPath $_.FullName -ExcludePath $ExcludePath)
            } |
            ForEach-Object {
                $file = $_
                $lineNumber = 1

                # Get file content once to avoid multiple reads for performance
                $fileContent = Get-Content $file.FullName -ErrorAction SilentlyContinue

                if ($fileContent) {
                    foreach ($searchTerm in $SearchReplacePairs.Keys) {
                        $lineNumber = 1
                        # Escape the search term for regex pattern matching
                        $escapedSearchTerm = [regex]::Escape($searchTerm)
                        foreach ($line in $fileContent) {
                            if ($line -match $escapedSearchTerm) {
                                [PSCustomObject]@{
                                    File       = $file.FullName
                                    LineNumber = $lineNumber
                                    Line       = $line
                                    SearchTerm = $searchTerm
                                }
                            }
                            $lineNumber++
                        }
                    }
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
        # Determine if we need to prompt or auto-replace
        $performReplacements = $AutoReplace

        if (-not $AutoReplace) {
            foreach ($searchTerm in $SearchReplacePairs.Keys) {
                $replaceWith = $SearchReplacePairs[$searchTerm]
                if ([string]::IsNullOrEmpty($replaceWith)) {
                    $confirmation = Read-Host "`nWould you like to replace all instances of '$searchTerm'? (Y/N)"
                    if ($confirmation -eq 'Y') {
                        $replaceWith = Read-Host "Enter the replacement string"
                        $SearchReplacePairs[$searchTerm] = $replaceWith
                        $performReplacements = $true
                    }
                } else {
                    $confirmation = Read-Host "`nWould you like to replace all instances of '$searchTerm' with '$replaceWith'? (Y/N)"
                    $performReplacements = $confirmation -eq 'Y'
                }

                if (-not $performReplacements) {
                    break
                }
            }
        }

        if ($performReplacements) {
            Write-ColorOutput "`nPerforming replacements..." -ForegroundColor Cyan

            # Process each search term
            foreach ($searchTerm in $SearchReplacePairs.Keys) {
                $replaceWith = $SearchReplacePairs[$searchTerm]

                # Skip if replacement is empty and this is auto-replace mode
                if ([string]::IsNullOrEmpty($replaceWith) -and $AutoReplace) {
                    Write-ColorOutput "Skipping '$searchTerm' as no replacement was provided." -ForegroundColor Yellow
                    continue
                }

                Write-ColorOutput "Replacing '$searchTerm' with '$replaceWith'..." -ForegroundColor White

                # Replace in file contents
                $termContentMatches = $contentMatches | Where-Object { $_.SearchTerm -eq $searchTerm }
                if ($termContentMatches) {
                    $termContentMatches | Select-Object -ExpandProperty File -Unique | ForEach-Object {
                        $filePath = $_
                        try {
                            (Get-Content $filePath) |
                                ForEach-Object { $_ -replace [regex]::Escape($searchTerm), $replaceWith } |
                                Set-Content $filePath
                                Write-ColorOutput "Updated content in: $filePath" -ForegroundColor Green
                            } catch {
                                Write-ColorOutput "Error updating $filePath : $_" -ForegroundColor Red
                            }
                        }
                    }

                    # Rename files and folders
                    $termNameMatches = $nameMatches | Where-Object { $_.SearchTerm -eq $searchTerm }
                    if ($termNameMatches) {
                        $termNameMatches | ForEach-Object {
                            try {
                                $newName = $_.Name -replace [regex]::Escape($searchTerm), $replaceWith
                                $newPath = Join-Path (Split-Path $_.FullName -Parent) $newName
                                Rename-Item -Path $_.FullName -NewName $newName -ErrorAction Stop
                                Write-ColorOutput "Renamed: $($_.FullName) to $newPath" -ForegroundColor Green
                            } catch {
                                Write-ColorOutput "Error renaming $($_.FullName): $_" -ForegroundColor Red
                            }
                        }
                    }
                }
                Write-ColorOutput "`nReplacement operation completed." -ForegroundColor Cyan
            } else {
                Write-ColorOutput "`nReplacement operation cancelled." -ForegroundColor Yellow
            }
        }
    } catch {
        Write-ColorOutput "Error: $_" -ForegroundColor Red
    } finally {
        if ($logFile) {
            Stop-Transcript
        }
    }
