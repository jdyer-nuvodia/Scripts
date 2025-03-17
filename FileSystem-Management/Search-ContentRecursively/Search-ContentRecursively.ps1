# =============================================================================
# Script: Search-ContentRecursively.ps1
# Created: 2025-03-17 21:00:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-17 21:00:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.0
# Additional Info: Initial script creation for recursive content searching
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
Searches for "ConfigMgr" in all files and directories under C:\Scripts

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

function Write-ColorOutput {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $true)]
        [string]$ForegroundColor
    )
    
    Write-Host $Message -ForegroundColor $ForegroundColor
}

# Validate the start path
if (-not (Test-Path -Path $StartPath)) {
    Write-ColorOutput "Error: The specified path '$StartPath' does not exist." -ForegroundColor Red
    exit 1
}

Write-ColorOutput "Starting search for keyword '$Keyword' in path '$StartPath'..." -ForegroundColor Cyan

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
        Where-Object { $_.Extension -notmatch '\.(exe|dll|zip|png|jpg|jpeg|gif|pdf|doc|docx|xls|xlsx)$' } |
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