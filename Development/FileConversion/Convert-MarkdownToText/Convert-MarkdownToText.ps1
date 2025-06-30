# =============================================================================
# Script: Convert-MarkdownToText.ps1
# Created: 2025-04-28 23:45:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-28 23:45:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.0
# Additional Info: Initial script creation
# =============================================================================

<#
.SYNOPSIS
Converts Markdown files to plain text files.

.DESCRIPTION
This script converts Markdown files to plain text files by:
1. Removing Markdown formatting syntax (headings, links, emphasis, etc.)
2. Extracting and simplifying content
3. Preserving the content hierarchy and structure
4. Writing the result to a text file with the same name but .txt extension

The script supports individual file conversion or batch processing of directories.
It includes -WhatIf support for all file creation/modification operations.

.PARAMETER Path
Path to the Markdown file or directory containing Markdown files to convert.
If a directory is specified, all .md files in the directory will be processed.

.PARAMETER Recurse
When specified, processes Markdown files in subdirectories as well.
Only applicable when Path points to a directory.

.PARAMETER Destination
Optional destination directory for the output text files.
If not specified, text files will be created in the same location as the source files.

.PARAMETER Force
When specified, overwrites existing text files without prompting.

.EXAMPLE
.\Convert-MarkdownToText.ps1 -Path "C:\Documents\README.md"
Converts a single Markdown file to a text file in the same directory.

.EXAMPLE
.\Convert-MarkdownToText.ps1 -Path "C:\Documents" -Recurse -Destination "C:\TextFiles"
Converts all Markdown files in the Documents directory and its subdirectories,
saving the resulting text files to the TextFiles directory.

.EXAMPLE
.\Convert-MarkdownToText.ps1 -Path "C:\Documents" -WhatIf
Shows what would happen if the script was run but does not actually convert any files.
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Path,

    [Parameter(Mandatory=$false)]
    [switch]$Recurse,

    [Parameter(Mandatory=$false)]
    [string]$Destination,

    [Parameter(Mandatory=$false)]
    [switch]$Force
)

# Function to convert a single Markdown file to text
function ConvertFrom-Markdown {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$MarkdownFile,

        [Parameter(Mandatory=$false)]
        [string]$OutputDirectory,

        [Parameter(Mandatory=$false)]
        [switch]$Force
    )

    try {
        Write-Host "Processing $MarkdownFile..." -ForegroundColor Cyan

        # Validate the input file
        if (-not (Test-Path -Path $MarkdownFile -PathType Leaf)) {
            Write-Host "Error: File not found - $MarkdownFile" -ForegroundColor Red
            return $false
        }

        # Determine the output path
        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($MarkdownFile)
        if ([string]::IsNullOrEmpty($OutputDirectory)) {
            $outputPath = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($MarkdownFile), "$fileName.txt")
        } else {
            $outputPath = [System.IO.Path]::Combine($OutputDirectory, "$fileName.txt")
        }

        # Check if output file already exists
        if ((Test-Path -Path $outputPath) -and -not $Force) {
            $response = Read-Host "Output file $outputPath already exists. Overwrite? (Y/N)"
            if ($response -ne 'Y') {
                Write-Host "Skipping file: $MarkdownFile" -ForegroundColor Yellow
                return $false
            }
        }

        # Read the Markdown content
        $content = Get-Content -Path $MarkdownFile -Raw

        # Apply Markdown to text conversion rules
        $plainText = $content

        # Remove HTML tags
        $plainText = [System.Text.RegularExpressions.Regex]::Replace($plainText, '<[^>]+>', '')

        # Convert headers to plain text with emphasis
        $plainText = [System.Text.RegularExpressions.Regex]::Replace($plainText, '#{1,6}\s*(.+)$', '$1', [System.Text.RegularExpressions.RegexOptions]::Multiline)

        # Convert links [text](url) to just text
        $plainText = [System.Text.RegularExpressions.Regex]::Replace($plainText, '\[([^\]]+)\]\([^\)]+\)', '$1')

        # Convert image notations to just alt text
        $plainText = [System.Text.RegularExpressions.Regex]::Replace($plainText, '!\[([^\]]+)\]\([^\)]+\)', '$1')

        # Remove bold and italic formatting
        $plainText = [System.Text.RegularExpressions.Regex]::Replace($plainText, '\*\*(.+?)\*\*', '$1')
        $plainText = [System.Text.RegularExpressions.Regex]::Replace($plainText, '__(.+?)__', '$1')
        $plainText = [System.Text.RegularExpressions.Regex]::Replace($plainText, '\*(.+?)\*', '$1')
        $plainText = [System.Text.RegularExpressions.Regex]::Replace($plainText, '_(.+?)_', '$1')

        # Convert code blocks, both inline and multi-line
        $plainText = [System.Text.RegularExpressions.Regex]::Replace($plainText, '```[\s\S]*?```', '')
        $plainText = [System.Text.RegularExpressions.Regex]::Replace($plainText, '`([^`]+)`', '$1')

        # Replace horizontal rules with a line of dashes
        $plainText = [System.Text.RegularExpressions.Regex]::Replace($plainText, '^\s*[-*_]{3,}\s*$', '-------------------------------------------', [System.Text.RegularExpressions.RegexOptions]::Multiline)

        # Convert bullet lists to plain text with dashes
        $plainText = [System.Text.RegularExpressions.Regex]::Replace($plainText, '^\s*[-*+]\s+(.+)$', '- $1', [System.Text.RegularExpressions.RegexOptions]::Multiline)

        # Convert numbered lists to plain text
        $plainText = [System.Text.RegularExpressions.Regex]::Replace($plainText, '^\s*\d+\.\s+(.+)$', '$1', [System.Text.RegularExpressions.RegexOptions]::Multiline)

        # Fix extra blank lines
        $plainText = [System.Text.RegularExpressions.Regex]::Replace($plainText, '\n{3,}', "`n`n")

        # Ensure output directory exists
        if (-not [string]::IsNullOrEmpty($OutputDirectory)) {
            if (-not (Test-Path -Path $OutputDirectory)) {
                if ($PSCmdlet.ShouldProcess($OutputDirectory, "Create Directory")) {
                    New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
                    Write-Host "Created output directory: $OutputDirectory" -ForegroundColor Green
                }
            }
        }

        # Write the plain text to the output file
        if ($PSCmdlet.ShouldProcess($outputPath, "Create Text File")) {
            Set-Content -Path $outputPath -Value $plainText
            Write-Host "Converted to text file: $outputPath" -ForegroundColor Green
            return $true
        }

        return $true
    }
    catch {
        Write-Host "Error processing file $MarkdownFile : $_" -ForegroundColor Red
        return $false
    }
}

# Main script execution
try {
    # Keep track of statistics
    $fileCount = 0
    $successCount = 0
    $errorCount = 0

    # Define the file extension filter for markdown files
    $mdExtensions = @("*.md", "*.markdown")

    # Process single file or directory
    if (Test-Path -Path $Path -PathType Container) {
        # It's a directory, process all Markdown files
        Write-Host "Processing directory: $Path" -ForegroundColor Cyan

        foreach ($extension in $mdExtensions) {
            $mdFiles = Get-ChildItem -Path $Path -Filter $extension -Recurse:$Recurse
            foreach ($file in $mdFiles) {
                $fileCount++

                if (ConvertFrom-Markdown -MarkdownFile $file.FullName -OutputDirectory $Destination -Force:$Force) {
                    $successCount++
                } else {
                    $errorCount++
                }
            }
        }
    }
    elseif (Test-Path -Path $Path -PathType Leaf) {
        # It's a file, process it directly
        $fileCount++

        # Make sure it's a markdown file
        $extension = [System.IO.Path]::GetExtension($Path).ToLower()
        if ($extension -eq ".md" -or $extension -eq ".markdown") {
            if (ConvertFrom-Markdown -MarkdownFile $Path -OutputDirectory $Destination -Force:$Force) {
                $successCount++
            } else {
                $errorCount++
            }
        } else {
            Write-Host "Error: File is not a Markdown file (.md or .markdown extension required): $Path" -ForegroundColor Red
            $errorCount++
        }
    }
    else {
        Write-Host "Error: Path not found - $Path" -ForegroundColor Red
        $errorCount++
    }

    # Output summary
    Write-Host "`nConversion Summary:" -ForegroundColor Cyan
    Write-Host "-------------------" -ForegroundColor DarkGray
    Write-Host "Total files processed: $fileCount" -ForegroundColor White
    Write-Host "Successfully converted: $successCount" -ForegroundColor Green

    if ($errorCount -gt 0) {
        Write-Host "Errors encountered: $errorCount" -ForegroundColor Red
    } else {
        Write-Host "Errors encountered: 0" -ForegroundColor Green
    }

    # Return exit code based on success
    if ($errorCount -eq 0 -and $successCount -gt 0) {
        # Success
                exit 0
    } elseif ($errorCount -gt 0) {
        # Error
                exit 1
    } else {
        # No files processed
                exit 2
    }
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
    exit 1
}
