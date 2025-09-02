# =============================================================================
# Script: Rename-FolderCase.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.4.2
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
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

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true,
        Position = 0,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
    [string]$StartPath,

    [Parameter(Mandatory = $false)]
    [switch]$Recursive
)

begin {
    # Color support variables and Write-ColorOutput function
    $Script:UseAnsiColors = $PSVersionTable.PSVersion.Major -ge 7
    $Script:Colors = if ($Script:UseAnsiColors) {
        @{
            'White' = "`e[37m"
            'Cyan' = "`e[36m"
            'Green' = "`e[32m"
            'Yellow' = "`e[33m"
            'Red' = "`e[31m"
            'Magenta' = "`e[35m"
            'DarkGray' = "`e[90m"
            'Reset' = "`e[0m"
        }
    } else {
        @{
            'White' = [ConsoleColor]::White
            'Cyan' = [ConsoleColor]::Cyan
            'Green' = [ConsoleColor]::Green
            'Yellow' = [ConsoleColor]::Yellow
            'Red' = [ConsoleColor]::Red
            'Magenta' = [ConsoleColor]::Magenta
            'DarkGray' = [ConsoleColor]::DarkGray
            'Reset' = ''
        }
    }

    function Write-ColorOutput {
        <#
        .SYNOPSIS
        Outputs colored text in a way that's compatible with PSScriptAnalyzer requirements.

        .DESCRIPTION
        This function provides colored output while maintaining compatibility with PSScriptAnalyzer
        by using only Write-Output and standard PowerShell cmdlets.
        #>
        param(
            [Parameter(Mandatory = $true)]
            [string]$Message,
            [Parameter(Mandatory = $false)]
            [string]$Color = "White"
        )

        # Always use Write-Output to satisfy PSScriptAnalyzer
        # For PowerShell 7+, include ANSI color codes in the output
        if ($Script:UseAnsiColors) {
            $colorCode = $Script:Colors[$Color]
            $resetCode = $Script:Colors.Reset
            Write-Output "${colorCode}${Message}${resetCode}"
        } else {
            # For PowerShell 5.1, just output the message
            # Color formatting will be handled by the terminal/host if supported
            Write-Output $Message
        }
    }

    function Convert-ToPascalCase {
        param([string]$text)

        Write-Verbose "Converting to PascalCase: $text"

        # Split by common delimiters
        $words = $text -split '[-_\s]'

        # Convert each word to proper case
        $words = $words | ForEach-Object {
            if ($_.Length -gt 0) {
                $_.Substring(0, 1).ToUpper() + $_.Substring(1).ToLower()
            }
        }

        # Rejoin with hyphens for PowerShell convention
        $result = $words -join '-'
        Write-Verbose "Converted to: $result"
        return $result
    }

    function Rename-FolderWithCase {
        [CmdletBinding(SupportsShouldProcess = $true)]
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
                Write-ColorOutput -Message "Renaming '$currentName' to '$newName'" -Color 'Cyan'
                $newPath = Join-Path -Path $parentPath -ChildPath $newName

                # Handle case where only case is different (needs temp rename)
                if ($newPath.ToLower() -eq $StartPath.ToLower()) {
                    $tempName = "_temp_" + [Guid]::NewGuid().ToString().Substring(0, 8)
                    $tempPath = Join-Path -Path $parentPath -ChildPath $tempName

                    if ($PSCmdlet.ShouldProcess($StartPath, "Rename to temp folder '$tempPath'")) {
                        Write-Verbose "Temporary rename: '$StartPath' -> '$tempPath'"
                        Rename-Item -LiteralPath $StartPath -NewName $tempName -ErrorAction Stop

                        Write-Verbose "Final rename: '$tempPath' -> '$newPath'"
                        Rename-Item -LiteralPath $tempPath -NewName $newName -ErrorAction Stop
                        Write-ColorOutput -Message "Successfully renamed: '$currentName' -> '$newName'" -Color 'Green'
                    }
                } else {
                    if ($PSCmdlet.ShouldProcess($StartPath, "Rename to '$newPath'")) {
                        Rename-Item -LiteralPath $StartPath -NewName $newName -ErrorAction Stop
                        Write-ColorOutput -Message "Successfully renamed: '$currentName' -> '$newName'" -Color 'Green'
                    }
                }
            } else {
                Write-Verbose "Skipping '$currentName' - already in correct case"
            }
        } catch {
            Write-Error "Error processing folder '$StartPath': $_"
            Write-ColorOutput -Message "Stack Trace: $($_.ScriptStackTrace)" -Color 'Red'
        }
    }

    # Enable verbose output
    $VerbosePreference = "Continue"
}

process {
    Write-ColorOutput -Message "Script started - Processing path: $StartPath" -Color 'Cyan'
    Write-Verbose "Recursive mode: $Recursive"

    try {
        Write-ColorOutput -Message "`nStarting folder case correction process..." -Color 'Cyan'

        # Verify path exists
        if (!(Test-Path -Path $StartPath)) {
            throw "Path '$StartPath' does not exist."
        }

        Write-ColorOutput -Message "Getting list of folders to process..." -Color 'Cyan'
        # Get folders to process
        $folders = @()
        if ($Recursive) {
            Write-Verbose "Getting all subfolders recursively..."
            $folders = Get-ChildItem -Path $StartPath -Directory -Recurse
            Write-ColorOutput -Message "Found $($folders.Count) subfolders" -Color 'Cyan'
        } else {
            Write-Verbose "Getting immediate subfolders only..."
            $folders = Get-ChildItem -Path $StartPath -Directory
            Write-ColorOutput -Message "Found $($folders.Count) folders" -Color 'Cyan'
        }

        # Process folders in reverse order (deepest first) to handle nested folders
        Write-ColorOutput -Message "Processing folders from deepest to shallowest..." -Color 'Cyan'
        $folders = $folders | Sort-Object -Property FullName -Descending

        # Process each folder
        foreach ($folder in $folders) {
            Write-ColorOutput -Message "`nProcessing folder: $($folder.FullName)" -Color 'Cyan'
            Rename-FolderWithCase -StartPath $folder.FullName
        }

        # Process the root folder if it's a directory
        if ((Get-Item -Path $StartPath).PSIsContainer) {
            Write-ColorOutput -Message "`nProcessing root folder: $StartPath" -Color 'Cyan'
            Rename-FolderWithCase -StartPath $StartPath
        }

        Write-ColorOutput -Message "`nFolder case correction process completed successfully." -Color 'Green'
    } catch {
        Write-Error "Script error: $_"
        Write-ColorOutput -Message "Stack Trace: $($_.ScriptStackTrace)" -Color 'Red'
        exit 1
    }
}
