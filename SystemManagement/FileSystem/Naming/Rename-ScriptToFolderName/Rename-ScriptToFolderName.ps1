# =============================================================================
# Script: Rename-ScriptToFolderName.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.3.8
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
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

    function Rename-ScriptFile {
        [CmdletBinding(SupportsShouldProcess = $true)]
        param(
            [string]$StartPath
        )

        try {
            # Get folder info to preserve exact case
            $folderInfo = Get-Item -Path $StartPath
            # This preserves the exact case
            $folderName = $folderInfo.Name

            Write-ColorOutput -Message "`nProcessing folder: $StartPath" -Color 'Cyan'
            Write-ColorOutput -Message "Folder name (exact case): $folderName" -Color 'DarkGray'

            # Get all .ps1 files in the folder
            $psFiles = Get-ChildItem -Path $StartPath -Filter "*.ps1" -File

            if ($psFiles.Count -eq 0) {
                Write-ColorOutput -Message "  No PowerShell scripts found in this folder" -Color 'Yellow'
                return
            }

            Write-ColorOutput -Message "  Found $($psFiles.Count) PowerShell script(s)" -Color 'DarkGray'

            if ($psFiles.Count -eq 1) {
                $psFile = $psFiles[0]
                $newName = "$folderName.ps1"

                # Skip if name already matches exactly (including case)
                if ($psFile.Name -ceq $newName) {
                    Write-ColorOutput -Message "  Script '$($psFile.Name)' already matches folder name and case" -Color 'Yellow'
                    return
                }

                # Show case difference if only case is different
                if ($psFile.Name -eq $newName) {
                    Write-ColorOutput -Message "  Case difference detected:" -Color 'Yellow'
                    Write-ColorOutput -Message "    Current: $($psFile.Name)" -Color 'DarkGray'
                    Write-ColorOutput -Message "    New:     $newName" -Color 'DarkGray'
                }

                # Rename the file
                if ($PSCmdlet.ShouldProcess($psFile.FullName, "Rename to '$newName'")) {
                    Write-ColorOutput -Message "  Renaming '$($psFile.Name)' to '$newName'" -Color 'DarkGray'

                    # Handle case-only changes
                    if ($psFile.Name -eq $newName) {
                        $tempName = "_temp_" + [Guid]::NewGuid().ToString().Substring(0, 8) + ".ps1"
                        Rename-Item -Path $psFile.FullName -NewName $tempName -Force
                        Rename-Item -Path (Join-Path -Path $StartPath -ChildPath $tempName) -NewName $newName -Force
                    } else {
                        Rename-Item -Path $psFile.FullName -NewName $newName -Force
                    }

                    Write-ColorOutput -Message "  Successfully renamed to '$newName'" -Color 'Green'
                }
            } else {
                # Multiple .ps1 files found - prompt for action
                Write-ColorOutput -Message "`n  Multiple PowerShell scripts found:" -Color 'Yellow'
                for ($i = 0; $i -lt $psFiles.Count; $i++) {
                    Write-ColorOutput -Message "    [$i] $($psFiles[$i].Name)" -Color "White"
                }

                $choice = Read-Host "`n  Do you want to rename one of these files to '$folderName.ps1'? (y/n)"
                if ($choice -eq 'y') {
                    $fileChoice = Read-Host "  Enter the number of the file to rename"
                    if ($fileChoice -match '^\d+$' -and [int]$fileChoice -lt $psFiles.Count) {
                        $selectedFile = $psFiles[[int]$fileChoice]
                        $newName = "$folderName.ps1"

                        if ($PSCmdlet.ShouldProcess($selectedFile.FullName, "Rename to '$newName'")) {
                            Write-ColorOutput -Message "  Renaming '$($selectedFile.Name)' to '$newName'" -Color 'DarkGray'

                            # Handle case-only changes
                            if ($selectedFile.Name -eq $newName) {
                                $tempName = "_temp_" + [Guid]::NewGuid().ToString().Substring(0, 8) + ".ps1"
                                Rename-Item -Path $selectedFile.FullName -NewName $tempName -Force
                                Rename-Item -Path (Join-Path -Path $StartPath -ChildPath $tempName) -NewName $newName -Force
                            } else {
                                Rename-Item -Path $selectedFile.FullName -NewName $newName -Force
                            }

                            Write-ColorOutput -Message "  Successfully renamed to '$newName'" -Color 'Green'
                        }
                    } else {
                        Write-Warning "Invalid file selection '$fileChoice'. Please enter a number between 0 and $($psFiles.Count - 1)"
                        Write-ColorOutput -Message "  Skipping folder: $StartPath" -Color 'Yellow'
                    }
                }
            }
        } catch {
            Write-Error "Error processing folder '$StartPath': $_"
            Write-ColorOutput -Message "Stack Trace: $($_.ScriptStackTrace)" -Color 'Red'
        }
    }
}

process {
    try {
        # Verify path exists
        if (!(Test-Path -Path $script:StartPath)) {
            throw "Path '$script:StartPath' does not exist."
        }

        Write-ColorOutput -Message "Starting script rename process..." -Color 'Cyan'
        Write-ColorOutput -Message "Processing path: $script:StartPath" -Color 'DarkGray'
        Write-ColorOutput -Message "Recursive mode: $Recursive" -Color 'DarkGray'

        # Initialize folders collection
        $folders = @()

        # Get the root folder
        if ((Get-Item -Path $script:StartPath).PSIsContainer) {
            $folders += Get-Item -Path $script:StartPath
        }

        # Add subfolders if recursive
        if ($Recursive) {
            Write-ColorOutput -Message "Getting all subfolders recursively..." -Color 'DarkGray'
            $subFolders = Get-ChildItem -Path $script:StartPath -Directory -Recurse
            $folders += $subFolders
            Write-ColorOutput -Message "Found $($subFolders.Count) subfolder(s)" -Color 'DarkGray'
        }

        # Sort folders by depth (deepest first) to handle nested folders properly
        $folders = $folders | Sort-Object { ($_.FullName -split '\\').Count } -Descending

        # Process each folder
        foreach ($folder in $folders) {
            Rename-ScriptFile -StartPath $folder.FullName
        }

        Write-ColorOutput -Message "`nScript rename process completed successfully." -Color 'Green'
    } catch {
        Write-Error "Script error: $_"
        Write-ColorOutput -Message "Stack Trace: $($_.ScriptStackTrace)" -Color 'Red'
        exit 1
    }
}
