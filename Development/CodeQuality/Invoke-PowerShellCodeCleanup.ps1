# =============================================================================
# Script: Invoke-PowerShellCodeCleanup.ps1
# Created: 2025-06-23 15:30
# Author: jdyer-nuvodia
# Last Updated: 2025-06-27 21:15:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2.0
# Additional Info: Added Fix-InlineComments function to automatically correct inline comment issues
# =============================================================================

<#
.SYNOPSIS
Performs code quality cleanup and analysis on PowerShell scripts in the current working directory.

.DESCRIPTION
This script provides four main functions for PowerShell code quality management:
1. Clear-TrailingWhitespace: Removes trailing whitespace from PowerShell script files
2. Find-NewlineError: Identifies common newline-related formatting issues in PowerShell scripts
3. Find-PotentialMissingCode: Detects ellipses patterns that may indicate incomplete code blocks
4. Repair-InlineComment: Automatically fixes inline comment issues by separating them into proper comment lines

The script automatically targets the first .ps1 file found in the current working directory.
Both functions help maintain consistent code formatting and identify potential parsing issues.

.PARAMETER ScriptPath
The path to the PowerShell script file to process. If not specified, the script will automatically
find the first .ps1 file in the current working directory.

.EXAMPLE
.\Invoke-PowerShellCodeCleanup.ps1
Processes the first .ps1 file found in the current directory for whitespace cleanup and newline error detection.

.EXAMPLE
Clear-TrailingWhitespace -ScriptPath "C:\Scripts\MyScript.ps1"
Removes trailing whitespace from the specified script file.

.EXAMPLE
Repair-InlineComment -ScriptPath "C:\Scripts\MyScript.ps1"
Automatically fixes inline comment issues by separating them into proper comment lines above the code.
#>

[CmdletBinding()]
param()

function Clear-TrailingWhitespace {
    <#
    .SYNOPSIS
    Removes trailing whitespace from a PowerShell script file.

    .DESCRIPTION
    Reads the specified PowerShell script file, trims trailing whitespace from each line,
    and saves the cleaned content back to the file. This helps maintain consistent
    code formatting and prevents issues with version control systems.

    .PARAMETER ScriptPath
    The path to the PowerShell script file to clean. If not specified, uses the first
    .ps1 file found in the current working directory.

    .EXAMPLE
    Clear-TrailingWhitespace -ScriptPath "C:\Scripts\MyScript.ps1"
    Removes trailing whitespace from MyScript.ps1

    .EXAMPLE
    Clear-TrailingWhitespace
    Removes trailing whitespace from the first .ps1 file in the current directory
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ScriptPath
    )

    try {
        # If no script path provided, find the first .ps1 file in current directory
        if (-not $ScriptPath) {
            $ScriptPath = Get-ChildItem -Path (Get-Location) -Filter "*.ps1" | Select-Object -First 1 -ExpandProperty FullName

            if (-not $ScriptPath) {
                Write-Error -Message "No PowerShell script files (.ps1) found in the current directory."
                return
            }

            Write-Information -MessageData "Processing file: $ScriptPath" -InformationAction Continue
        }

        # Verify the file exists
        if (-not (Test-Path -Path $ScriptPath)) {
            Write-Error -Message "Script file not found: $ScriptPath"
            return
        }

        Write-Information -MessageData "Clearing trailing whitespace from: $ScriptPath" -InformationAction Continue

        # Read content, trim trailing whitespace, and save back
        (Get-Content -Path $ScriptPath) | ForEach-Object { $_.TrimEnd() } | Set-Content -Path $ScriptPath

        Write-Information -MessageData "Successfully cleared trailing whitespace from: $ScriptPath" -InformationAction Continue
    }
    catch {
        Write-Error -Message "Failed to clear trailing whitespace: $($_.Exception.Message)"
    }
}

function Find-NewlineError {
    <#
    .SYNOPSIS
    Identifies common newline-related formatting issues in PowerShell scripts.

    .DESCRIPTION
    Searches for patterns that commonly indicate newline or formatting issues in PowerShell scripts:
    - Multiple consecutive spaces not followed by comments
    - Inline comments that appear after code on the same line (which should be avoided)
    - Excludes properly formatted comment blocks and comment lines

    .PARAMETER ScriptPath
    The path to the PowerShell script file to analyze. If not specified, uses the first
    .ps1 file found in the current working directory.

    .EXAMPLE
    Find-NewlineErrors -ScriptPath "C:\Scripts\MyScript.ps1"
    Analyzes MyScript.ps1 for newline-related formatting issues

    .EXAMPLE
    Find-NewlineError
    Analyzes the first .ps1 file in the current directory for formatting issues
    #>
    [CmdletBinding()]
        param(
        [Parameter(Mandatory = $false)]
        [string]$ScriptPath
    )

    try {
        # If no script path provided, find the first .ps1 file in current directory
        if (-not $ScriptPath) {
            $ScriptPath = Get-ChildItem -Path (Get-Location) -Filter "*.ps1" | Select-Object -First 1 -ExpandProperty FullName

            if (-not $ScriptPath) {
                Write-Error -Message "No PowerShell script files (.ps1) found in the current directory."
                return
            }

            Write-Information -MessageData "Processing file: $ScriptPath" -InformationAction Continue
        }

        # Verify the file exists
        if (-not (Test-Path -Path $ScriptPath)) {
            Write-Error -Message "Script file not found: $ScriptPath"
            return
        }

        Write-Information -MessageData "Searching for newline errors in: $ScriptPath" -InformationAction Continue

        # Search for the specified patterns
        # Pattern 1: Multiple consecutive spaces not followed by comments
        # Pattern 2: Inline comments after code (but exclude lines that start with whitespace + comment markers)
        # Also exclude lines that contain quoted strings with hash symbols to avoid false positives
        $results = Select-String -Path $ScriptPath -Pattern @(
            '\S {2,}(?!#)\S',
            "^(?!\s*<#)(?!\s*#>)(?!\s*#)(?!.*'.*#.*')(?!.*`".*#.*`").*\S.*#"
        )

        if ($results) {
            Write-Information -MessageData "Found potential newline/formatting issues:" -InformationAction Continue
        }
        else {
            Write-Information -MessageData "No newline errors detected in: $ScriptPath" -InformationAction Continue
        }

        return $results
    }
    catch {
        Write-Error -Message "Failed to find newline errors: $($_.Exception.Message)"
    }
}

function Find-PotentialMissingCode {
    <#
    .SYNOPSIS
    Detects ellipses patterns that may indicate incomplete code blocks in PowerShell scripts.

    .DESCRIPTION
    Searches for patterns of three or more consecutive dots (...) which commonly indicate
    placeholder text or incomplete code blocks that need to be filled in. This is particularly
    useful for identifying template code or documentation examples that need implementation.

    .PARAMETER ScriptPath
    The path to the PowerShell script file to analyze. If not specified, uses the first
    .ps1 file found in the current working directory.

    .EXAMPLE
    Find-PotentialMissingCode -ScriptPath "C:\Scripts\MyScript.ps1"
    Analyzes MyScript.ps1 for ellipses patterns that may indicate incomplete code

    .EXAMPLE
    Find-PotentialMissingCode
    Analyzes the first .ps1 file in the current directory for potential missing code blocks
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ScriptPath
    )

    try {
        # If no script path provided, find the first .ps1 file in current directory
        if (-not $ScriptPath) {
            $ScriptPath = Get-ChildItem -Path (Get-Location) -Filter "*.ps1" | Select-Object -First 1 -ExpandProperty FullName

            if (-not $ScriptPath) {
                Write-Error -Message "No PowerShell script files (.ps1) found in the current directory."
                return
            }

            Write-Information -MessageData "Processing file: $ScriptPath" -InformationAction Continue
        }

        # Verify the file exists
        if (-not (Test-Path -Path $ScriptPath)) {
            Write-Error -Message "Script file not found: $ScriptPath"
            return
        }

        Write-Information -MessageData "Searching for potential missing code patterns in: $ScriptPath" -InformationAction Continue

        # Search for ellipses patterns that may indicate incomplete code
        $results = Select-String -Path $ScriptPath -Pattern '\.{3,}'

        if ($results) {
            Write-Information -MessageData "Found potential missing code indicators (ellipses):" -InformationAction Continue
        }
        else {
            Write-Information -MessageData "No potential missing code patterns detected in: $ScriptPath" -InformationAction Continue
        }

        return $results
    }
    catch {
        Write-Error -Message "Failed to find potential missing code patterns: $($_.Exception.Message)"
    }
}

function Repair-InlineComment {
    <#
    .SYNOPSIS
    Automatically fixes inline comment issues by separating them into proper comment lines.

    .DESCRIPTION
    Identifies lines with inline comments (comments that appear after code on the same line)
    and reformats them by placing the comment on a separate line above the code. This helps
    maintain clean code formatting and follows PowerShell best practices.

    .PARAMETER ScriptPath
    The path to the PowerShell script file to fix. If not specified, uses the first
    .ps1 file found in the current working directory.

    .EXAMPLE
    Repair-InlineComment -ScriptPath "C:\Scripts\MyScript.ps1"
    Fixes inline comments in the specified script file

    .EXAMPLE
    Repair-InlineComment
    Fixes inline comments in the first .ps1 file in the current directory
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ScriptPath
    )

    try {
        # If no script path provided, find the first .ps1 file in current directory
        if (-not $ScriptPath) {
            $ScriptPath = Get-ChildItem -Path (Get-Location) -Filter "*.ps1" | Select-Object -First 1 -ExpandProperty FullName

            if (-not $ScriptPath) {
                Write-Error -Message "No PowerShell script files (.ps1) found in the current directory."
                return
            }

            Write-Information -MessageData "Processing file: $ScriptPath" -InformationAction Continue
        }

        # Verify the file exists
        if (-not (Test-Path -Path $ScriptPath)) {
            Write-Error -Message "Script file not found: $ScriptPath"
            return
        }

        Write-Information -MessageData "Fixing inline comments in: $ScriptPath" -InformationAction Continue

        # Read the file content
        $content = Get-Content -Path $ScriptPath
        $fixedContent = @()
        $fixCount = 0

        foreach ($line in $content) {
            # Check if line has inline comment (excluding properly formatted comments and quoted strings)
            if ($line -match "^(?!\s*<#)(?!\s*#>)(?!\s*#)(?!.*'.*#.*')(?!.*`".*#.*`").*\S.*#(.*)$") {
                # Extract the code part and comment part
                $parts = $line -split '#', 2
                $codePart = $parts[0].TrimEnd()
                $commentPart = $parts[1].Trim()

                # Get the indentation from the original line
                $indentation = ""
                if ($line -match "^(\s*)") {
                    $indentation = $matches[1]
                }

                # Add the comment line first, then the code line
                $fixedContent += "$indentation# $commentPart"
                $fixedContent += "$indentation$codePart"
                $fixCount++

                Write-Information -MessageData "Fixed inline comment: $line" -InformationAction Continue
            }
            else {
                # Keep the line as-is
                $fixedContent += $line
            }
        }

        if ($fixCount -gt 0) {
            # Write the fixed content back to the file
            $fixedContent | Set-Content -Path $ScriptPath
            Write-Information -MessageData "Fixed $fixCount inline comment(s) in: $ScriptPath" -InformationAction Continue

            # Clean up trailing whitespace after the fixes
            Write-Information -MessageData "Cleaning up trailing whitespace after fixes..." -InformationAction Continue
            Clear-TrailingWhitespace -ScriptPath $ScriptPath
        }
        else {
            Write-Information -MessageData "No inline comments found to fix in: $ScriptPath" -InformationAction Continue
        }

        return $fixCount
    }
    catch {
        Write-Error -Message "Failed to fix inline comments: $($_.Exception.Message)"
    }
}

# Main execution block
if ($MyInvocation.InvocationName -ne '.') {
    # Get the first .ps1 file in the current working directory
    $scriptPath = Get-ChildItem -Path (Get-Location) -Filter "*.ps1" | Select-Object -First 1 -ExpandProperty FullName

    if ($scriptPath) {
        Write-Information -MessageData "Found PowerShell script: $scriptPath" -InformationAction Continue

        # Execute all functions
        Clear-TrailingWhitespace -ScriptPath $scriptPath
        $inlineCommentResults = Find-NewlineError -ScriptPath $scriptPath
        Find-PotentialMissingCode -ScriptPath $scriptPath

        # Automatically fix inline comments if found
        if ($inlineCommentResults) {
            Write-Information -MessageData "Inline comments detected. Automatically fixing..." -InformationAction Continue
            Repair-InlineComment -ScriptPath $scriptPath
            Write-Information -MessageData "Re-running analysis after fixes..." -InformationAction Continue
            Find-NewlineError -ScriptPath $scriptPath
        }
        Repair-InlineComment -ScriptPath $scriptPath
    }
    else {
        Write-Warning -Message "No PowerShell script files (.ps1) found in the current directory."
    }
}
