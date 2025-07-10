# =============================================================================
# Script: Invoke-PowerShellCodeCleanup.ps1
# Created: 2025-06-23 15:30
# Author: jdyer-nuvodia
# Last Updated: 2025-07-08 21:15:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.6.3
# Additional Info: Fixed regex pattern to exclude string interpolation from brace spacing fixes
# =============================================================================

<#
.SYNOPSIS
Performs code quality cleanup and analysis on PowerShell scripts.

This script provides six main functions for PowerShell code quality management:
1. Clear-TrailingWhitespace: Removes trailing whitespace from PowerShell script files
2. Find-NewlineError: Identifies common newline-related formatting issues in PowerShell scripts
3. Find-PotentialMissingCode: Detects ellipses patterns that may indicate incomplete code blocks
4. Repair-InlineComment: Automatically fixes inline comment issues by separating them into proper comment lines
5. Repair-CloseBraceNewline: Fixes PSPlaceCloseBrace issues where closing braces are followed by newlines instead of being on the same line as branch statements
6. Repair-WhitespaceConsistency: Fixes PSUseConsistentWhitespace violations for operators, braces, and commas

.DESCRIPTION
The script can target a specific file via the ScriptPath parameter or automatically target the first .ps1 file found in the current working directory.
These functions help maintain consistent code formatting and identify potential parsing issues.

.PARAMETER ScriptPath
The full path to the PowerShell script file to process. If not specified, the script will automatically
find the first .ps1 file in the current working directory.

.EXAMPLE
.\Invoke-PowerShellCodeCleanup.ps1
Processes the first .ps1 file found in the current directory for whitespace cleanup and newline error detection.

.EXAMPLE
.\Invoke-PowerShellCodeCleanup.ps1 -ScriptPath "C:\Scripts\MyScript.ps1"
Processes the specified script file for code quality cleanup and analysis.

.EXAMPLE
Clear-TrailingWhitespace -ScriptPath "C:\Scripts\MyScript.ps1"
Removes trailing whitespace from the specified script file.

.EXAMPLE
Repair-InlineComment -ScriptPath "C:\Scripts\MyScript.ps1"
Automatically fixes inline comment issues by separating them into proper comment lines above the code.

.EXAMPLE
Repair-CloseBraceNewline -ScriptPath "C:\Scripts\MyScript.ps1"
Fixes PSPlaceCloseBrace violations where closing braces are followed by newlines instead of being on the same line as branch statements.

.EXAMPLE
Repair-WhitespaceConsistency -ScriptPath "C:\Scripts\MyScript.ps1"
Fixes PSUseConsistentWhitespace violations including space before open brace issues.
#>

[CmdletBinding()]
[OutputType([int])]
param(
    [Parameter(Mandatory = $false)]
    [string]$ScriptPath
)

function Clear-TrailingWhitespace {
    <#
    .SYNOPSIS
    Removes trailing whitespace from a PowerShell script file. "   this is ok" but not trailing spaces

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
    } catch {
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
        # Pattern 1: Multiple consecutive spaces not followed by comments (excluding content inside quotes)
        # Pattern 2: Inline comments after code (but exclude lines that start with whitespace + comment markers)
        # Also exclude lines that contain quoted strings with hash symbols to avoid false positives
        $results = Select-String -Path $ScriptPath -Pattern @(
            '(?=^(?:[^"'']*"[^"]*")*[^"'']*$)(?=^(?:[^"'']*''[^'']*'')*[^"'']*$)\S{2,}(?!#)\S',
            "^(?!\s*<#)(?!\s*#>)(?!\s*#)(?!.*'.*#.*')(?!.*`".*#.*`").*\S.*#"
        )

        if ($results) {
            Write-Information -MessageData "Found potential newline/formatting issues:" -InformationAction Continue
        } else {
            Write-Information -MessageData "No newline errors detected in: $ScriptPath" -InformationAction Continue
        }

        return $results
    } catch {
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

        # Search for ellipses patterns that may indicate incomplete code (excluding content inside quotes)
        $results = Select-String -Path $ScriptPath -Pattern '(?=^(?:[^"'']*"[^"]*")*[^"'']*$)(?=^(?:[^"'']*''[^'']*'')*[^"'']*$)\.{3,}'

        if ($results) {
            Write-Information -MessageData "Found potential missing code indicators (ellipses):" -InformationAction Continue
        } else {
            Write-Information -MessageData "No potential missing code patterns detected in: $ScriptPath" -InformationAction Continue
        }

        return $results
    } catch {
        Write-Error -Message "Failed to find potential missing code patterns: $($_.Exception.Message)"
    }
}

function Repair-CloseBraceNewline {
    <#
    .SYNOPSIS
    Fixes PSPlaceCloseBrace issues where closing braces are followed by newlines instead of being on the same line as branch statements.

    .DESCRIPTION
    Identifies lines where closing braces are followed by a newline and then a branch statement
    (catch, else, elseif, finally, etc.) and reformats them to place the brace and branch statement
    on the same line. This fixes PSScriptAnalyzer PSPlaceCloseBrace violations.

    .PARAMETER ScriptPath
    The path to the PowerShell script file to fix. If not specified, uses the first
    .ps1 file found in the current working directory.

    .EXAMPLE
    Repair-CloseBraceNewline -ScriptPath "C:\Scripts\MyScript.ps1"
    Fixes close brace newline issues in the specified script file

    .EXAMPLE
    Repair-CloseBraceNewline
    Fixes close brace newline issues in the first .ps1 file in the current directory
    #>
    [CmdletBinding()]
    [OutputType([int])]
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

        Write-Information -MessageData "Fixing close brace newline issues in: $ScriptPath" -InformationAction Continue

        # Read the file content
        $content = Get-Content -Path $ScriptPath
        $fixedContent = @()
        $fixCount = 0
        $i = 0

        while ($i -lt $content.Length) {
            $currentLine = $content[$i]

            # Check if current line ends with a closing brace and the next line starts with a branch statement
            if ($i -lt ($content.Length - 1)) {
                $nextLine = $content[$i + 1]

                # Pattern: line ending with } (possibly with whitespace after)
                # Next line: starts with whitespace and then branch keywords
                if ($currentLine -match "^(\s*)}(\s*)$") {
                    $braceIndentation = $matches[1]

                    if ($nextLine -match "^(\s*)(catch|else|elseif|finally)(\s|\{)") {
                        $nextIndentation = $matches[1]
                        $branchKeyword = $matches[2]
                        $branchRemainder = $nextLine.Substring($nextIndentation.Length + $branchKeyword.Length)

                        # Create the fixed line by combining brace and branch statement
                        $fixedLine = "$braceIndentation} $branchKeyword$branchRemainder"
                        $fixedContent += $fixedLine

                        Write-Information -MessageData "Fixed close brace newline: '$currentLine' + '$nextLine' -> '$fixedLine'" -InformationAction Continue

                        $fixCount++
                        $i += 2
                        continue
                    }
                }
            }

            # No fix needed, keep the line as-is
            $fixedContent += $currentLine
            $i++
        }

        if ($fixCount -gt 0) {
            # Write the fixed content back to the file
            $fixedContent | Set-Content -Path $ScriptPath
            Write-Information -MessageData "Fixed $fixCount close brace newline issue(s) in: $ScriptPath" -InformationAction Continue

            # Clean up trailing whitespace after the fixes
            Write-Information -MessageData "Cleaning up trailing whitespace after fixes..." -InformationAction Continue
            Clear-TrailingWhitespace -ScriptPath $ScriptPath
        } else {
            Write-Information -MessageData "No close brace newline issues found to fix in: $ScriptPath" -InformationAction Continue
        }

        return $fixCount
    } catch {
        Write-Error -Message "Failed to fix close brace newline issues: $($_.Exception.Message)"
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
    [OutputType([int])]
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
            # Check if line has inline comment that should be moved to its own line
            # This is more restrictive - only move comments that appear to be substantive documentation
            # Exclude: assignment comments, short explanatory comments, quotes with #, existing comment lines
            if ($line -match "^(?!\s*<#)(?!\s*#>)(?!\s*#)(?!.*'.*#.*')(?!.*`".*#.*`").*\S\s+#\s*([A-Z].*|[a-z]{4,}.*)$") {
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
            } else {
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
        } else {
            Write-Information -MessageData "No inline comments found to fix in: $ScriptPath" -InformationAction Continue
        }

        return $fixCount
    } catch {
        Write-Error -Message "Failed to fix inline comments: $($_.Exception.Message)"
    }
}

function Repair-WhitespaceConsistency {
    <#
    .SYNOPSIS
    Fixes PSUseConsistentWhitespace violations in PowerShell scripts.

    .DESCRIPTION
    Identifies and fixes common whitespace consistency issues that violate PSScriptAnalyzer PSUseConsistentWhitespace rules:
    - Ensures exactly one space before and after binary and assignment operators (=, +=, -=, *=, /=, %=, -eq, -ne, -lt, -gt, -le, -ge, -like, -notlike, -match, -notmatch, -contains, -notcontains, -in, -notin, -and, -or, -xor, -not, +, -, *, /, %)
    - Ensures exactly one space before opening braces { (e.g., switch statements, function definitions)
    - Ensures exactly one space after opening braces { and before closing braces }
    - Ensures exactly one space after commas
    - Preserves existing proper spacing and avoids over-correcting

    .PARAMETER ScriptPath
    The path to the PowerShell script file to fix. If not specified, uses the first
    .ps1 file found in the current working directory.

    .EXAMPLE
    Repair-WhitespaceConsistency -ScriptPath "C:\Scripts\MyScript.ps1"
    Fixes whitespace consistency issues in the specified script file

    .EXAMPLE
    Repair-WhitespaceConsistency
    Fixes whitespace consistency issues in the first .ps1 file in the current directory
    #>
    [CmdletBinding()]
    [OutputType([int])]
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

        Write-Information -MessageData "Fixing whitespace consistency issues in: $ScriptPath" -InformationAction Continue

        # Read the file content
        $content = Get-Content -Path $ScriptPath
        $fixedContent = @()
        $fixCount = 0

        foreach ($line in $content) {
            $originalLine = $line
            $modifiedLine = $line

            # Skip comment lines and empty lines
            if ($line -match '^\s*#' -or $line -match '^\s*$') {
                $fixedContent += $line
                continue
            }

            # Fix binary and assignment operators - ensure exactly one space before and after =
            # Handle assignment operators like = += -= *= /= %=
            $modifiedLine = $modifiedLine -replace '(\w)\s*=\s*([^=])', '$1 = $2'
            $modifiedLine = $modifiedLine -replace '(\w)\s*\+=\s*(\w)', '$1 += $2'
            $modifiedLine = $modifiedLine -replace '(\w)\s*-=\s*(\w)', '$1 -= $2'
            $modifiedLine = $modifiedLine -replace '(\w)\s*\*=\s*(\w)', '$1 *= $2'
            $modifiedLine = $modifiedLine -replace '(\w)\s*/=\s*(\w)', '$1 /= $2'
            $modifiedLine = $modifiedLine -replace '(\w)\s*%=\s*(\w)', '$1 %= $2'

            # Fix braces - ensure exactly one space before and after opening brace
            # Fix missing space before opening brace (e.g., "switch ($var){" -> "switch ($var) {")
            $modifiedLine = $modifiedLine -replace '(\w|\)|")\{', '$1 {'
            # Fix multiple spaces before opening brace (e.g., "item"    { -> "item" {)
            $modifiedLine = $modifiedLine -replace '\s{2,}\{', ' {'
            # Fix exactly one space after opening brace, but exclude string interpolation patterns
            # Only match standalone braces, not those preceded by $ (string interpolation)
            $modifiedLine = $modifiedLine -replace '(?<!\$)\{([^\s\}])', '{ $1'

            # Fix commas - ensure exactly one space after comma
            $modifiedLine = $modifiedLine -replace ',([^\s\r\n])', ', $1'

            if ($modifiedLine -ne $originalLine) {
                $fixCount++
                Write-Information -MessageData "Fixed whitespace: '$originalLine' -> '$modifiedLine'" -InformationAction Continue
            }

            $fixedContent += $modifiedLine
        }

        # Write the fixed content back to the file if changes were made
        if ($fixCount -gt 0) {
            $fixedContent | Set-Content -Path $ScriptPath
            Write-Information -MessageData "Fixed $fixCount whitespace consistency issue(s) in: $ScriptPath" -InformationAction Continue

            # Clean up trailing whitespace after the fixes
            Write-Information -MessageData "Cleaning up trailing whitespace after fixes..." -InformationAction Continue
            Clear-TrailingWhitespace -ScriptPath $ScriptPath
        } else {
            Write-Information -MessageData "No whitespace consistency issues found to fix in: $ScriptPath" -InformationAction Continue
        }

        return $fixCount
    } catch {
        Write-Error -Message "Failed to fix whitespace consistency: $($_.Exception.Message)"
    }
}

# Main execution block
if ($MyInvocation.InvocationName -ne '.') {
    # Determine the target script path
    if ($ScriptPath) {
        # Use the provided script path
        if (-not (Test-Path -Path $ScriptPath)) {
            Write-Error -Message "Specified script file not found: $ScriptPath"
            exit 1
        }

        if (-not ($ScriptPath -like "*.ps1")) {
            Write-Error -Message "Specified file is not a PowerShell script (.ps1): $ScriptPath"
            exit 1
        }

        $targetScriptPath = $ScriptPath
        Write-Information -MessageData "Using specified PowerShell script: $targetScriptPath" -InformationAction Continue
    } else {
        # Get the first .ps1 file in the current working directory
        $targetScriptPath = Get-ChildItem -Path (Get-Location) -Filter "*.ps1" | Select-Object -First 1 -ExpandProperty FullName

        if (-not $targetScriptPath) {
            Write-Warning -Message "No PowerShell script files (.ps1) found in the current directory."
            exit 1
        }

        Write-Information -MessageData "Found PowerShell script in current directory: $targetScriptPath" -InformationAction Continue
    }

    # Execute all functions on the target script
    Clear-TrailingWhitespace -ScriptPath $targetScriptPath
    $inlineCommentResults = Find-NewlineError -ScriptPath $targetScriptPath
    Find-PotentialMissingCode -ScriptPath $targetScriptPath

    # Automatically fix close brace newline issues
    Write-Information -MessageData "Checking for close brace newline issues..." -InformationAction Continue
    Repair-CloseBraceNewline -ScriptPath $targetScriptPath

    # Automatically fix whitespace consistency issues
    Write-Information -MessageData "Checking for whitespace consistency issues..." -InformationAction Continue
    Repair-WhitespaceConsistency -ScriptPath $targetScriptPath

    # Automatically fix inline comments if found
    if ($inlineCommentResults) {
        Write-Information -MessageData "Inline comments detected. Automatically fixing..." -InformationAction Continue
        Repair-InlineComment -ScriptPath $targetScriptPath
        Write-Information -MessageData "Re-running analysis after fixes..." -InformationAction Continue
        Find-NewlineError -ScriptPath $targetScriptPath
    }
    Repair-InlineComment -ScriptPath $targetScriptPath
}
