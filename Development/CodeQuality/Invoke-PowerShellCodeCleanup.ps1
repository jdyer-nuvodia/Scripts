# =============================================================================
# Script: Invoke-PowerShellCodeCleanup.ps1
# Created: 2025-06-23 15:30:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-06-23 15:45:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.2
# Additional Info: Removed duplicate issue reporting, keeping only Select-String original format
# =============================================================================

<#
.SYNOPSIS
Performs code quality cleanup and analysis on PowerShell scripts in the current working directory.

.DESCRIPTION
This script provides two main functions for PowerShell code quality management:
1. Clear-TrailingWhitespace: Removes trailing whitespace from PowerShell script files
2. Find-NewlineError: Identifies common newline-related formatting issues in PowerShell scripts

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
Find-NewlineError -ScriptPath "C:\Scripts\MyScript.ps1"
Searches for newline-related formatting errors in the specified script file.
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
    - Lines ending with non-XML content followed by comments without proper spacing
    - Three or more consecutive dots (which may indicate incomplete ellipsis patterns)

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
        $results = Select-String -Path $ScriptPath -Pattern @('\S {2,}(?!#)\S', '^(?!<).*?\S.*#', '\.{3,}')

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

# Main execution block
if ($MyInvocation.InvocationName -ne '.') {
    # Get the first .ps1 file in the current working directory
    $scriptPath = Get-ChildItem -Path (Get-Location) -Filter "*.ps1" | Select-Object -First 1 -ExpandProperty FullName

    if ($scriptPath) {
        Write-Information -MessageData "Found PowerShell script: $scriptPath" -InformationAction Continue

        # Execute both functions
        Clear-TrailingWhitespace -ScriptPath $scriptPath
        Find-NewlineError -ScriptPath $scriptPath
    }
    else {
        Write-Warning -Message "No PowerShell script files (.ps1) found in the current directory."
    }
}
