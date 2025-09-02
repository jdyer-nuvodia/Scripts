# =============================================================================
# Script: Convert-WriteHostToColorOutput.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.4.4
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
# =============================================================================

<#
.SYNOPSIS
Converts Write-Host calls to Write-ColorOutput function calls for PSScriptAnalyzer compliance.

.DESCRIPTION
This script systematically replaces all Write-Host calls in PowerShell scripts with Write-ColorOutput
calls to ensure PSScriptAnalyzer compliance. The script handles various Write-Host patterns including:
- Write-Host with double quotes and -ForegroundColor
- Write-Host with single quotes and -ForegroundColor
- Write-Host with variables and -ForegroundColor
- Write-Host with variable color parameters (e.g., Write-Host $Message -ForegroundColor $Color)
- Write-Host with string interpolation

The script also adds the required Write-ColorOutput function and color support variables to the
target script if they do not already exist.

.PARAMETER FilePath
The path to the PowerShell script file to process. Can be a single file or multiple files.

.PARAMETER BackupOriginal
Creates a backup of the original file before making changes. Default is $true.

.PARAMETER WhatIf
Shows what changes would be made without actually modifying files.

.EXAMPLE
.\Convert-WriteHostToColorOutput.ps1 -FilePath "C:\Scripts\MyScript.ps1"
Converts Write-Host calls to Write-ColorOutput in the specified script.

.EXAMPLE
.\Convert-WriteHostToColorOutput.ps1 -FilePath "C:\Scripts\*.ps1" -BackupOriginal $false
Converts Write-Host calls in all PowerShell scripts in the directory without creating backups.

.EXAMPLE
.\Convert-WriteHostToColorOutput.ps1 -FilePath ".\MyScript.ps1" -WhatIf
Shows what changes would be made without actually modifying the file.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string[]]$FilePath,

    [Parameter(Mandatory = $false)]
    [bool]$BackupOriginal = $true
)

# Color support variables and Write-ColorOutput function template
$Script:ColorSupportCode = @'
# Color support variables and Write-ColorOutput function
$Script:UseAnsiColors = $PSVersionTable.PSVersion.Major -ge 7
$Script:Colors = if ($Script:UseAnsiColors) {
    @{
        'White'    = "`e[37m"
        'Cyan'     = "`e[36m"
        'Green'    = "`e[32m"
        'Yellow'   = "`e[33m"
        'Red'      = "`e[31m"
        'Magenta'  = "`e[35m"
        'DarkGray' = "`e[90m"
        'Reset'    = "`e[0m"
    }
} else {
    @{
        'White'    = [ConsoleColor]::White
        'Cyan'     = [ConsoleColor]::Cyan
        'Green'    = [ConsoleColor]::Green
        'Yellow'   = [ConsoleColor]::Yellow
        'Red'      = [ConsoleColor]::Red
        'Magenta'  = [ConsoleColor]::Magenta
        'DarkGray' = [ConsoleColor]::DarkGray
        'Reset'    = ''
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
'@

function Add-ColorSupportToScript {
    param(
        [string]$Content
    )

    # Check if Write-ColorOutput function already exists
    if ($Content -match 'function Write-ColorOutput') {
        # Function already exists, return content unchanged
        # Note: FilePath parameter kept for consistency with function signature
        return $Content
    }

    # Find the best location to insert the color support code
    # Look for the end of the param block more reliably
    $lines = $Content -split "`r?`n"
    $insertIndex = -1
    $inParamBlock = $false
    $braceCount = 0

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i].Trim()

        # Check if we're starting a param block
        if ($line -match '^param\s*\(') {
            $inParamBlock = $true
            $braceCount = ($line.ToCharArray() | Where-Object { $_ -eq '(' }).Count - ($line.ToCharArray() | Where-Object { $_ -eq ')' }).Count
        } elseif ($inParamBlock) {
            # Count braces to find end of param block
            $braceCount += ($line.ToCharArray() | Where-Object { $_ -eq '(' }).Count - ($line.ToCharArray() | Where-Object { $_ -eq ')' }).Count

            if ($braceCount -le 0) {
                # Found the end of param block
                $insertIndex = $i + 1
                break
            }
        }
    }

    # If we found a param block, insert after it
    if ($insertIndex -gt 0) {
        $beforeLines = $lines[0..($insertIndex - 1)]
        $afterLines = $lines[$insertIndex..($lines.Count - 1)]
        $newContent = ($beforeLines + "" + $Script:ColorSupportCode.Split("`n") + "" + $afterLines) -join "`n"
    } else {
        # Insert at the beginning after any initial comments
        $lines = $Content -split "`n"
        $insertIndex = 0

        # Skip initial comments and empty lines
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $line = $lines[$i].Trim()
            if (-not ($line -match '^#' -or $line -eq '')) {
                $insertIndex = $i
                break
            }
        }

        $beforeLines = $lines[0..($insertIndex - 1)]
        $afterLines = $lines[$insertIndex..($lines.Count - 1)]
        $newContent = ($beforeLines + '' + $Script:ColorSupportCode.Split("`n") + '' + $afterLines) -join "`n"
    }

    return $newContent
}

function Convert-WriteHostCall {
    param(
        [string]$Content
    )

    # Define color mapping from ForegroundColor to our Color parameter
    $colorMap = @{
        'Cyan'     = 'Cyan'
        'Green'    = 'Green'
        'Yellow'   = 'Yellow'
        'Red'      = 'Red'
        'DarkGray' = 'DarkGray'
        'Magenta'  = 'Magenta'
        'White'    = 'White'
        # Map Blue to Cyan as fallback
        'Blue'     = 'Cyan'
        # Map Gray to DarkGray as fallback
        'Gray'     = 'DarkGray'
    }

    $replacementCount = 0
    $modifiedContent = $Content

    # Replace Write-Host calls with Write-ColorOutput
    foreach ($color in $colorMap.Keys) {
        # Pattern 1: Write-Host "message" -ForegroundColor Color
        $pattern1 = "Write-Host\s+`"([^`"]+)`"\s+-ForegroundColor\s+$color"
        $replacement1 = "Write-ColorOutput -Message `"`$1`" -Color '$($colorMap[$color])'"
        $beforeCount = ($modifiedContent | Select-String -Pattern $pattern1 -AllMatches).Matches.Count
        $modifiedContent = $modifiedContent -replace $pattern1, $replacement1
        $afterCount = ($modifiedContent | Select-String -Pattern $pattern1 -AllMatches).Matches.Count
        $replacementCount += ($beforeCount - $afterCount)

        # Pattern 2: Write-Host 'message' -ForegroundColor Color
        $pattern2 = "Write-Host\s+'([^']+)'\s+-ForegroundColor\s+$color"
        $replacement2 = "Write-ColorOutput -Message '$1' -Color '$($colorMap[$color])'"
        $beforeCount = ($modifiedContent | Select-String -Pattern $pattern2 -AllMatches).Matches.Count
        $modifiedContent = $modifiedContent -replace $pattern2, $replacement2
        $afterCount = ($modifiedContent | Select-String -Pattern $pattern2 -AllMatches).Matches.Count
        $replacementCount += ($beforeCount - $afterCount)

        # Pattern 3: Write-Host $variable -ForegroundColor Color
        $pattern3 = "Write-Host\s+(\$[a-zA-Z_][a-zA-Z0-9_]*)\s+-ForegroundColor\s+$color"
        $replacement3 = "Write-ColorOutput -Message `$1 -Color '$($colorMap[$color])'"
        $beforeCount = ($modifiedContent | Select-String -Pattern $pattern3 -AllMatches).Matches.Count
        $modifiedContent = $modifiedContent -replace $pattern3, $replacement3
        $afterCount = ($modifiedContent | Select-String -Pattern $pattern3 -AllMatches).Matches.Count
        $replacementCount += ($beforeCount - $afterCount)

        # Pattern 4: Write-Host expression -ForegroundColor Color (for complex expressions)
        $pattern4 = "Write-Host\s+([^`"']+(?:`"[^`"]*`"[^`"']*)*)\s+-ForegroundColor\s+$color"
        $replacement4 = "Write-ColorOutput -Message $1 -Color '$($colorMap[$color])'"
        $beforeCount = ($modifiedContent | Select-String -Pattern $pattern4 -AllMatches).Matches.Count
        $modifiedContent = $modifiedContent -replace $pattern4, $replacement4
        $afterCount = ($modifiedContent | Select-String -Pattern $pattern4 -AllMatches).Matches.Count
        $replacementCount += ($beforeCount - $afterCount)
    }

    # Handle basic Write-Host calls without -ForegroundColor parameter
    # These patterns use negative lookahead to ensure we don't match calls that already have -ForegroundColor

    # Pattern 5: Write-Host "message" (no color parameter) - must be precise with quotes
    $pattern5 = 'Write-Host\s+"([^"]*)"(?!\s+-ForegroundColor)'
    $replacement5 = 'Write-ColorOutput -Message "$1" -Color "White"'
    $beforeCount = ($modifiedContent | Select-String -Pattern $pattern5 -AllMatches).Matches.Count
    $modifiedContent = $modifiedContent -replace $pattern5, $replacement5
    $afterCount = ($modifiedContent | Select-String -Pattern $pattern5 -AllMatches).Matches.Count
    $replacementCount += ($beforeCount - $afterCount)

    # Pattern 6: Write-Host 'message' (no color parameter) - must be precise with quotes
    $pattern6 = "Write-Host\s+'([^']*)'(?!\s+-ForegroundColor)"
    $replacement6 = "Write-ColorOutput -Message '$1' -Color 'White'"
    $beforeCount = ($modifiedContent | Select-String -Pattern $pattern6 -AllMatches).Matches.Count
    $modifiedContent = $modifiedContent -replace $pattern6, $replacement6
    $afterCount = ($modifiedContent | Select-String -Pattern $pattern6 -AllMatches).Matches.Count
    $replacementCount += ($beforeCount - $afterCount)

    # Pattern 7: Write-Host $variable (no color parameter) - end of line or whitespace
    $pattern7 = 'Write-Host\s+(\$[a-zA-Z_][a-zA-Z0-9_]*)\s*(?:\r?\n|$)'
    $replacement7 = 'Write-ColorOutput -Message $1 -Color "White"`r`n'
    $beforeCount = ($modifiedContent | Select-String -Pattern $pattern7 -AllMatches).Matches.Count
    $modifiedContent = $modifiedContent -replace $pattern7, $replacement7
    $afterCount = ($modifiedContent | Select-String -Pattern $pattern7 -AllMatches).Matches.Count
    $replacementCount += ($beforeCount - $afterCount)

    # Pattern 7b: Write-Host $variable (no color parameter) - followed by closing brace or other punctuation
    $pattern7b = 'Write-Host\s+(\$[a-zA-Z_][a-zA-Z0-9_]*)\s*(?=\s*[})])'
    $replacement7b = 'Write-ColorOutput -Message $1 -Color "White"'
    $beforeCount = ($modifiedContent | Select-String -Pattern $pattern7b -AllMatches).Matches.Count
    $modifiedContent = $modifiedContent -replace $pattern7b, $replacement7b
    $afterCount = ($modifiedContent | Select-String -Pattern $pattern7b -AllMatches).Matches.Count
    $replacementCount += ($beforeCount - $afterCount)

    # Pattern 8: Write-Host with parenthesized expressions (no color parameter)
    $pattern8 = 'Write-Host\s+(\([^)]+\))(?!\s+-ForegroundColor)'
    $replacement8 = 'Write-ColorOutput -Message $1 -Color "White"'
    $beforeCount = ($modifiedContent | Select-String -Pattern $pattern8 -AllMatches).Matches.Count
    $modifiedContent = $modifiedContent -replace $pattern8, $replacement8
    $afterCount = ($modifiedContent | Select-String -Pattern $pattern8 -AllMatches).Matches.Count
    $replacementCount += ($beforeCount - $afterCount)

    # Pattern 9: Write-Host with variable color parameter
    # Handles cases like: Write-Host $Message -ForegroundColor $Color
    $pattern9 = 'Write-Host\s+([^-]+?)\s+-ForegroundColor\s+(\$[a-zA-Z_][a-zA-Z0-9_]*)'
    $replacement9 = 'Write-ColorOutput -Message $1 -Color $2'
    $beforeCount = ($modifiedContent | Select-String -Pattern $pattern9 -AllMatches).Matches.Count
    $modifiedContent = $modifiedContent -replace $pattern9, $replacement9
    $afterCount = ($modifiedContent | Select-String -Pattern $pattern9 -AllMatches).Matches.Count
    $replacementCount += ($beforeCount - $afterCount)

    return @{
        Content = $modifiedContent
        ReplacementCount = $replacementCount
    }
}

function ConvertScript {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$FilePath,
        [bool]$BackupOriginal = $true
    )

    try {
        # Validate file exists and is a PowerShell script
        if (-not (Test-Path -Path ${FilePath})) {
            Write-Warning "File not found: ${FilePath}"
            return
        }

        if (-not (${FilePath} -match '\.ps1$')) {
            Write-Warning "File is not a PowerShell script: ${FilePath}"
            return
        }

        Write-Output "Processing: ${FilePath}"

        # Read the file content
        $originalContent = Get-Content -Path ${FilePath} -Raw -Encoding UTF8

        if ([string]::IsNullOrWhiteSpace($originalContent)) {
            Write-Warning "File is empty or could not be read: ${FilePath}"
            return
        }

        # Check if the file contains Write-Host calls
        $writeHostMatches = $originalContent | Select-String -Pattern "Write-Host" -AllMatches
        if (-not $writeHostMatches.Matches) {
            Write-Output "No Write-Host calls found in ${FilePath}"
            return
        }

        Write-Output "Found $($writeHostMatches.Matches.Count) Write-Host call(s) in ${FilePath}"

        # Create backup if requested
        if ($BackupOriginal) {
            $backupPath = "${FilePath}.backup-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
            if ($PSCmdlet.ShouldProcess(${FilePath}, "Create backup at $backupPath")) {
                Copy-Item -Path ${FilePath} -Destination $backupPath -Force
                Write-Output "Created backup: $backupPath"
            }
        }

        if ($PSCmdlet.ShouldProcess(${FilePath}, "Convert Write-Host calls to Write-ColorOutput")) {
            # Add color support code if needed
            $contentWithColorSupport = Add-ColorSupportToScript -Content $originalContent -FilePath ${FilePath}

            # Convert Write-Host calls
            $conversionResult = Convert-WriteHostCall -Content $contentWithColorSupport
            $modifiedContent = $conversionResult.Content
            $replacementCount = $conversionResult.ReplacementCount

            # Save the updated content
            $modifiedContent | Out-File -FilePath ${FilePath} -Encoding UTF8 -NoNewline

            Write-Output "Successfully converted $replacementCount Write-Host call(s) in ${FilePath}"
        }

    } catch {
        Write-Error "Error processing ${FilePath}`: $_"
    }
}

# Main execution
foreach ($path in ${FilePath}) {
    # Handle wildcards and multiple files
    $resolvedPaths = Get-ChildItem -Path $path -Filter "*.ps1" -ErrorAction SilentlyContinue

    if ($resolvedPaths) {
        foreach ($resolvedPath in $resolvedPaths) {
            ConvertScript -FilePath $resolvedPath.FullName -BackupOriginal $BackupOriginal
        }
    } else {
        # Try as direct file path
        ConvertScript -FilePath $path -BackupOriginal $BackupOriginal
    }
}

Write-Output "Write-Host to Write-ColorOutput conversion completed."
