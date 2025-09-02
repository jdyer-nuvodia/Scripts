# =============================================================================
# Script: Invoke-PSScriptAnalyzerCheck.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-15 23:40:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2.3
# Additional Info: Reverted ANSI escape sequences back to backtick-e format (false positive)
# =============================================================================

<#
.SYNOPSIS
    Runs comprehensive PSScriptAnalyzer checks on PowerShell scripts for code quality compliance.

.DESCRIPTION
    This script performs comprehensive static code analysis using PSScriptAnalyzer with:
    - All default rules (Error, Warning, Information)
    - Code formatting rules (PSUseConsistentWhitespace, PSPlaceCloseBrace, etc.)
    - Best practice rules
    - Security rules
    - Performance rules

    EXCLUDED RULES:
    - PSUseConsistentWhitespace: Excluded to allow vertical alignment of operators
      per organizational coding standards in copilot-instructions.md
    - PSPlaceCloseBrace: Excluded to allow organizational brace placement style

    The script provides detailed output showing all issues found and categorizes them by severity.
    It also provides a summary count of issues by type and severity level.

.PARAMETER FilePath
    The absolute path to the PowerShell script file to analyze.

.PARAMETER IncludeDefaultRules
    Include all default PSScriptAnalyzer rules. Default is $true.

.PARAMETER IncludeCodeFormatting
    Include code formatting rules. Default is $true.

.PARAMETER ShowSummaryOnly
    Show only the summary of issues without detailed output. Default is $false.

.EXAMPLE
    .\Invoke-PSScriptAnalyzerCheck.ps1 -FilePath "C:\Scripts\MyScript.ps1"
    Runs comprehensive analysis on MyScript.ps1 and displays all issues found.

.EXAMPLE
    .\Invoke-PSScriptAnalyzerCheck.ps1 -FilePath "C:\Scripts\MyScript.ps1" -ShowSummaryOnly
    Runs analysis and shows only a summary of issues found.

.NOTES
    Security Level: Low
    Required Permissions: Read access to target script file
    Dependencies: PSScriptAnalyzer module must be installed

    ORGANIZATIONAL STANDARDS:
    This script excludes PSUseConsistentWhitespace and PSPlaceCloseBrace rules to
    allow vertical alignment of operators and organizational brace placement style
    as specified in copilot-instructions.md. These decisions prioritize code
    readability and organizational standards over strict PSScriptAnalyzer formatting rules.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({
            if (-not (Test-Path -Path $_ -PathType Leaf)) {
                throw "File path '$_' does not exist or is not a file."
            }
            if (-not ($_ -match '\.ps1$|\.psm1$|\.psd1$')) {
                throw "File '$_' is not a PowerShell script file (.ps1, .psm1, .psd1)."
            }
            return $true
        })]
    [string]$FilePath,

    [Parameter(Mandatory = $false)]
    [bool]$IncludeDefaultRules = $true,

    [Parameter(Mandatory = $false)]
    [bool]$IncludeCodeFormatting = $true,

    [Parameter(Mandatory = $false)]
    [bool]$ShowSummaryOnly = $false
)

# Color support variables and Write-ColorOutput function
# Note: This script intentionally uses vertical alignment which may trigger
# PSUseConsistentWhitespace warnings. This is per organizational standards.
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
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$Message = "",
        [Parameter(Mandatory = $false)]
        [string]$Color = "White"
    )

    if ($Script:UseAnsiColors) {
        $colorCode = $Script:Colors[$Color]
        $resetCode = $Script:Colors.Reset
        Write-Output "${colorCode}${Message}${resetCode}"
    } else {
        Write-Output $Message
    }
}

# Check if PSScriptAnalyzer module is available
try {
    Import-Module PSScriptAnalyzer -ErrorAction Stop
} catch {
    Write-ColorOutput -Message "ERROR: PSScriptAnalyzer module is not installed or available." -Color 'Red'
    Write-ColorOutput -Message "Please install it using: Install-Module -Name PSScriptAnalyzer -Force" -Color 'Yellow'
    exit 1
}

# Resolve the full path
$ScriptPath = Resolve-Path -Path $FilePath
$ScriptName = Split-Path -Path $ScriptPath -Leaf

Write-ColorOutput -Message "=======================================================" -Color 'White'
Write-ColorOutput -Message "PSScriptAnalyzer Comprehensive Analysis" -Color 'White'
Write-ColorOutput -Message "=======================================================" -Color 'White'
Write-ColorOutput -Message "Script: $ScriptName" -Color 'Cyan'
Write-ColorOutput -Message "Path: $ScriptPath" -Color 'DarkGray'
Write-ColorOutput -Message "Analysis started: $(Get-Date)" -Color 'DarkGray'
Write-ColorOutput -Message "-------------------------------------------------------" -Color 'DarkGray'

# Initialize results collections
$AllIssues = @()
$ErrorCount = 0
$WarningCount = 0
$InformationCount = 0

try {
    # Define rules to exclude (operator spacing conflicts with vertical alignment)
    $ExcludeRules = @('PSUseConsistentWhitespace', 'PSPlaceCloseBrace', 'PSUseBOMForUnicodeEncodedFile')

    # Run default rules analysis
    if ($IncludeDefaultRules) {
        Write-ColorOutput -Message "Running default PSScriptAnalyzer rules..." -Color 'Cyan'
        $DefaultResults = Invoke-ScriptAnalyzer -Path $ScriptPath -Severity @('Error', 'Warning', 'Information') -ExcludeRule $ExcludeRules
        $AllIssues += $DefaultResults
    }

    # Run code formatting analysis
    if ($IncludeCodeFormatting) {
        Write-ColorOutput -Message "Running code formatting rules..." -Color 'Cyan'
        $FormattingResults = Invoke-ScriptAnalyzer -Path $ScriptPath -Settings CodeFormatting -ExcludeRule $ExcludeRules
        $AllIssues += $FormattingResults
    }

    # Remove duplicates (in case same issue is found by multiple rule sets)
    $UniqueIssues = $AllIssues | Sort-Object RuleName, Line, Column | Get-Unique -AsString

    # Count issues by severity
    $ErrorCount = ($UniqueIssues | Where-Object { $_.Severity -eq 'Error' }).Count
    $WarningCount = ($UniqueIssues | Where-Object { $_.Severity -eq 'Warning' }).Count
    $InformationCount = ($UniqueIssues | Where-Object { $_.Severity -eq 'Information' }).Count
    $TotalIssues = $UniqueIssues.Count

    Write-ColorOutput -Message "-------------------------------------------------------" -Color 'DarkGray'
    Write-ColorOutput -Message "ANALYSIS RESULTS:" -Color 'White'

    if ($TotalIssues -eq 0) {
        Write-ColorOutput -Message "No issues found! Script passes all PSScriptAnalyzer checks." -Color 'Green'
    } else {
        Write-ColorOutput -Message "Total Issues Found: $TotalIssues" -Color 'Yellow'
        Write-ColorOutput -Message "  Errors: $ErrorCount" -Color 'Red'
        Write-ColorOutput -Message "  Warnings: $WarningCount" -Color 'Yellow'
        Write-ColorOutput -Message "  Information: $InformationCount" -Color 'Cyan'

        if (-not $ShowSummaryOnly) {
            Write-ColorOutput -Message "" -Color 'White'
            Write-ColorOutput -Message "DETAILED ISSUES:" -Color 'White'
            Write-ColorOutput -Message "-------------------------------------------------------" -Color 'DarkGray'

            # Group and display issues by severity
            $SeverityOrder = @('Error', 'Warning', 'Information')
            foreach ($Severity in $SeverityOrder) {
                $IssuesOfSeverity = $UniqueIssues | Where-Object { $_.Severity -eq $Severity }
                if ($IssuesOfSeverity) {
                    $SeverityColor = switch ($Severity) {
                        'Error' { 'Red' }
                        'Warning' { 'Yellow' }
                        'Information' { 'Cyan' }
                        default { 'White' }
                    }

                    Write-ColorOutput -Message "" -Color 'White'
                    Write-ColorOutput -Message "${Severity}s:" -Color $SeverityColor

                    foreach ($Issue in $IssuesOfSeverity) {
                        $LineInfo = if ($Issue.Line) { "Line $($Issue.Line)" } else { "N/A" }
                        $ColumnInfo = if ($Issue.Column) { ", Column $($Issue.Column)" } else { "" }
                        $LocationInfo = "$LineInfo$ColumnInfo"

                        Write-ColorOutput -Message "  [$LocationInfo] $($Issue.RuleName): $($Issue.Message)" -Color $SeverityColor
                    }
                }
            }
        }

        Write-ColorOutput -Message "" -Color 'White'
        Write-ColorOutput -Message "RECOMMENDATIONS:" -Color 'White'
        Write-ColorOutput -Message "-------------------------------------------------------" -Color 'DarkGray'

        if ($ErrorCount -gt 0) {
            Write-ColorOutput -Message "* Fix all ERRORS immediately - these can cause script failures" -Color 'Red'
        }
        if ($WarningCount -gt 0) {
            Write-ColorOutput -Message "* Address WARNINGS for best practices and maintainability" -Color 'Yellow'
        }
        if ($InformationCount -gt 0) {
            Write-ColorOutput -Message "* Review INFORMATION items for code style improvements" -Color 'Cyan'
        }

        # Provide exit code based on severity
        if ($ErrorCount -gt 0) {
            Write-ColorOutput -Message "" -Color 'White'
            Write-ColorOutput -Message "RESULT: FAILED - Script has PSScriptAnalyzer errors that must be fixed." -Color 'Red'
            exit 1
        } elseif ($WarningCount -gt 0) {
            Write-ColorOutput -Message "" -Color 'White'
            Write-ColorOutput -Message "RESULT: WARNINGS - Script has warnings that should be addressed." -Color 'Yellow'
            exit 2
        } else {
            Write-ColorOutput -Message "" -Color 'White'
            Write-ColorOutput -Message "RESULT: PASSED - Only informational items found." -Color 'Green'
            exit 0
        }
    }

} catch {
    Write-ColorOutput -Message "ERROR: Failed to run PSScriptAnalyzer analysis." -Color 'Red'
    Write-ColorOutput -Message "Error details: $($_.Exception.Message)" -Color 'Red'
    exit 1
} finally {
    Write-ColorOutput -Message "" -Color 'White'
    Write-ColorOutput -Message "Analysis completed: $(Get-Date)" -Color 'DarkGray'
    Write-ColorOutput -Message "=======================================================" -Color 'White'
}
