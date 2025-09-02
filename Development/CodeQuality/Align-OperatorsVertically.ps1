# =============================================================================
# Script: Align-OperatorsVertically.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-15 23:35:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.0
# Additional Info: Identifies operator alignment opportunities for improved code readability
# =============================================================================

<#
.SYNOPSIS
Identifies operator alignment opportunities in a specific PowerShell script file.

.DESCRIPTION
This script analyzes a single PowerShell file to identify potential operator
alignment issues that could improve code readability. It detects:
1. Consecutive variable assignments that could benefit from vertical alignment
2. Hashtable assignments with unaligned operators
3. Other operator alignment opportunities

The script does not modify files, only reports alignment opportunities for manual review.

.PARAMETER FilePath
The path to the specific PowerShell script file to analyze.

.PARAMETER WhatIf
Shows what would be analyzed without making actual changes (informational only).

.EXAMPLE
.\Align-OperatorsVertically.ps1 -FilePath "C:\Scripts\MyScript.ps1"
Analyzes the specified PowerShell file for alignment opportunities.

.EXAMPLE
.\Align-OperatorsVertically.ps1 -FilePath ".\Test-Script.ps1"
Analyzes a PowerShell file in the current directory for alignment opportunities.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({
            if (-not (Test-Path -Path $_ -PathType Leaf)) {
                throw "The file '$_' does not exist."
            }
            if (-not ($_ -match '\.ps1$')) {
                throw "The file '$_' must be a PowerShell script file (.ps1)."
            }
            return $true
        })]
    [string]$FilePath
)

# Analyze the specified PowerShell file
$file = Get-Item -Path $FilePath
Write-Output "Analyzing PowerShell file: $($file.Name)"

$alignmentIssues = @()

try {
    $lines = Get-Content -Path $file.FullName

    # Check for potential alignment issues

    # Check for consecutive variable assignments that could benefit from alignment
    for ($i = 0; $i -lt ($lines.Count - 1); $i++) {
        $currentLine = $lines[$i]
        $nextLine = $lines[$i + 1]

        if ($currentLine -match '^\s*\$\w+\s*=\s*' -and $nextLine -match '^\s*\$\w+\s*=\s*') {
            $alignmentIssues += "Lines $($i + 1)-$($i + 2): Consecutive variable assignments"
        }
    }

    # Check for hashtable assignments with unaligned operators
    $hashtableLines = $lines | Select-String -Pattern "^\s*'[^']*'\s*=\s*" | Where-Object { $_.Line -notmatch "^\s*'[^']*'\s{4,}=\s*" }
    if ($hashtableLines) {
        $alignmentIssues += "Hashtable with unaligned operators at lines: $($hashtableLines.LineNumber -join ', ')"
    }

} catch {
    Write-Error "Error analyzing $($file.FullName): $($_.Exception.Message)"
}

Write-Output "`nAnalysis complete!"
Write-Output "File analyzed: $($file.Name)"

if ($alignmentIssues.Count -gt 0) {
    Write-Output "`nPotential alignment opportunities found:"
    Write-Output "======================================"
    foreach ($issue in $alignmentIssues) {
        Write-Output "  - $issue"
    }

    Write-Output "`nNote: Manual review recommended to ensure alignment improves readability"
    Write-Output "while maintaining PSScriptAnalyzer compliance per copilot-instructions.md"
} else {
    Write-Output "`nNo obvious alignment opportunities detected in the analyzed file."
}
