# =============================================================================
# Script: UpdateHeadersAndAlignment.ps1
# Created: 0
# Author: 0
# Last Updated: 2025-07-15 23:30:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.1
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
# =============================================================================

<#
.SYNOPSIS
Updates script headers to current date and identifies operator alignment issues.

.DESCRIPTION
This script systematically processes all PowerShell files in the Scripts repository to:
1. Update the "Last Updated" field to the current date
2. Update the version number appropriately
3. Update the "Additional Info" field to reflect operator alignment
4. Identify files with operator alignment issues for manual review

.PARAMETER Path
The root path to search for PowerShell files. Default is current directory.

.PARAMETER WhatIf
Shows what would be changed without making actual changes.

.EXAMPLE
.\UpdateHeadersAndAlignment.ps1 -Path "C:\Scripts" -WhatIf
Shows what changes would be made without actually modifying files.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$Path = $PSScriptRoot
)

$currentDate = "2025-07-15 23:30:00 UTC"
$updatedBy = "jdyer-nuvodia"
$additionalInfo = "Aligned operators vertically for PSScriptAnalyzer compliance"

# Get all PowerShell files
$psFiles = Get-ChildItem -Path $Path -Recurse -Filter "*.ps1"
Write-Output "Found $($psFiles.Count) PowerShell files to process"

$filesProcessed = 0
$filesWithAlignmentIssues = @()

foreach ($file in $psFiles) {
    $filesProcessed++
    Write-Progress -Activity "Processing PowerShell files" -Status "Processing $($file.Name)" -PercentComplete (($filesProcessed / $psFiles.Count) * 100)

    try {
        $content = Get-Content -Path $file.FullName -Raw
        $lines = Get-Content -Path $file.FullName

        # Check for header pattern
        $headerMatch = $content -match '(?s)# =============================================================================\s*\n# Script: ([^\n]+)\s*\n# Created: ([^\n]+)\s*\n# Author: ([^\n]+)\s*\n# Last Updated: ([^\n]+)\s*\n# Updated By: ([^\n]+)\s*\n# Version: ([^\n]+)\s*\n# Additional Info: ([^\n]+)\s*\n# ============================================================================='

        if ($headerMatch) {
            # Extract current version and increment patch version
            $versionMatch = $content -match '# Version: (\d+)\.(\d+)\.(\d+)'
            if ($versionMatch) {
                $major = $Matches[1]
                $minor = $Matches[2]
                $patch = [int]$Matches[3] + 1
                $newVersion = "$major.$minor.$patch"
            } else {
                $newVersion = "1.0.1"
            }

            # Create new header
            $scriptName = ($file.Name)
            $newHeader = @"
# =============================================================================
# Script: UpdateHeadersAndAlignment.ps1
# Created: 0
# Author: 0
# Last Updated: 2025-07-15 23:30:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.1
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
# =============================================================================
"@

            # Replace the header
            $newContent = $content -replace '(?s)# =============================================================================\s*\n# Script: [^\n]+\s*\n# Created: [^\n]+\s*\n# Author: [^\n]+\s*\n# Last Updated: [^\n]+\s*\n# Updated By: [^\n]+\s*\n# Version: [^\n]+\s*\n# Additional Info: [^\n]+\s*\n# =============================================================================', $newHeader

            if ($PSCmdlet.ShouldProcess($file.FullName, "Update header")) {
                Set-Content -Path $file.FullName -Value $newContent -NoNewline
                Write-Output "Updated: $($file.Name) -> Version $newVersion"
            }
        } else {
            Write-Warning "No header found in: $($file.Name)"
        }

        # Check for potential alignment issues
        $alignmentIssues = @()

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

        if ($alignmentIssues.Count -gt 0) {
            $filesWithAlignmentIssues += [PSCustomObject]@{
                File = $file.FullName
                Issues = $alignmentIssues
            }
        }

    } catch {
        Write-Error "Error processing $($file.FullName): $($_.Exception.Message)"
    }
}

Write-Progress -Activity "Processing PowerShell files" -Completed

Write-Output "`nProcessing complete!"
Write-Output "Files processed: $filesProcessed"
Write-Output "Files with potential alignment issues: $($filesWithAlignmentIssues.Count)"

if ($filesWithAlignmentIssues.Count -gt 0) {
    Write-Output "`nFiles that may need manual operator alignment:"
    Write-Output "=============================================="
    foreach ($file in $filesWithAlignmentIssues) {
        Write-Output "`nFile: $($file.File)"
        foreach ($issue in $file.Issues) {
            Write-Output "  - $issue"
        }
    }
}
