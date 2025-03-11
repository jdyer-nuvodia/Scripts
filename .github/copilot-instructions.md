# =============================================================================
# File: copilot-instructions.md
# Created: 2025-02-07 21:21:53 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-06 17:30:00 UTC
# Updated By: jdyer-nuvodia
# Version: 4.0.0
# Additional Info: Changed to semantic versioning (MAJOR.MINOR.PATCH)
# =============================================================================

## ⚠️ MANDATORY VERSION CONTROL REQUIREMENTS - STRICT ENFORCEMENT

1. VERSION CONTROL IS NOT OPTIONAL. ALL script modifications MUST include:
   - Version number increment (Required)
   - UTC timestamp update (Required)
   - Updated By field revision (Required)
   - Change summary in Additional Info (Required)
   - NO EXCEPTIONS OR OMISSIONS PERMITTED

2. STRICT Version numbering policy:
   - Format: MAJOR.MINOR.PATCH (Example: 1.0.0, 1.1.0, 2.0.0)
   - Patch changes: +0.0.1 (Example: 1.0.0 -> 1.0.1) for backwards compatible bug fixes
   - Minor changes: +0.1.0 (Example: 1.0.0 -> 1.1.0) for backwards compatible features
   - Major changes: +1.0.0 (Example: 1.9.9 -> 2.0.0) for breaking changes
   - ALL changes require version increment

3. TIMESTAMP REQUIREMENTS (CRITICAL):
   - UTC timezone ONLY
   - Format MUST be: YYYY-MM-DD HH:MM:SS UTC
   - CURRENT timestamps only - NO placeholders
   - NO historical or future dates

Example header update:
```
# Last Updated: 2025-03-07 15:30:00 UTC  # Always current UTC
# Updated By: editor-name                 # Person making changes
# Version: 1.0.1                         # Incremented from 1.0.0
# Additional Info: Fixed logging bug      # What changed
```

## ⚠️ MANDATORY CODING REQUIREMENTS - STRICT ENFORCEMENT

1. COMMIT MESSAGE REQUIREMENTS (MANDATORY):
   - Every script creation or modification MUST include a commit message
   - Follow the guidelines at: https://www.gitkraken.com/learn/git/best-practices/git-commit-message
   - Format: <type>(<scope>): <description>
   - Example: "feat(logging): implement new error handling system"
   - NO EXCEPTIONS OR OMISSIONS PERMITTED

2. SCRIPT DOCUMENTATION POLICY (REQUIRED):
   - ALL scripts MUST include complete header documentation
   - ALL parameters MUST be documented with examples
   - ALL functions MUST include usage examples
   - Additional functionality MUST be reflected in header updates
   - Version number MUST be incremented (see Version Control section)

3. FILE TYPE REQUIREMENTS:
   - ALL log files MUST use .log extension
   - .txt extension for log files is STRICTLY PROHIBITED

4. POWERSHELL CODING STANDARDS:
   - .NET methods MUST be used when available.
   - If functionality can be done in Powershell with modules, use PowerShell instead of calling other programs.
   - Native PowerShell alternatives are ONLY acceptable when .NET methods are unavailable
   - Example: Use [System.IO.File]::ReadAllText() instead of Get-Content
   - Example: Use [System.IO.Directory]::GetFiles() instead of Get-ChildItem

5. SCRIPT GRAMMAR REQUIREMENTS
   - Do NOT use contractions in comments or documentation. 
      - Contractions are words that end in 're, 's, 'nt, etc.

These requirements are NON-NEGOTIABLE and MUST be followed without exception.
Failure to comply will result in automatic rejection of contributions.

# Current date info
The current year is 2025.
The month is 03.

## Mandatory File & Script Header Format

# =============================================================================
# Script: <ScriptName>.ps1
# Created: <YYYY-MM-DD HH:MM:SS UTC>
# Author: <AuthorName>
# Last Updated: <YYYY-MM-DD HH:MM:SS UTC>
# Updated By: <AuthorName or Collaborator>
# Version: <VersionNumber>
# Additional Info: <Additional contextual data>
# =============================================================================

<#
.SYNOPSIS
    [Brief statement of the script's purpose]
.DESCRIPTION
    [Detailed explanation of the script's functionality, including:
     - Key actions
     - Dependencies or prerequisites
     - Usage or examples (if needed)]
.PARAMETER <ParameterName>
    [Description of parameter usage if applicable]
.EXAMPLE
    .\<ScriptName>.ps1
    [Example usage, describing outcomes or key steps]
#>

## Standard Output Color Scheme

Use these standard colors for script output to maintain consistency:

| Color    | Usage                                        | Example Write-Host Usage                               |
|----------|----------------------------------------------|------------------------------------------------------|
| White    | Standard information, default messages       | `Write-Host "Processing file..." -ForegroundColor White` |
| Cyan     | Process starting, status updates             | `Write-Host "Starting process..." -ForegroundColor Cyan` |
| Green    | Success messages, completion confirmations    | `Write-Host "Operation completed" -ForegroundColor Green` |
| Yellow   | Warnings, retries, fallback actions          | `Write-Host "Retrying..." -ForegroundColor Yellow` |
| Red      | Errors, critical issues                      | `Write-Host "Critical failure!" -ForegroundColor Red` |
| Magenta  | Debug information                            | `Write-Host "Debug: $variable = $value" -ForegroundColor Magenta` |
| DarkGray | Technical details, less important info       | `Write-Host "Registry key: $key" -ForegroundColor DarkGray` |

### Color Usage Guidelines

1. **White** - Use for:
   - Standard information messages
   - Default text output
   - Section headers or dividers

2. **Cyan** - Use for:
   - Process initiation messages
   - Status updates
   - Configuration steps
   - Progress indicators

3. **Green** - Use for:
   - Successful completion messages
   - Resource creation confirmations
   - Validation successes
   - Positive metrics or results

4. **Yellow** - Use for:
   - Warning messages
   - Retry attempts
   - Using fallback options
   - Non-critical issues
   - Attention-requiring information

5. **Red** - Use for:
   - Error messages
   - Critical failures
   - Exception reporting
   - Security issues
   - Required actions

6. **Magenta** - Use for:
   - Debug information
   - Verbose technical details
   - Developer-oriented messages
   - Trace information

7. **DarkGray** - Use for:
   - Technical details
   - Background information
   - Less important messages
   - File paths, commands, or code examples