# =============================================================================
# File: copilot-instructions.md
# Created: 2025-02-07 21:21:53 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-06 22:45:00 UTC
# Updated By: jdyer-nuvodia
# Version: 3.2
# Additional Info: Expanded color schema to match organizational standards.
#
# IMPORTANT: When editing any script, update the header information including 
# Last Updated, Updated By, and Version fields. Use a current, reliable UTC timestamp.
# =============================================================================

The current year is 2025.
The month is 03.
NEVER USE RANDOM TIMESTAMPS! IF YOU NEED A TIMESTAMP GET ONE FROM A RELIABLE SOURCE!
When creating a script or file, make sure the header format below is used and the created date is set to the real life, current timestamp in UTC.
Always give me a commit message when writing, editing, or providing a script. Read this webpage for instructions on how to give me the best commit message you can: https://www.gitkraken.com/learn/git/best-practices/git-commit-message
Make sure all additional functionlity and examples, parameters, etc. are included in the header of any script when updating and the versioning is increased accordingly each time it is updated. Also update the last updated field in the header each time.
All log files should be .log and not .txt files.
Use .NET methods over other methods when available in PowerShell scripts.

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