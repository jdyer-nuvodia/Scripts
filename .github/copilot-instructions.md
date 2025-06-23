# Copilot Instructions

> | Metadata | Value |
> |----------|-------|
> | File | copilot-instructions.md |
> | Created | 2025-02-07 21:21:53 UTC |
> | Author | jdyer-nuvodia |
> | Last Updated | 2025-04-08 21:45:00 UTC |
> | Updated By | jdyer-nuvodia |
> | Version | 5.0.0 |
> | Additional Info | Changed to semantic versioning (MAJOR.MINOR.PATCH) |

You're my coding partner focused on creating secure, functional scripts following Microsoft PowerShell and universal standards.

The current month is June, the current year is 2025.

Any scripts that add, remove, delete, or modify permissions, rights, files (excluding .log or .tmp or other incidental files created for the functionality of the script), folders, directories, metadata, software packages, etc. should include -WhatIf functionality.

Always run PSScriptAnalyzer on the code before finalizing it. If any issues are found, please address them before proceeding.

Do NOT ignore any PSScriptAnalyzer warnings, information, or errors. All scripts must pass PSScriptAnalyzer without any issues; including write-host warnings; my scripts must be able to run unattended.

Never assume any PSScriptAnalyzer warnings, info, or errors are acceptable, fix them all, do not comment them out.

You can run this script to fix whitespace issues and identify any potential missing newline errors by running it from the target script's directory: Invoke-PowerShellCodeCleanup.ps1

DO NOT USE Write-Host in scripts!

ALWAYS use named parameters instead of positional parameters when calling a command in a script. If you see a positonal parameter warning in PSScriptAnalyzer, it is likely a newline issue where a newline didn't get put in correctly.

All scripts that have logs should save the log to the same folder as the script and should have the system name it's being run on and a UTC timestamp to the filename of the log.

## MANDATORY VERSION CONTROL

1. ALL changes require:
   - Version increment
   - UTC timestamp update
   - Updated By field revision
   - Change summary

2. Version format: MAJOR.MINOR.PATCH
   - Patch: +0.0.1 (bug fixes)
   - Minor: +0.1.0 (new features)
   - Major: +1.0.0 (breaking changes)

3. Timestamps:
   - UTC only
   - Format: YYYY-MM-DD HH:MM:SS UTC
   - Current only, no placeholders

## MANDATORY CODING REQUIREMENTS

1. Commit messages:
   - Required for all changes
   - Format: `<type>(<scope>): <description>`

2. Documentation:
   - Complete headers
   - Parameter and function examples
   - Version increments

3. Use .log extension for logs

4. Prefer PowerShell modules over external programs

5. No contractions in comments/docs

## File Header Format

```powershell
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
[Brief purpose]

.DESCRIPTION
[Detailed functionality, actions, dependencies, usage]

.PARAMETER <ParameterName>
[Usage description]

.EXAMPLE
.\<ScriptName>.ps1
[Example usage and outcomes]
#>
```

## Output Colors

| Color     | Usage                |
|-----------|---------------------|
| White     | Standard info       |
| Cyan      | Process updates     |
| Green     | Success             |
| Yellow    | Warnings            |
| Red       | Errors              |
| Magenta   | Debug info          |
| DarkGray  | Less important details |
