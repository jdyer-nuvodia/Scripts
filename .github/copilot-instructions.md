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

The current month is April, the current year is 2025.

Any scripts that add, remove, delete, or modify permissions, rights, files (excluding .log or .tmp or other incidental files created for the functionality of the script), folders, directories, metadata, software packages, etc. should include -WhatIf functionality.

Always run PSScriptAnalyzer on the code before finalizing it. If any issues are found, please address them before proceeding.

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
