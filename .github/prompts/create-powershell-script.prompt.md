---
description: "Create a new PowerShell script following organizational best practices and standards"
mode: "edit"
---

# Create New PowerShell Script

Create a new PowerShell script that follows all organizational coding standards and best practices. The script should include:

## Required Elements:
1. **Complete header** with proper metadata format including:
   - Script name, creation date (UTC), author, version (1.0.0 for new scripts)
   - Complete .SYNOPSIS and .DESCRIPTION sections
   - .PARAMETER documentation for all parameters
   - .EXAMPLE section with realistic usage

2. **Parameter block** with:
   - Named parameters with appropriate types and validation
   - Sensible default values where applicable
   - No prompting for user input (automation-ready)

3. **Error handling**:
   - Try/catch/finally blocks where appropriate
   - Proper exit codes for automation compatibility
   - -WhatIf support for state-changing operations

4. **Logging implementation**:
   - Log file naming: `ScriptName_${env:COMPUTERNAME}_$(Get-Date -Format 'yyyyMMdd_HHmmss').log`
   - Use Write-Output, Write-Verbose, Write-Warning, Write-Error (never Write-Host)
   - Save logs to script directory with .log extension

5. **Code quality**:
   - Full cmdlet names (no aliases)
   - Named parameters in all command invocations
   - PSScriptAnalyzer compliant code
   - Proper indentation and formatting

## Input Variables:
- **Script Purpose**: ${input:purpose:What is the main purpose of this script?}
- **Script Category**: ${input:category:Which category does this belong to? (e.g., CloudServices/365, SystemManagement, etc.)}
- **Parameters Needed**: ${input:parameters:What parameters does the script need? (describe each)}

## Requirements:
- Must be automation-friendly (no interactive prompts)
- Must include -WhatIf support for any state-changing operations
- Must pass PSScriptAnalyzer without any warnings or errors
- Must include comprehensive error handling
- Must follow the exact header format from [copilot-instructions.md](../copilot-instructions.md)

Create the complete script with all required elements, following the established patterns from existing scripts in the workspace.
