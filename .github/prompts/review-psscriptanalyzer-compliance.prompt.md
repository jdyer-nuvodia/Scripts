---
description: "Review PowerShell scripts for PSScriptAnalyzer compliance and organizational standards"
mode: "agent"
---

# Review PowerShell Script for PSScriptAnalyzer Compliance

Perform a comprehensive review of the selected PowerShell script(s) to ensure full PSScriptAnalyzer compliance and adherence to organizational coding standards.

## Review Checklist:

### PSScriptAnalyzer Compliance:
1. **No Write-Host usage** - Must use Write-Output, Write-Verbose, Write-Warning, or Write-Error
2. **Named parameters** - All command invocations must use named parameters (no positional)
3. **No aliases** - Use full cmdlet names throughout
4. **Proper formatting** - Check for whitespace and newline issues
5. **Variable scoping** - Ensure proper variable declarations and scoping
6. **Security best practices** - No hardcoded credentials or sensitive data

### Organizational Standards:
1. **Header format** - Verify complete metadata header matching the required format
2. **Documentation** - Check .SYNOPSIS, .DESCRIPTION, .PARAMETER, and .EXAMPLE sections
3. **Error handling** - Confirm try/catch/finally blocks and proper exit codes
4. **Logging implementation** - Verify proper logging with correct file naming convention
5. **Parameterization** - Ensure no interactive prompts, proper defaults
6. **Automation readiness** - Check for -WhatIf support on state-changing operations

### Code Quality:
1. **Idempotency** - Can the script be run multiple times safely?
2. **Modularity** - Are functions properly structured and reusable?
3. **Performance** - Any obvious performance issues or inefficiencies?
4. **Security** - Any security vulnerabilities or concerns?

## Current File Context:
Selected file(s): ${selection}

## Action Items:
For each issue found:
1. **Identify the specific problem** with line numbers if applicable
2. **Explain why it violates standards** (PSScriptAnalyzer rule or org standard)
3. **Provide the exact fix** with corrected code
4. **Suggest improvements** for better practices

## Validation:
After suggesting fixes, run PSScriptAnalyzer validation to confirm all issues are resolved. The script must pass with zero warnings, errors, or information messages.

Reference the coding standards from [copilot-instructions.md](../copilot-instructions.md) for complete organizational requirements.
