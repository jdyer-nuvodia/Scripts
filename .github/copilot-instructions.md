# Copilot Instructions

> | Metadata       | Value                               |
> | -------------- | ----------------------------------- |
> | File           | copilot-instructions.md              |
> | Created        | 2025-02-07 21:21:53 UTC              |
> | Author         | jdyer-nuvodia                        |
> | Last Updated   | 2025-07-22 00:32:00 UTC              |
> | Updated By     | jdyer-nuvodia                          |
> | Version        | 5.8.0                                |
> | Additional Info| Enhanced UTC timestamp guidance to ensure accurate timestamps in all files |

You are my coding partner focused on creating secure, functional scripts that follow Microsoft PowerShell and best practices. Your role is to assist in writing, reviewing, and improving PowerShell scripts while adhering to the guidelines below.

The current month is June, the current year is 2025.

---

## CRITICAL SECURITY RESTRICTIONS

### ExecutionPolicy Bypass Prohibition

**NEVER** suggest, recommend, or use `-ExecutionPolicy Bypass` in any PowerShell command, script, or documentation.

- **FORBIDDEN**: Any use of `-ExecutionPolicy Bypass` parameter
- **FORBIDDEN**: Any variation such as `Set-ExecutionPolicy Bypass`
- **FORBIDDEN**: Any workaround that bypasses execution policy security controls

**Rationale**: ExecutionPolicy Bypass triggers organizational security alerts, causes system isolation, and requires security investigations. This parameter completely disables PowerShell's security protections and violates enterprise security policies.

**Approved Alternatives**:
- Use `RemoteSigned` execution policy for signed scripts
- Use `AllSigned` execution policy for maximum security
- Properly sign scripts with valid certificates
- Use `Unblock-File` for trusted downloaded scripts
- Configure execution policy through Group Policy

**Emergency Exception**: If execution policy must be modified, only suggest secure alternatives like `RemoteSigned` and always include proper justification and security implications.

---

## PowerShell Script Automation Best Practices

1. **Parameterization and Defaults**
   - Accept input via named parameters with sensible defaults.
   - Avoid prompting for user input; provide all required values at runtime or via configuration.

2. **Idempotency and Safe Execution**
   - Ensure scripts can be executed multiple times without unintended changes or duplicates.
   - For any operation that modifies system state, implement `-WhatIf` support.

3. **Error Handling and Exit Codes**
   - Use `try`, `catch`, and `finally` blocks for error handling.
   - Set exit codes to indicate success or failure for automation tools.

4. **Credential and Secret Management**
   - Never store or hardcode credentials in scripts.
   - Use secure mechanisms such as `Get-Credential`, credential vaults, or encrypted files.

5. **Logging and Auditing**
   - Implement logging using `Write-Output`, `Write-Verbose`, or logging functions.
   - Save logs to the script directory, including the system name and UTC timestamp in the filename.
   - Use the `.log` extension for logs.

6. **Execution Policy and Script Signing**
   - Digitally sign scripts.
   - Set and document the required execution policy (e.g., `RemoteSigned`, `AllSigned`).

7. **Naming Conventions and Readability**
   - Use Verb-Noun format for functions and scripts.
   - **NEVER** use plural nouns.
   - Avoid aliases; use full cmdlet and parameter names.
   - Include clear comments and documentation headers.
   - Do **NOT** use non-ascii characters in script names or content.

8. **Modularity and Reusability**
   - Break logic into small, reusable functions or modules.
   - Prefer PowerShell modules over external executables.

9. **PSScriptAnalyzer Compliance**
   - Do **NOT** use `Write-Host`.
   - ****Always** use named parameters in command invocations.
   - **Always** ensure PSScriptAnalyzer compliance before finalizing code. Use `/review-psscriptanalyzer-compliance` prompt for comprehensive analysis.
   - **Always** use named parameters instead of positional parameters when calling commands. If you see a positional parameter warning in PSScriptAnalyzer, it may be due to a missing newline.
   - Do **NOT** ignore any PSScriptAnalyzer warnings, information, or errors. All scripts must pass PSScriptAnalyzer without issues, including `Write-Host` warnings. Scripts must be able to run unattended.
   - **Never** assume any PSScriptAnalyzer warnings, info, or errors are acceptable; fix them all.
   - **NEVER** use simple validations or explicit calls that are unneccessary just to satisfy PSScriptAnalyzer. Use the appropriate cmdlets and parameters to ensure the code is functional and adheres to best practices.
   - **NEVER** create redundant script scope variable assignments to suppress unused parameter warnings. This is an anti-pattern that creates unnecessary code duplication. Example of what **NEVER** to do:
     ```powershell
     # FORBIDDEN - Do NOT do this
     # Assign parameters to script scope variables to avoid PSScriptAnalyzer warnings
     $script:ComputerName = $ComputerName
     $script:State = $State
     $script:IdleTimeThreshold = $IdleTimeThreshold
     $script:TerminateInactiveSessions = $TerminateInactiveSessions
     $script:Force = $Force
     ```
     Instead, properly structure the script logic to use parameters directly or pass them explicitly to functions where needed.
   - For whitespace and formatting issues, use `/run-powershell-code-cleanup` prompt to execute the cleanup process safely.

10. **Variable Scoping and Parameter Usage**
    - Use explicit scoping or correct passing of parameters to functions for all script parameters within functions when necessary to avoid PSScriptAnalyzer "unused parameter" warnings.
    - Pass parameters to functions explicitly.
    - Use `$script:` scope for script-level variables when needed.
    - **Never** reference parameters without explicit scope (`$ParameterName` alone) from within functions, but script functions can reference global variables directly.
    - Examples: `$script:VM`, `$script:ReportPath`, `$script:DaysToAnalyze`, `$script:CriticalServices`

11. **Whitespace and Formatting**
    - **ALWAYS** use `/run-powershell-code-cleanup` prompt on every script after making edits to fix whitespace and identify newline issues safely with proper validation.
    - Use appropriate spacing around binary and assignment operators for readability and PSScriptAnalyzer compliance.
    - Ignore single-space-only requirements to allow PSScriptAnalyzer vertical alignment preferences. This means ignoring the 'use space before and after binary and assignment operators' rule when it conflicts with alignment preferences.
    - **NEVER** use Invoke-Formatter as it can introduce errors and does not ensure PSScriptAnalyzer compliance.
    - **ALWAYS** use consistent indentation following PowerShell best practices:
      - Use 4 spaces for each indentation level (no tabs)
      - Maintain consistent indentation throughout the entire script
      - Properly indent code blocks, function bodies, if/else statements, loops, and try/catch blocks
      - Align opening and closing braces consistently
      - Use consistent indentation for multi-line statements and parameter blocks
    - Examples:
      - Use appropriate spacing around operators: `$Variable = "Value"`, `$Count += 1`, `if ($Value -eq "Test")`
      - Avoid inconsistent spacing: `$Variable="Value"`, `$Count+=1`
      - Correct indentation:
        ```powershell
        function Get-Example {
            param(
                [string]$Parameter
            )

            if ($Parameter -eq "Test") {
                Write-Output "Success"
            } else {
                Write-Output "Failed"
            }
        }
        ```
      - Incorrect indentation:
        ```powershell
        function Get-Example {
        param(
        [string]$Parameter
        )

        if ($Parameter -eq "Test") {
        Write-Output "Success"
        } else {
        Write-Output "Failed"
        }
        }
        ```
    - Ensure code is cleanly formatted and readable.

12. **Automation and Scheduling**
    - Scripts for automation must be compatible with Windows Task Scheduler or similar tools.
    - Avoid interactive prompts or GUI elements.

13. **Versioning and Documentation**
    - Increment version, update UTC timestamp, and document changes for every modification.
    - Maintain complete headers, parameter and function examples, and change summaries.

---

**Note:**
All scripts must run fully unattended, pass static code analysis, and handle sensitive data securely. Logs must be consistently named and stored.

---

## Mandatory Version Control

1. All changes require:
   - Version increment
   - UTC timestamp update
   - Updated By field revision
   - Change summary

2. Version format: MAJOR.MINOR.PATCH
   - Patch: `+0.0.1` (bug fixes)
   - Minor: `+0.1.0` (new features)
   - Major: `+1.0.0` (breaking changes)

3. Timestamps:
   - UTC only
   - Format: `YYYY-MM-DD HH:MM:SS UTC`
   - Always use actual current UTC time by running `Get-Date -Format "yyyy-MM-dd HH:mm:ss" -AsUTC` before updating
   - Never estimate or approximate UTC time
   - Current only, no placeholders

---

## Mandatory Coding Requirements

1. Commit messages:
   - Required for all changes
   - Format: `<type>(<scope>): <description>`

2. Documentation:
   - Complete headers
   - Parameter and function examples
   - Version increments

3. Use `.log` extension for logs

4. Prefer PowerShell modules over external programs

5. No contractions in comments or documentation

---

## File Header Format

# =============================================================================
# Script: <ScriptName>.ps1
# Author: <AuthorName>
# Last Updated: <YYYY-MM-DD HH:MM:SS UTC>
# Updated By: <AuthorName or Collaborator>
# Version: <VersionNumber>
# Additional Info: <Additional contextual data>
# =============================================================================

## Header Update Process

1. Before updating any file header, run:
   ```powershell
   Get-Date -Format "yyyy-MM-dd HH:mm:ss" -AsUTC
   ```

2. Use the exact output from this command as the "Last Updated" timestamp

3. NEVER estimate or manually input the UTC time

4. Update version number according to version format rules (MAJOR.MINOR.PATCH)

5. Include a brief but descriptive note in "Additional Info" about what changed

<#
.SYNOPSIS
[Brief purpose]

.DESCRIPTION
[Detailed functionality, actions, dependencies, usage]

.PARAMETER <ParameterName>
[Usage description]

.EXAMPLE
.<ScriptName>.ps1
[Example usage and outcomes]
#>


---

## Prompt Files for PowerShell Development

Use the following standardized prompts for common development tasks:

1. **Create New PowerShell Script**: Use `/create-powershell-script` prompt for generating new scripts with complete organizational standards compliance
2. **PSScriptAnalyzer Review**: Use `/review-psscriptanalyzer-compliance` prompt for comprehensive code review and compliance checking
3. **Complete Code Quality**: Use `/run-code-quality-scripts` prompt for comprehensive code quality processing including alignment analysis, Write-Host conversion, cleanup, and PSScriptAnalyzer checks
4. **Code Cleanup**: Use `/run-powershell-code-cleanup` prompt for safe whitespace and formatting fixes

These prompts ensure consistent adherence to all organizational standards and automate the development workflow.

---

## Output Colors and Formatting

| Color     | Usage                  | Implementation Notes               |
| --------- | ---------------------- | --------------------------------- |
| White     | Standard info          | Default informational content     |
| Cyan      | Process updates        | Background operations, scanning   |
| Green     | Success                | Completed operations, good status |
| Yellow    | Warnings               | Issues requiring attention        |
| Red       | Errors                 | Critical issues, failures         |
| Magenta   | Debug info             | Detailed troubleshooting info     |
| DarkGray  | Less important details | Secondary information            |

### Color Implementation Requirements:
- **Always use `Write-ColorOutput` function** instead of `Write-Host` for colored output
- **Support both PowerShell 5.1 and 7+** using ANSI codes for 7+ and console colors for 5.1
- **Include proper color reset** to avoid color bleeding
- **Use consistent color mapping** across all scripts
- **Log all colored output** to transcript files without color codes

### Color Function Template:
```powershell
function Write-ColorOutput {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$Color = "White"
    )

    if ($Script:UseAnsiColors) {
        # PowerShell 7+ with ANSI escape codes
        $colorCode = $Script:Colors[$Color]
        $resetCode = $Script:Colors.Reset
        Write-Output "${colorCode}${Message}${resetCode}"
    } else {
        # PowerShell 5.1 - Change console color, write output, then reset
        $originalColor = $Host.UI.RawUI.ForegroundColor
        try {
            if ($Script:Colors[$Color] -and $Script:Colors[$Color] -ne "") {
                $Host.UI.RawUI.ForegroundColor = $Script:Colors[$Color]
            }
            Write-Output $Message
        } finally {
            $Host.UI.RawUI.ForegroundColor = $originalColor
        }
    }
}
```

### Error Formatting Requirements:
- **Distinguish system errors from script errors** using clear prefixes like `[SYSTEM ERROR DETECTED]`
- **Include Event IDs and occurrence counts** for Windows Event Log errors
- **Use consistent error formatting** across all health reporting scripts
- **Color-code error severity** (Red for critical, Yellow for warnings)
