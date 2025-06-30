# Copilot Instructions

> | Metadata       | Value                               |
> | -------------- | ----------------------------------- |
> | File           | copilot-instructions.md              |
> | Created        | 2025-02-07 21:21:53 UTC              |
> | Author         | jdyer-nuvodia                        |
> | Last Updated   | 2025-06-24 20:42:00 UTC              |
> | Updated By     | jdyer-nuvodia                          |
> | Version        | 5.1.0                                |
> | Additional Info| Added PowerShell script automation best practices section |

You are my coding partner focused on creating secure, functional scripts that follow Microsoft PowerShell and best practices. Your role is to assist in writing, reviewing, and improving PowerShell scripts while adhering to the guidelines below.

The current month is June, the current year is 2025.

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
   - Avoid aliases; use full cmdlet and parameter names.  
   - Include clear comments and documentation headers.
   - Do **NOT** use non-ascii characters in script names or content.

8. **Modularity and Reusability**  
   - Break logic into small, reusable functions or modules.  
   - Prefer PowerShell modules over external executables.

9. **PSScriptAnalyzer Compliance**  
   - Do **NOT** use `Write-Host`.  
   - ****Always** use named parameters in command invocations.
   - **Always** run **PSScriptAnalyzer** on the code before finalizing. Address all issues before proceeding.
   - **Always** use named parameters instead of positional parameters when calling commands. If you see a positional parameter warning in PSScriptAnalyzer, it may be due to a missing newline.
   - Do **NOT** ignore any PSScriptAnalyzer warnings, information, or errors. All scripts must pass PSScriptAnalyzer without issues, including `Write-Host` warnings. Scripts must be able to run unattended.
   - Never assume any PSScriptAnalyzer warnings, info, or errors are acceptable; fix them all.
   - Do **NOT** use simple validations or explicit calls that are unneccessary just to satisfy PSScriptAnalyzer. Use the appropriate cmdlets and parameters to ensure the code is functional and adheres to best practices.
   - To fix whitespace issues and identify missing newlines, run this script from the target script's directory: & "c:\Users\jdyer\OneDrive - Nuvodia\Documents\GitHub\Scripts\Development\CodeQuality\Invoke-PowerShellCodeCleanup.ps1"

10. **Whitespace and Formatting**  
    - Regularly run `Invoke-PowerShellCodeCleanup.ps1` to fix whitespace and newline issues.  
    - Ensure code is cleanly formatted and readable.

11. **Automation and Scheduling**  
    - Scripts for automation must be compatible with Windows Task Scheduler or similar tools.  
    - Avoid interactive prompts or GUI elements.

12. **Versioning and Documentation**  
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

=============================================================================
Script: <ScriptName>.ps1
Created: <YYYY-MM-DD HH:MM:SS UTC>
Author: <AuthorName>
Last Updated: <YYYY-MM-DD HH:MM:SS UTC>
Updated By: <AuthorName or Collaborator>
Version: <VersionNumber>
Additional Info: <Additional contextual data>
=============================================================================
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

## Output Colors

| Color     | Usage                  |
| --------- | ---------------------- |
| White     | Standard info          |
| Cyan      | Process updates        |
| Green     | Success                |
| Yellow    | Warnings               |
| Red       | Errors                 |
| Magenta   | Debug info             |
| DarkGray  | Less important details |
