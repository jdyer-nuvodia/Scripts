# =============================================================================
# File: Initialize-Prompt-Scripts.txt
# Created: 2025-02-07 21:21:53 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-12 18:17:30 UTC
# Updated By: jdyer-nuvodia
# Version: 3.0
# Additional Info: Updated update notes filename format to include full timestamp
# =============================================================================

If you are GitHub Copilot reading this, use the remainder of this document for instructions on how to interact with me.

# -----------------------------------------------------------------------------
# MANDATORY SCRIPT FORMATTING REQUIREMENTS
# -----------------------------------------------------------------------------
1. All scripts MUST follow these formatting rules without deviation:
   - An example script can be found in the repository entitled Script-Template.ps1
   - Minimize unnecessary whitespace - no empty lines between related code blocks
   - Use single empty lines only between major sections or logical breaks
   - Align parameter blocks and variable assignments for readability
   - Use concise, descriptive variable names following PowerShell conventions
   - Allow new lines in command parameters or multi-line declarations when it affects syntax to prevent parsing errors.
   - When default value assignments in parameter blocks cause parsing errors, it is allowed to omit them provided that the explicit default values are defined elsewhere in the script.
   - Place opening braces on the same line as the condition/function declaration
   - Include error handling in try/catch blocks for critical operations
   - Use consistent indentation (4 spaces recommended)
   - Group related operations into logical blocks
   - Place closing braces on their own lines
   - IMPORTANT: Minimizing scripts MUST NOT remove any functionality, including additional logging, extended error handling, or extra features, unless such changes are explicitly discussed first.
   - IMPORTANT: DO NOT remove explicitly defined variables in favor of parameters passed in at runtime.
    
2. Complete Script Requirements (MANDATORY):
   - ALL scripts MUST be provided in their complete form
   - NO partial scripts or code snippets are allowed
   - NEVER use phrases like "Rest of script remains the same" or "..." to indicate omitted content
   - ALWAYS provide the ENTIRE script content when making updates or changes
   - When updating a script, the COMPLETE updated version MUST be provided
   - Scripts MUST include proper error handling
   - Scripts MUST include proper logging mechanisms
   - Scripts MUST include all necessary function declarations
   - Scripts MUST include all required parameter declarations
   - Scripts MUST include proper completion handling
   - Scripts MUST include proper resource cleanup where applicable
   - Scripts MUST enforce TCP connections for network operations to ensure data integrity
   - Scripts MUST include connection validation and retry logic
   - IMPORTANT: Any reference to existing content being unchanged MUST still include that content in full
    
3. Long Script Handling (MANDATORY):
   - For scripts exceeding 100 lines, content MUST be provided via GitHub Gist
   - When providing a Gist, the following MUST be included:
     * The complete Gist URL
     * A brief description of the script's purpose
     * Any relevant setup or usage instructions
   - The Gist MUST maintain all formatting requirements specified in this document
   - The Gist MUST include proper file naming and extension
   - IMPORTANT: Do not split long scripts into multiple chunks
   - IMPORTANT: Ensure the entire script is available in a single Gist
   - For scripts under 100 lines, content may be provided directly in the conversation
   - ALL scripts, regardless of length, MUST maintain proper formatting and documentation
   - IMPORTANT: Even when using Gists, the ENTIRE script content must be included

4. Repository Update Notes (MANDATORY):
   - An example repository update can be found in the repository entitled Update-Notes.txt
   - ALL changes MUST be documented in a separate txt file, regardless of file type.
   - Update notes MUST be provided for all executable code changes, as well as for policy and documentation enhancements or any file within the repository.
   - Update notes MUST follow the template format below.
   - NO commits are allowed without corresponding update notes.
   - Update notes MUST be created before script implementation.
    
# -----------------------------------------------------------------------------
# MANDATORY SCRIPT HEADER FORMAT
# -----------------------------------------------------------------------------
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

# -----------------------------------------------------------------------------
# REPOSITORY UPDATE NOTES TEMPLATE (MANDATORY)
# -----------------------------------------------------------------------------
# Update notes MUST be provided in a separate txt file using this format:
#
# -----------------------------------------------------------------------------
# Date: <YYYY-MM-DD HH:MM:SS UTC>
# Commit: <Brief description of the commit or change>
# Author: <AuthorName or Collaborator>
#
# Summary of Changes:
# - <List each key change or improvement>
#
# Impact:
# - <Note how these changes affect functionality, dependencies, or users>
# - <Highlight any potential risks or required follow-up actions>
#
# File Name: <ScriptName>_UpdateNotes_<YYYYMMDD_HHMMSS>.txt
# -----------------------------------------------------------------------------