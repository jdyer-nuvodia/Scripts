# =============================================================================
# File: copilot-instructions.md
# Created: 2025-02-07 21:21:53 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-27 18:49:00 UTC
# Updated By: jdyer-nuvodia
# Version: 3.1
# Additional Info: Added instruction for commit messages.
# =============================================================================

The current year is 2025.
When creating a script or file, make sure the header format below is used and the created date is set to the current timestamp, use a reliable time server to get the current time.
Always give me a commit message when writing, editing, or providing a script. Read this webpage for instructions on how to give me the best commit message you can: https://www.gitkraken.com/learn/git/best-practices/git-commit-message
Make sure all additional functionlity and examples, parameters, etc. are included in the header of any script when updating and the versioning is increased accordingly each time it is updated. Also update the last updated field in the header each time.

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