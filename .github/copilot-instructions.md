# =============================================================================
# File: copilot-instructions.md
# Created: 2025-02-07 21:21:53 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-20 17:10:00 UTC
# Updated By: jdyer-nuvodia
# Version: 3.1
# Additional Info: Added instruction for commit messages.
# =============================================================================

Always give me a commit message when writing, editing, or providing a script.
Make sure all additional functionlity and examples, parameters, etc. are included in the header of any script when updating and the versioning is increased accordingly each time it is updated.

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