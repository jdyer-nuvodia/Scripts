# =============================================================================
# Script: Delete-ShadowCopies.ps1
# Created: 2024-02-20 17:15:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2024-02-25 23:07:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0
# Additional Info: Initial script creation with standard header
# =============================================================================

<#
.SYNOPSIS
    Manages Windows Volume Shadow Copies by keeping only the newest restore point.
.DESCRIPTION
    This script performs the following actions:
    - Lists all existing volume shadow copies
    - Identifies the newest shadow copy
    - Deletes all shadow copies except the most recent one
    - Uses vssadmin commands for shadow copy management
    
    Dependencies:
    - Requires administrative privileges
    - Windows Volume Shadow Copy Service must be running
.EXAMPLE
    .\Delete-ShadowCopies.ps1
    Keeps the newest shadow copy and removes all others.
.NOTES
    Security Level: High
    Required Permissions: Administrative privileges
    Validation Requirements: Verify shadow copy retention after execution
#>

# List all shadow copies
$vssList = vssadmin list shadows

# Filter out the newest restore point and delete the rest
$shadowCopies = $vssList | Where-Object {$_ -match "Shadow Copy ID:"}
$shadowIds = $shadowCopies | ForEach-Object { $_.Split(":")[1].Trim() }

# Keep only the newest restore point (assuming it's the first in the list)
$keepId = $shadowIds[0]

# Delete all other restore points
foreach ($id in $shadowIds | Where-Object {$_ -ne $keepId}) {
    vssadmin delete shadows /shadow=$id /quiet
}