# =============================================================================
# Script: Remove-GroupsFromDisabledUsers.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Version: 3.5.4
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
# =============================================================================

<#
.SYNOPSIS
    Removes all group memberships from disabled AD users and updates their description.
.DESCRIPTION
    This script performs the following actions on disabled AD user accounts:
     - Removes all group memberships (except default primary group)
     - Updates the user description with disabled date
     - Key actions are logged to a transcript file
     - Automatically opens the log file upon completion
     - Keeps the PowerShell window open until user interaction

    Supports -WhatIf parameter to preview changes without making them.

    Dependencies:
     - Active Directory PowerShell module
     - Appropriate AD permissions to modify users and groups
     - Windows Forms assembly for completion notification

    IMPORTANT: This script always runs in live mode and will make immediate changes to Active Directory unless -WhatIf is specified.
.PARAMETER None
    This script does not accept parameters. Configuration is done via variables.
.EXAMPLE
    .\Remove-GroupsFromDisabledUsers.ps1
    Processes all disabled users, logging actions to C:\Temp\DisabledUsers_[timestamp].log
.NOTES
    Security Level: High
    Required Permissions: Domain Admin or delegated AD permissions
    Validation Requirements:
    - Review log file after completion
    - Verify users have appropriate group membership
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param()


# Color support variables and Write-ColorOutput function
$Script:UseAnsiColors = $PSVersionTable.PSVersion.Major -ge 7
$Script:Colors = if ($Script:UseAnsiColors) {
    @{
        'White'    = "`e[37m"
        'Cyan'     = "`e[36m"
        'Green'    = "`e[32m"
        'Yellow'   = "`e[33m"
        'Red'      = "`e[31m"
        'Magenta'  = "`e[35m"
        'DarkGray' = "`e[90m"
        'Reset'    = "`e[0m"
    }
} else {
    @{
        'White'    = [ConsoleColor]::White
        'Cyan'     = [ConsoleColor]::Cyan
        'Green'    = [ConsoleColor]::Green
        'Yellow'   = [ConsoleColor]::Yellow
        'Red'      = [ConsoleColor]::Red
        'Magenta'  = [ConsoleColor]::Magenta
        'DarkGray' = [ConsoleColor]::DarkGray
        'Reset'    = ''
    }
}

function Write-ColorOutput {
    <#
    .SYNOPSIS
    Outputs colored text in a way that's compatible with PSScriptAnalyzer requirements.

    .DESCRIPTION
    This function provides colored output while maintaining compatibility with PSScriptAnalyzer
    by using only Write-Output and standard PowerShell cmdlets.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$Color = "White"
    )

    # Always use Write-Output to satisfy PSScriptAnalyzer
    # For PowerShell 7+, include ANSI color codes in the output
    if ($Script:UseAnsiColors) {
        $colorCode = $Script:Colors[$Color]
        $resetCode = $Script:Colors.Reset
        Write-Output "${colorCode}${Message}${resetCode}"
    } else {
        # For PowerShell 5.1, just output the message
        # Color formatting will be handled by the terminal/host if supported
        Write-Output $Message
    }
}

# Import required assembly for MessageBox
Add-Type -AssemblyName System.Windows.Forms

# Get the current domain
$CurrentDomain       = Get-ADDomain
# Default primary group
$DomainUsersGroup    = "Domain Users"

# Set log file location
$logfilename         = "C:\Temp\DisabledUsers_" + (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss") + ".log"

# Begin Logging
Start-Transcript -Path $logfilename

# Display script header
Write-ColorOutput -Message "=======================================================" -Color 'White'
Write-ColorOutput -Message "Remove-GroupsFromDisabledUsers.ps1" -Color 'White'
Write-ColorOutput -Message "=======================================================" -Color 'White'

# Display environment information
Write-ColorOutput -Message "CONFIGURATION:" -Color 'White'
Write-ColorOutput -Message "Current domain: $($CurrentDomain.DNSRoot)" -Color 'Cyan'
Write-ColorOutput -Message "Start time: $(Get-Date)" -Color 'White'
Write-ColorOutput -Message "Log file: $logfilename" -Color 'DarkGray'
Write-ColorOutput -Message "-------------------------------------------------------" -Color 'DarkGray'

# Get Domain Users group information
try {
    $DomainUsersInfo     = Get-ADGroup -Identity $DomainUsersGroup -Properties primaryGroupToken
    $DomainUsersPGID     = $DomainUsersInfo.primaryGroupToken
    Write-ColorOutput -Message "Domain Users group token: $DomainUsersPGID" -Color 'DarkGray'
} catch {
    Write-ColorOutput -Message "Error retrieving Domain Users group info: $($_.Exception.Message)" -Color 'Red'
    Stop-Transcript
    exit
}

# Get all disabled users
$DisabledUsers       = Search-ADAccount -AccountDisabled -UsersOnly -ResultPageSize 2000 -ResultSetSize $null
$DisabledUsersCount  = $DisabledUsers.Count

# Output the total number of users identified
Write-ColorOutput -Message "PROCESSING:" -Color 'White'
Write-ColorOutput -Message "Identified $DisabledUsersCount disabled user accounts" -Color 'Cyan'
Write-ColorOutput -Message "-------------------------------------------------------" -Color 'DarkGray'

# Counter for tracking progress
$UserCounter          = 0
$GroupsRemovedCounter = 0

try {
    # Loop through all users from Disabled Users search
    foreach ($User in $DisabledUsers.SamAccountName) {
        $UserCounter++
        # Calculate and display progress percentage
        $PercentComplete = [math]::Round(($UserCounter / $DisabledUsersCount) * 100, 1)
        Write-ColorOutput -Message "Processing user $UserCounter of $DisabledUsersCount ($PercentComplete%): ${ User}" -Color 'Cyan'

        try {
            # Get user details
            $UserInfo = Get-ADUser -Identity $User -Properties Description, PrimaryGroupID

            # Get all group memberships
            $UserGroups = Get-ADPrincipalGroupMembership $User
            $GroupCount = $UserGroups.Count
            Write-ColorOutput -Message "User is a member of $GroupCount groups" -Color 'DarkGray'

            # Process group removals
            if ($GroupCount -gt 0) {
                # Ensure Domain Users is the primary group before removing other groups
                if ($UserInfo.PrimaryGroupID -ne $DomainUsersPGID) {
                    try {
                        # Add to Domain Users if not already a member
                        if (-not ($UserGroups | Where-Object { $_.Name -eq $DomainUsersGroup })) {
                            if ($PSCmdlet.ShouldProcess($User, "Add to Domain Users group")) {
                                Add-ADGroupMember -Identity $DomainUsersGroup -Members $User
                            }
                        }
                        # Set Domain Users as primary group
                        if ($PSCmdlet.ShouldProcess($User, "Set Domain Users as primary group")) {
                            Set-ADUser -Identity $User -Replace @{ primaryGroupID = $DomainUsersPGID }
                            Write-ColorOutput -Message "Reset primary group to Domain Users" -Color 'Green'
                        }
                    } catch {
                        Write-ColorOutput -Message "Error setting Domain Users as primary group: $($_.Exception.Message)" -Color 'Red'
                    }
                }

                # Now remove all group memberships
                foreach ($Group in $UserGroups) {
                    # Skip if it's Domain Users and it's now the primary group
                    if ($Group.Name -ne $DomainUsersGroup -or $UserInfo.PrimaryGroupID -ne $DomainUsersPGID) {
                        try {
                            if ($PSCmdlet.ShouldProcess($User, "Remove from group '$($Group.Name)'")) {
                                Remove-ADGroupMember -Identity $Group.Name -Members $User -Confirm:$false
                                Write-ColorOutput -Message "Removed from group: $($Group.Name)" -Color 'Green'
                                $GroupsRemovedCounter++
                            }
                        } catch {
                            Write-ColorOutput -Message "Error removing from group $($Group.Name) - $($_.Exception.Message)" -Color 'Red'
                        }
                    }
                }
            } else {
                Write-ColorOutput -Message "User is not a member of any groups" -Color 'Yellow'
            }

            # Update User Description
            $DisabledDate = Get-Date
            if ($null -eq $UserInfo.Description -or -not $UserInfo.Description.StartsWith("User disabled")) {
                try {
                    if ($PSCmdlet.ShouldProcess($User, "Update description to 'User disabled $DisabledDate'")) {
                        Set-ADUser -Identity $User -Description "User disabled $DisabledDate"
                        Write-ColorOutput -Message "Updated user description" -Color 'Green'
                    }
                } catch {
                    Write-ColorOutput -Message "Error updating description: $($_.Exception.Message)" -Color 'Red'
                }
            } else {
                Write-ColorOutput -Message "User description already indicates disabled status" -Color 'Green'
            }
        } catch {
            Write-ColorOutput -Message "Error processing user ${ User} - $($_.Exception.Message)" -Color 'Red'
        }

        Write-ColorOutput -Message "-------------------------------------------------------" -Color 'DarkGray'
    }
} catch {
    Write-ColorOutput -Message "Critical error in main processing loop: $($_.Exception.Message)" -Color 'Red'
}

# End logging transcript
Write-ColorOutput -Message "SUMMARY:" -Color 'White'
Write-ColorOutput -Message "End time: $(Get-Date)" -Color 'White'
Write-ColorOutput -Message "Total users processed: $UserCounter of $DisabledUsersCount" -Color 'Cyan'
Write-ColorOutput -Message "Total group memberships removed: $GroupsRemovedCounter" -Color 'Green'
Write-ColorOutput -Message "Log file saved to: $logfilename" -Color 'DarkGray'
Write-ColorOutput -Message "=======================================================" -Color 'White'
Stop-Transcript

# Prompt end of script with log location
[System.Windows.Forms.MessageBox]::Show("Operation complete! Processed $UserCounter disabled users.`nRemoved $GroupsRemovedCounter group memberships.`n`nA log file has been saved to:`n$logfilename`n`nThe log file will open automatically when you click OK.", "Script Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

# Automatically open the log file
try {
    Invoke-Item -Path $logfilename
} catch {
    [System.Windows.Forms.MessageBox]::Show("Could not open log file: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}

# Keep console window open regardless of how script was launched
Write-ColorOutput -Message "`n`n" -Color "White" -NoNewline
Write-ColorOutput -Message "Script execution complete." -Color 'Green'
Write-ColorOutput -Message "Window will remain open for your review." -Color 'Cyan'
Write-ColorOutput -Message "Press any key to close this window..." -Color 'Yellow'

# This technique works better with "Run with PowerShell" than Read-Host
function Wait-ForKey {
    if ($psISE) {
        # Running in PowerShell ISE
        $null = Read-Host "Press Enter to continue..."
    } else {
        # Running in console or "Run with PowerShell"
        Write-ColorOutput -Message "Press any key to continue..." -Color "White"
        $host.UI.RawUI.ReadKey("NoEcho, IncludeKeyDown") | Out-Null
    }
}

# Execute the wait function to keep window open
Wait-ForKey
