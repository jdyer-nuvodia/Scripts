# =============================================================================
# Script: Remove-GroupsFromDisabledUsers.ps1
# Created: 2024-02-20 17:15:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-05 23:17:15 UTC
# Updated By: jdyer-nuvodia
# Version: 3.0
# Additional Info: Removed report-only mode, script now always runs in live mode
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
     
    Dependencies:
     - Active Directory PowerShell module
     - Appropriate AD permissions to modify users and groups
     - Windows Forms assembly for completion notification

    IMPORTANT: This script always runs in live mode and will make immediate changes to Active Directory.
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

# Import required assembly for MessageBox
Add-Type -AssemblyName System.Windows.Forms

# Get the current domain
$CurrentDomain = Get-ADDomain
$DomainDN = $CurrentDomain.DistinguishedName
$DomainUsersGroup = "Domain Users" # Default primary group

# Set log file location
$logfilename = "C:\Temp\DisabledUsers_" + (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss") + ".log"

# Begin Logging
Start-Transcript -Path $logfilename

# Display script header
Write-Host "=======================================================" -ForegroundColor White
Write-Host "Remove-GroupsFromDisabledUsers.ps1" -ForegroundColor White
Write-Host "=======================================================" -ForegroundColor White

# Display environment information
Write-Host "CONFIGURATION:" -ForegroundColor White
Write-Host "Current domain: $($CurrentDomain.DNSRoot)" -ForegroundColor Cyan
Write-Host "Start time: $(Get-Date)" -ForegroundColor White
Write-Host "Log file: $logfilename" -ForegroundColor DarkGray
Write-Host "-------------------------------------------------------" -ForegroundColor DarkGray

# Get Domain Users group information
try {
    $DomainUsersInfo = Get-ADGroup -Identity $DomainUsersGroup -Properties primaryGroupToken
    $DomainUsersPGID = $DomainUsersInfo.primaryGroupToken
    Write-Host "Domain Users group token: $DomainUsersPGID" -ForegroundColor DarkGray
} catch {
    Write-Host "Error retrieving Domain Users group info: $($_.Exception.Message)" -ForegroundColor Red
    Stop-Transcript
    exit
}

# Get all disabled users
$DisabledUsers = Search-ADAccount -AccountDisabled -UsersOnly -ResultPageSize 2000 -ResultSetSize $null
$DisabledUsersCount = $DisabledUsers.Count

# Output the total number of users identified
Write-Host "PROCESSING:" -ForegroundColor White
Write-Host "Identified $DisabledUsersCount disabled user accounts" -ForegroundColor Cyan
Write-Host "-------------------------------------------------------" -ForegroundColor DarkGray

# Counter for tracking progress
$UserCounter = 0
$GroupsRemovedCounter = 0

Try {
    # Loop through all users from Disabled Users search
    Foreach ($User in $DisabledUsers.SamAccountName) {
        $UserCounter++
        # Calculate and display progress percentage
        $PercentComplete = [math]::Round(($UserCounter / $DisabledUsersCount) * 100, 1)
        Write-Host "Processing user $UserCounter of $DisabledUsersCount ($PercentComplete%): $User" -ForegroundColor Cyan
        
        try {
            # Get user details
            $UserInfo = Get-ADUser -Identity $User -Properties Description, PrimaryGroupID
            
            # Get all group memberships
            $UserGroups = Get-ADPrincipalGroupMembership $User
            $GroupCount = $UserGroups.Count
            Write-Host "User is a member of $GroupCount groups" -ForegroundColor DarkGray
            
            # Process group removals
            if ($GroupCount -gt 0) {
                # Ensure Domain Users is the primary group before removing other groups
                if ($UserInfo.PrimaryGroupID -ne $DomainUsersPGID) {
                    try {
                        # Add to Domain Users if not already a member
                        if (-not ($UserGroups | Where-Object { $_.Name -eq $DomainUsersGroup })) {
                            Add-ADGroupMember -Identity $DomainUsersGroup -Members $User
                        }
                        # Set Domain Users as primary group
                        Set-ADUser -Identity $User -Replace @{primaryGroupID = $DomainUsersPGID}
                        Write-Host "Reset primary group to Domain Users" -ForegroundColor Green
                    } catch {
                        Write-Host "Error setting Domain Users as primary group: $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
                
                # Now remove all group memberships
                foreach ($Group in $UserGroups) {
                    # Skip if it's Domain Users and it's now the primary group
                    if ($Group.Name -ne $DomainUsersGroup -or $UserInfo.PrimaryGroupID -ne $DomainUsersPGID) {
                        try {
                            Remove-ADGroupMember -Identity $Group.Name -Members $User -Confirm:$false
                            Write-Host "Removed from group: $($Group.Name)" -ForegroundColor Green
                            $GroupsRemovedCounter++
                        } catch {
                            Write-Host "Error removing from group $($Group.Name): $($_.Exception.Message)" -ForegroundColor Red
                        }
                    }
                }
            } else {
                Write-Host "User is not a member of any groups" -ForegroundColor Yellow
            }
            
            # Update User Description
            $DisabledDate = Get-Date
            if ($null -eq $UserInfo.Description -or -not $UserInfo.Description.StartsWith("User disabled")) {
                try {
                    Set-ADUser -Identity $User -Description "User disabled $DisabledDate"
                    Write-Host "Updated user description" -ForegroundColor Green
                } catch {
                    Write-Host "Error updating description: $($_.Exception.Message)" -ForegroundColor Red
                }
            } else {
                Write-Host "User description already indicates disabled status" -ForegroundColor Green
            }
        } catch {
            Write-Host "Error processing user $User: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        Write-Host "-------------------------------------------------------" -ForegroundColor DarkGray
    }
} catch {
    Write-Host "Critical error in main processing loop: $($_.Exception.Message)" -ForegroundColor Red
}

# End logging transcript
Write-Host "SUMMARY:" -ForegroundColor White
Write-Host "End time: $(Get-Date)" -ForegroundColor White
Write-Host "Total users processed: $UserCounter of $DisabledUsersCount" -ForegroundColor Cyan
Write-Host "Total group memberships removed: $GroupsRemovedCounter" -ForegroundColor Green
Write-Host "Log file saved to: $logfilename" -ForegroundColor DarkGray
Write-Host "=======================================================" -ForegroundColor White
Stop-Transcript

# Prompt end of script with log location
[System.Windows.Forms.MessageBox]::Show("Operation complete! Processed $UserCounter disabled users.`nRemoved $GroupsRemovedCounter group memberships.`n`nA log file has been saved to:`n$logfilename`n`nThe log file will open automatically when you click OK.", "Script Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

# Automatically open the log file
try {
    Invoke-Item -Path $logfilename
} catch {
    [System.Windows.Forms.MessageBox]::Show("Could not open log file: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}
