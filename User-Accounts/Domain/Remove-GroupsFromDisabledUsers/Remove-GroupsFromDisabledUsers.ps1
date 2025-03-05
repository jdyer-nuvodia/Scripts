# =============================================================================
# Script: Remove-GroupsFromDisabledUsers.ps1
# Created: 2024-02-20 17:15:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-05 23:10:30 UTC
# Updated By: jdyer-nuvodia
# Version: 2.0
# Additional Info: Simplified to remove all group memberships from disabled users
# =============================================================================

<#
.SYNOPSIS
    Removes all group memberships from disabled AD users and moves them to a designated OU.
.DESCRIPTION
    This script performs the following actions on disabled AD user accounts:
     - Removes all group memberships (except default primary group)
     - Moves the user to a designated Disabled Users OU
     - Updates the user description with disabled date
     - Key actions are logged to a transcript file
     
    Dependencies:
     - Active Directory PowerShell module
     - Appropriate AD permissions to modify users and groups
     - Windows Forms assembly for completion notification
.PARAMETER None
    This script does not accept parameters. Configuration is done via variables.
.EXAMPLE
    .\Remove-GroupsFromDisabledUsers.ps1
    Processes all disabled users, logging actions to C:\Temp\DisabledUsers_[timestamp].log
.NOTES
    Security Level: High
    Required Permissions: Domain Admin or delegated AD permissions
    Validation Requirements: 
    - Verify $ReportOnly is set correctly before execution
    - Review log file after completion
    - Verify users are in correct OU with appropriate group membership
#>

# Import required assembly for MessageBox
Add-Type -AssemblyName System.Windows.Forms

# Get the current domain
$CurrentDomain = Get-ADDomain
$DomainDN = $CurrentDomain.DistinguishedName
$DomainUsersGroup = "Domain Users" # Default primary group

# Set target OU based on detected domain
$TargetOU = "OU=Disabled Users,$DomainDN"
$logfilename = "C:\Temp\DisabledUsers_" + (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss") + ".log"
$ReportOnly = $true  # Change this to $false to actively make changes to AD

# Begin Logging
Start-Transcript -Path $logfilename

# Display environment information
Write-Host "Current domain: $($CurrentDomain.DNSRoot)" -ForegroundColor Cyan
Write-Host "Target OU: $TargetOU" -ForegroundColor Cyan
Write-Host "Report only mode: $ReportOnly" -ForegroundColor Cyan
Write-Host "Start time: $(Get-Date)" -ForegroundColor White

# Check if target OU exists, create if needed
if (-not $ReportOnly) {
    try {
        $OUExists = Get-ADOrganizationalUnit -Identity $TargetOU -ErrorAction SilentlyContinue
        if (-not $OUExists) {
            Write-Host "Creating target OU: $TargetOU" -ForegroundColor Yellow
            New-ADOrganizationalUnit -Name "Disabled Users" -Path $DomainDN
            Write-Host "Successfully created Disabled Users OU" -ForegroundColor Green
        } else {
            Write-Host "Target OU already exists" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "Error checking/creating target OU: $($_.Exception.Message)" -ForegroundColor Red
    }
}

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
Write-Host "Identified $DisabledUsersCount disabled user accounts" -ForegroundColor Cyan
Write-Host ("-" * 100) -ForegroundColor DarkGray

# Counter for tracking progress
$UserCounter = 0

Try {
    # Loop through all users from Disabled Users search
    Foreach ($User in $DisabledUsers.SamAccountName) {
        $UserCounter++
        # Calculate and display progress percentage
        $PercentComplete = [math]::Round(($UserCounter / $DisabledUsersCount) * 100, 1)
        Write-Host "Processing user $UserCounter of $DisabledUsersCount ($PercentComplete%): $User" -ForegroundColor Cyan
        
        try {
            # Get user details
            $UserInfo = Get-ADUser -Identity $User -Properties DistinguishedName, PrimaryGroupID, Description
            $UserOU = $UserInfo.DistinguishedName.Substring($UserInfo.DistinguishedName.IndexOf('OU=', [System.StringComparison]::CurrentCultureIgnoreCase))
            
            # Get all group memberships
            $UserGroups = Get-ADPrincipalGroupMembership $User
            $GroupCount = $UserGroups.Count
            Write-Host "User is a member of $GroupCount groups" -ForegroundColor DarkGray
            
            # Process group removals
            if ($GroupCount -gt 0) {
                if (-not $ReportOnly) {
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
                            } catch {
                                Write-Host "Error removing from group $($Group.Name): $($_.Exception.Message)" -ForegroundColor Red
                            }
                        }
                    }
                } else {
                    # Report mode - just list the groups
                    Write-Host "Would remove user from the following groups:" -ForegroundColor Yellow
                    foreach ($Group in $UserGroups) {
                        Write-Host " - $($Group.Name)" -ForegroundColor DarkGray
                    }
                }
            } else {
                Write-Host "User is not a member of any groups" -ForegroundColor Yellow
            }
            
            # Move user object to Target OU for disabled users
            if ($UserOU -ne $TargetOU) {
                if (-not $ReportOnly) {
                    try {
                        Move-ADObject -Identity $UserInfo.DistinguishedName -TargetPath $TargetOU 
                        Write-Host "Successfully moved user to Disabled Users OU" -ForegroundColor Green
                    } catch {
                        Write-Host "Error moving user: $($_.Exception.Message)" -ForegroundColor Red
                    }
                } else {
                    Write-Host "Would move user to: $TargetOU" -ForegroundColor Yellow
                }
            } else {
                Write-Host "User already in Disabled Users OU" -ForegroundColor Green
            }
            
            # Update User Description
            $DisabledDate = Get-Date
            if ($null -eq $UserInfo.Description -or -not $UserInfo.Description.StartsWith("User disabled")) {
                if (-not $ReportOnly) {
                    try {
                        Set-ADUser -Identity $User -Description "User disabled $DisabledDate"
                        Write-Host "Updated user description" -ForegroundColor Green
                    } catch {
                        Write-Host "Error updating description: $($_.Exception.Message)" -ForegroundColor Red
                    }
                } else {
                    Write-Host "Would update description to: User disabled $DisabledDate" -ForegroundColor Yellow
                }
            } else {
                Write-Host "User description already indicates disabled status" -ForegroundColor Green
            }
        } catch {
            Write-Host "Error processing user $User: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        Write-Host ("-" * 100) -ForegroundColor DarkGray
    }
} catch {
    Write-Host "Critical error in main processing loop: $($_.Exception.Message)" -ForegroundColor Red
}

# End logging transcript
Write-Host "End time: $(Get-Date)" -ForegroundColor White
Write-Host "Total users processed: $UserCounter of $DisabledUsersCount" -ForegroundColor Cyan
Stop-Transcript

# Prompt end of script with log location
[System.Windows.Forms.MessageBox]::Show("Operation complete! Processed $UserCounter disabled users. View logs at:`n$logfilename", "Script Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
