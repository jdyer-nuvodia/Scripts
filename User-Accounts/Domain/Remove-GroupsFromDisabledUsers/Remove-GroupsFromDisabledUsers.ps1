# =============================================================================
# Script: Remove-GroupsFromDisabledUsers.ps1
# Created: 2024-02-20 17:15:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-05 23:05:25 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2
# Additional Info: Modified to dynamically find any group with 'disabled' in the name
# =============================================================================

<#
.SYNOPSIS
    Manages disabled AD user accounts by removing group memberships and moving them to a designated OU.
.DESCRIPTION
    This script performs the following actions on disabled AD user accounts:
     - Sets a "disabled" group as the primary group (dynamically finds groups containing "disabled")
     - Removes all other group memberships
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

# Default group name if none found
$DefaultDisabledGroupName = "DisabledUsers"
# Get the current domain
$CurrentDomain = Get-ADDomain
$DomainDN = $CurrentDomain.DistinguishedName

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

# Check if target OU exists, create if needed
if (-not $ReportOnly) {
    try {
        $OUExists = Get-ADOrganizationalUnit -Identity $TargetOU -ErrorAction SilentlyContinue
        if (-not $OUExists) {
            Write-Host "Creating target OU: $TargetOU" -ForegroundColor Yellow
            New-ADOrganizationalUnit -Name "Disabled Users" -Path $DomainDN
        }
    }
    catch {
        Write-Host "Error checking/creating target OU: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Find any group with 'disabled' in the name (case-insensitive)
try {
    Write-Host "Searching for groups containing 'disabled' in the name..." -ForegroundColor Cyan
    $DisabledGroups = Get-ADGroup -Filter "Name -like '*disabled*'" -Properties primarygrouptoken -ErrorAction SilentlyContinue
    
    # Check if any groups were found
    if ($DisabledGroups -and $DisabledGroups.Count -gt 0) {
        # If multiple groups found, use the first one
        if ($DisabledGroups -is [array]) {
            $PrimaryGroup = $DisabledGroups[0]
            Write-Host "Found multiple groups with 'disabled' in the name. Using: $($PrimaryGroup.Name)" -ForegroundColor Yellow
        } else {
            # Just one group found
            $PrimaryGroup = $DisabledGroups
            Write-Host "Found group with 'disabled' in the name: $($PrimaryGroup.Name)" -ForegroundColor Green
        }
        
        $PrimaryGroupName = $PrimaryGroup.Name
        $PrimaryGroupToken = $PrimaryGroup | Select-Object primarygrouptoken
    } else {
        Write-Host "No groups with 'disabled' in the name found." -ForegroundColor Yellow
        
        if (-not $ReportOnly) {
            Write-Host "Creating group '$DefaultDisabledGroupName'..." -ForegroundColor Cyan
            $PrimaryGroup = New-ADGroup -Name $DefaultDisabledGroupName -SamAccountName $DefaultDisabledGroupName `
                           -GroupCategory Security -GroupScope Global -DisplayName $DefaultDisabledGroupName -PassThru
            
            # Get the primarygrouptoken after creation
            $PrimaryGroupToken = Get-ADGroup -Identity $DefaultDisabledGroupName -Properties primarygrouptoken | Select-Object primarygrouptoken
            $PrimaryGroupName = $DefaultDisabledGroupName
            Write-Host "Created new disabled users group: $DefaultDisabledGroupName" -ForegroundColor Green
        }
        else {
            Write-Host "Running in report-only mode. Group would be created in actual run." -ForegroundColor Yellow
            Write-Host "Exiting script as a disabled group is required." -ForegroundColor Red
            Stop-Transcript
            exit
        }
    }
}
catch {
    Write-Host "Error searching for disabled groups: $($_.Exception.Message)" -ForegroundColor Red
    Stop-Transcript
    exit
}

# Get all disabled users
$DisabledUsers = Search-ADAccount -AccountDisabled -UsersOnly -ResultPageSize 2000 -ResultSetSize $null
$DisabledUsersCount = $DisabledUsers.Count

# Output the total number of users identified
Write-Host "$(Get-Date): Identified $DisabledUsersCount disabled user accounts" -ForegroundColor Cyan
Write-Host ("-" * 100) -ForegroundColor DarkGray

Try {
    # Loop through all users from Disabled Users search
    Foreach ($User in $DisabledUsers.SamAccountName) {
        # Get OU location from User DN and Primary Group ID
        $UserDN = Get-ADUser -Identity $User | Select-Object DistinguishedName
        $UserOU = $UserDN.DistinguishedName.Substring($UserDN.DistinguishedName.IndexOf('OU=', [System.StringComparison]::CurrentCultureIgnoreCase))
        $UserPGID = Get-ADUser -Identity $User -Properties PrimaryGroupID

        # Logging
        Write-Host "$(Get-Date): Updating user $User" -ForegroundColor Cyan

        # Add user as member of Primary Group
        $UserGroups = Get-ADPrincipalGroupMembership $User | Select-Object name, groupscope

        # Set Primary Group to the disabled group
        If ($UserPGID.PrimaryGroupID -ne $PrimaryGroupToken.primarygrouptoken) {
            Write-Host "$(Get-Date): Updating Primary Group to $PrimaryGroupName | $($PrimaryGroupToken.primarygrouptoken)" -ForegroundColor Yellow
            If (-not $ReportOnly) {
                Try {
                    Add-ADGroupMember -Identity $PrimaryGroupName -Members $User
                    Get-ADUser $User | Set-ADUser -Replace @{primaryGroupID = $PrimaryGroupToken.primarygrouptoken}
                    Write-Host "$(Get-Date): Successfully set primary group" -ForegroundColor Green
                } Catch {
                    Write-Host "$(Get-Date): $($_.Exception.Message)" -ForegroundColor Red
                }
            }
        } Else {
            Write-Host "$(Get-Date): $User Already in $PrimaryGroupName" -ForegroundColor Green
        }

        # Remove all groups but the disabled group
        Foreach ($UserGroup in $UserGroups) {
            If ($UserGroup.Name -ne $PrimaryGroupName) {
                If (-not $ReportOnly) {
                    Try {
                        Remove-ADGroupMember -Identity $UserGroup.Name -Members $User -Confirm:$false
                        Write-Host "$(Get-Date): Successfully removed from group: $($UserGroup.Name)" -ForegroundColor Green
                    } Catch {
                        Write-Host "$(Get-Date): Error removing from group $($UserGroup.Name): $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
                Write-Host "$(Get-Date): -Removing Group: $($UserGroup.Name)" -ForegroundColor Yellow
            }
        }

        # Move user object to Target OU for disabled users
        If ($UserOU -ne $TargetOU) {
            If (-not $ReportOnly) {
                Try {
                    Move-ADObject -Identity (Get-ADUser -Identity $User) -TargetPath $TargetOU 
                    Write-Host "$(Get-Date): Successfully moved user to target OU" -ForegroundColor Green
                } Catch {
                    Write-Host "$(Get-Date): Error moving user: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
            Write-Host "$(Get-Date): Moving $User To $TargetOU" -ForegroundColor Yellow
        } Else {
            Write-Host "$(Get-Date): User already in $TargetOU" -ForegroundColor Green
        }

        # Update User Description
        $UserDescription = Get-ADUser -Identity $User -Properties Description
        $DisabledDate = Get-Date
        If ($null -eq $UserDescription.Description -or -not $UserDescription.Description.StartsWith("User disabled")) {
            If (-not $ReportOnly) {
                Try {
                    Set-ADUser -Identity $User -Description "User disabled $DisabledDate"
                    Write-Host "$(Get-Date): Description updated successfully" -ForegroundColor Green
                } Catch {
                    Write-Host "$(Get-Date): Error updating description: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
            Write-Host "$(Get-Date): User description updated to read 'User disabled $DisabledDate'" -ForegroundColor Yellow
        }

        Write-Host ("-" * 100) -ForegroundColor DarkGray
    }
} Catch {
    Write-Host "$(Get-Date): $($_.Exception.Message)" -ForegroundColor Red
}

# End logging transcript
Stop-Transcript

# Prompt end of script with log location
[System.Windows.Forms.MessageBox]::Show("Operation is complete, you can view the logs at $logfilename", "", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
