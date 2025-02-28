# Set static variables
$PrimaryGroupName = "DisabledPrimary"
$TargetOU = "OU=Disabled Users,DC=YourDomain,DC=com"
$logfilename = "C:\Temp\DisabledUsers_" + (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss") + ".log"
$ReportOnly = $true  # Change this to $false to actively make changes to AD

# Begin Logging
Start-Transcript -Path $logfilename

# Get Primary Group token and all disabled users
$PrimaryGroupToken = Get-ADGroup -Identity $PrimaryGroupName -Properties primarygrouptoken | Select-Object primarygrouptoken
$DisabledUsers = Search-ADAccount -AccountDisabled -UsersOnly -ResultPageSize 2000 -ResultSetSize $null
$DisabledUsersCount = $DisabledUsers.Count

# Output the total number of users identified
Write-Output "$(Get-Date): Identified $DisabledUsersCount disabled user accounts"
Write-Output "-" * 100

Try {
    # Loop through all users from Disabled Users search
    Foreach ($User in $DisabledUsers.SamAccountName) {
        # Get OU location from User DN and Primary Group ID
        $UserDN = Get-ADUser -Identity $User | Select-Object DistinguishedName
        $UserOU = $UserDN.DistinguishedName.Substring($UserDN.DistinguishedName.IndexOf('OU=', [System.StringComparison]::CurrentCultureIgnoreCase))
        $UserPGID = Get-ADUser -Identity $User -Properties PrimaryGroupID

        # Logging
        Write-Output "$(Get-Date): Updating user $User"

        # Add user as member of Primary Group to DisabledPrimary
        $UserGroups = Get-ADPrincipalGroupMembership $User | Select-Object name, groupscope

        # Set Primary Group to DisabledPrimary
        If ($UserPGID.PrimaryGroupID -ne $PrimaryGroupToken.primarygrouptoken) {
            Write-Output "$(Get-Date): Updating Primary Group to $PrimaryGroupName | $($PrimaryGroupToken.primarygrouptoken)"
            If (-not $ReportOnly) {
                Try {
                    Add-ADGroupMember -Identity $PrimaryGroupName -Members $User
                    Get-ADUser $User | Set-ADUser -Replace @{primaryGroupID = $PrimaryGroupToken.primarygrouptoken}
                } Catch {
                    Write-Output "$(Get-Date): $($_.Exception.GetType().FullName)"
                }
            }
        } Else {
            Write-Output "$(Get-Date): $User Already in $PrimaryGroupName"
        }

        # Remove all groups but DisabledPrimary
        Foreach ($UserGroup in $UserGroups) {
            If ($UserGroup.Name -ne $PrimaryGroupName) {
                If (-not $ReportOnly) {
                    Try {
                        Remove-ADGroupMember -Identity $UserGroup.Name -Members $User -Confirm:$false
                    } Catch {
                        Write-Output "$(Get-Date): $($_.Exception.GetType().FullName)"
                    }
                }
                Write-Output "$(Get-Date): -Removing Group: $($UserGroup.Name)"
            }
        }

        # Move user object to Target OU for disabled users
        If ($UserOU -ne $TargetOU) {
            If (-not $ReportOnly) {
                Try {
                    Move-ADObject -Identity (Get-ADUser -Identity $User) -TargetPath $TargetOU -Verbose
                } Catch {
                    Write-Output "$(Get-Date): $($_.Exception.GetType().FullName)"
                }
            }
            Write-Output "$(Get-Date): Moving $User To $TargetOU"
        } Else {
            Write-Output "$(Get-Date): User already in $TargetOU"
        }

        # Update User Description
        $UserDescription = Get-ADUser -Identity $User -Properties Description
        $DisabledDate = Get-Date
        If ($null -eq $UserDescription.Description -or -not $UserDescription.Description.StartsWith("User disabled")) {
            If (-not $ReportOnly) {
                Try {
                    Set-ADUser -Identity $User -Description "User disabled $DisabledDate"
                } Catch {
                    Write-Output "$(Get-Date): $($_.Exception.GetType().FullName)"
                }
            }
            Write-Output "$(Get-Date): User description updated to read 'User disabled $DisabledDate'"
        }

        Write-Output "-" * 100
    }
} Catch {
    Write-Output "$(Get-Date): $($_.Exception.GetType().FullName)"
}

# End logging transcript
Stop-Transcript

# Prompt end of script with log location
[System.Windows.Forms.MessageBox]::Show("Operation is complete, you can view the logs at $logfilename", "", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
