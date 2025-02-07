# Create the new user
New-ADUser -Name "JB Dyer" -GivenName "JB" -Surname "Dyer" -SamAccountName "pa-jdyer" -UserPrincipalName "pa-jdyer@inlandtarp.com" -AccountPassword (ConvertTo-SecureString "6largehowlermonkeyS!" -AsPlainText -Force) -Enabled $true -Path "OU=Special Admins,DC=bch,DC=local"
# Add user to Domain Admins AD group.
Add-ADGroupMember -Identity "Domain Admins" -Members "pa-jdyer"

