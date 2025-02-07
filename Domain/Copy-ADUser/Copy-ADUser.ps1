# Define the source user and the new user's details
$sourceUser = "pa-gbullock"   # Replace with the username of the user to be copied
$newUserName = "pa-jdyer"     # Replace with the new user's username
$newUserGivenName = "JB"  # Replace with the new user's given name
$newUserSurname = "Dyer"      # Replace with the new user's surname
$newUserPassword = "12ravenousgiantpandaS!" # Replace with the new user's password
$newUserDescription = "Nuvodia" # Replace with the new user's description

# Load the Active Directory module
Import-Module ActiveDirectory

# Get the source user's details
$sourceUserDetails = Get-ADUser -Identity $sourceUser -Properties *

# Create the new user with the different name properties and description
New-ADUser `
    -Name "$newUserGivenName $newUserSurname" `
    -GivenName $newUserGivenName `
    -Surname $newUserSurname `
    -SamAccountName $newUserName `
    -UserPrincipalName "$newUserName@$(($sourceUserDetails.UserPrincipalName).Split('@')[1])" `
    -Path $sourceUserDetails.DistinguishedName `
    -Enabled $true `
    -AccountPassword (ConvertTo-SecureString $newUserPassword -AsPlainText -Force) `
    -Description $newUserDescription

# Add the new user to the same groups as the source user
$sourceUserGroups = Get-ADUser -Identity $sourceUser -Properties MemberOf | Select-Object -ExpandProperty MemberOf
foreach ($group in $sourceUserGroups) {
    Add-ADGroupMember -Identity $group -Members $newUserName
}

Write-Host "New user $newUserName created with different name properties and description, and added to the same groups as $sourceUser."
