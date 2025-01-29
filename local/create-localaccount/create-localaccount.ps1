$username = "jsavikko"
$password = (ConvertTo-SecureString "Password!123" -AsPlainText -Force)
$fullname = "Jason Savikko"
$description = "Local account for Jason Savikko - #1759968"

# Create user Account

New-LocalUser -Name $username -Password $password -FullName $fullname -Description $description

# Force Password Change
net user $username /logonpasswordchg:yes

# Ensure the account is active
Enable-LocalUser -Name $username

# Add user to Users group
Add-LocalGroupMember -Group Users -Member $username