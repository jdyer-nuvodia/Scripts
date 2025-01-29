$Username = "username"
$NewPassword = "Password123!"

$SecurePassword = ConvertTo-SecureString $newpassword -AsPlainText -Force
Set-LocalUser -Name $username -Password $SecurePassword