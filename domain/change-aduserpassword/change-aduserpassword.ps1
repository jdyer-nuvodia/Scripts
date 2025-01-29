Import-Module ActiveDirectory

$username = "username"
$newPassword = ConvertTo-SecureString "Password123!" -AsPlainText -Force

try {
    Set-ADAccountPassword -Identity $username -NewPassword $newPassword -Reset
    Write-Host "Password changed successfully for user $username"
} catch {
    Write-Host "Failed to change password. Error: $($_.Exception.Message)"
}
