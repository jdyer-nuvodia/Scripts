$csvPath = "C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\addUserListTo365Group\users.csv"
$groupName = "ConfRmCal - Author"

$users = Import-Csv -Path $csvPath

foreach ($user in $users) {
    Add-DistributionGroupMember -Identity $groupName -Member $user.UserPrincipalName -BypassSecurityGroupManagerCheck
}
