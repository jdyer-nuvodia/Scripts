$user = "DOMAIN\username"
$rootFolder = "D:\"

Get-ChildItem -Directory -Path $rootFolder -Recurse -Force | ForEach-Object {
    $folder = $_.FullName
    $acl = Get-Acl $folder
    $userAccess = $acl.Access | Where-Object { $_.IdentityReference -eq $user }
    
    if ($userAccess) {
        [PSCustomObject]@{
            Folder = $folder
            User = $user
            Permissions = $userAccess.FileSystemRights
            IsInherited = $userAccess.IsInherited
        }
    }
} | Format-Table -AutoSize
