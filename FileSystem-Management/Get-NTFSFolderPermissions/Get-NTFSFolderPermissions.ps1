# Define the folder path
$FolderPath = "C:\YourFolderPath"

# Output file (optional)
$OutputCsv = "C:\NTFSPermissionsReport.csv"

# Initialize an array to store results
$Results = @()

# Get all folders and subfolders recursively
$Folders = Get-ChildItem -Path $FolderPath -Recurse -Directory

# Include the root folder in the list
$Folders += Get-Item -Path $FolderPath

# Loop through each folder and retrieve NTFS permissions
foreach ($Folder in $Folders) {
    $Acl = Get-Acl -Path $Folder.FullName

    foreach ($Access in $Acl.Access) {
        # Create a custom object for each permission entry
        $Result = [PSCustomObject]@{
            FolderPath       = $Folder.FullName
            IdentityReference = $Access.IdentityReference
            FileSystemRights  = $Access.FileSystemRights
            AccessControlType = $Access.AccessControlType
            IsInherited       = $Access.IsInherited
            InheritanceFlags  = $Access.InheritanceFlags
            PropagationFlags  = $Access.PropagationFlags
        }
        # Add the result to the array
        $Results += $Result
    }
}

# Display results in a human-readable table format
$Results | Format-Table -AutoSize

# Export results to CSV (optional)
if ($OutputCsv) {
    $Results | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
    Write-Host "NTFS permissions exported to: $OutputCsv"
}
