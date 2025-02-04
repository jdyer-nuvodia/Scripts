param (
    [string]$Path = "C:\"
)

Write-Host "Analyzing folders in: $Path"

function Get-FolderSizes {
    param (
        [string]$FolderPath
    )

    $folders = Get-ChildItem -Path $FolderPath -Directory -ErrorAction SilentlyContinue
    foreach ($folder in $folders) {
        try {
            $folderSize = (Get-ChildItem -Path $folder.FullName -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
            [PSCustomObject]@{
                Folder = $folder.FullName
                SizeMB = [math]::round($folderSize / 1MB, 2)
            }
        } catch {
            Write-Warning "Access to the path '$($folder.FullName)' is denied."
        }
    }
}

# Collect the folder sizes into a variable
$folderSizes = Get-FolderSizes -FolderPath $Path

# Display the results as a table
$folderSizes | Format-Table -Property Folder, SizeMB -AutoSize