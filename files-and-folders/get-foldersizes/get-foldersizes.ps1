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
                SizeGB = [math]::round($folderSize / 1GB)
            }
        } catch {
            Write-Warning "Access to the path '$($folder.FullName)' is denied."
        }
    }
}

# Collect the folder sizes into a variable
$folderSizes = Get-FolderSizes -FolderPath $Path

# Select the top 3 largest folders and display the results as a table
$folderSizes | Sort-Object -Property SizeGB -Descending | Select-Object -First 3 | Format-Table -Property Folder, SizeGB -AutoSize