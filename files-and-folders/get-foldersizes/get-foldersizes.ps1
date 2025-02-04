param (
    [string]$Path = "C:\",
    [int]$MaxDepth = 2  # Limit the depth of recursion
)

Write-Host "Analyzing folders in: $Path"

function Get-FolderSizes {
    param (
        [string]$FolderPath
    )

    $folders = Get-ChildItem -Path $FolderPath -Directory -ErrorAction SilentlyContinue
    $folderSizes = @()

    foreach ($folder in $folders) {
        try {
            # Calculate folder size using Measure-Object in a more efficient manner
            $folderSize = Get-ChildItem -Path $folder.FullName -File -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum
            $folderSizes += [PSCustomObject]@{
                Folder = $folder.FullName
                SizeGB = [math]::round($folderSize.Sum / 1GB)
            }
        } catch {
            Write-Warning "Access to the path '$($folder.FullName)' is denied."
        }
    }

    return $folderSizes
}

function Get-LargestFile {
    param (
        [string]$FolderPath
    )

    try {
        $largestFile = Get-ChildItem -Path $FolderPath -Recurse -File -ErrorAction SilentlyContinue | Sort-Object -Property Length -Descending | Select-Object -First 1
        return $largestFile
    } catch {
        Write-Warning "Access to the path '$FolderPath' is denied."
        return $null
    }
}

$currentPath = $Path
$currentDepth = 0

while ($currentDepth -le $MaxDepth) {
    $folderSizes = Get-FolderSizes -FolderPath $currentPath

    if ($folderSizes.Count -eq 0) {
        $largestFile = Get-LargestFile -FolderPath $currentPath
        if ($null -ne $largestFile) {
            Write-Output "Largest file: $($largestFile.FullName), Size: $([math]::round($largestFile.Length / 1GB)) GB"
        } else {
            Write-Output "No files found in $currentPath"
        }
        break
    }

    # Display the top 3 largest folders
    $topFolders = $folderSizes | Sort-Object -Property SizeGB -Descending | Select-Object -First 3
    $topFolders | Format-Table -Property Folder, SizeGB -AutoSize

    # Descend into the largest folder
    $largestFolder = $topFolders | Select-Object -First 1
    Write-Output "Descending into largest folder: $($largestFolder.Folder), Size: $($largestFolder.SizeGB) GB"
    $currentPath = $largestFolder.Folder
    $currentDepth++
}