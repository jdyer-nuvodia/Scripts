param (
    [string]$Path = "C:\"
)

Write-Host "Analyzing folders in: $Path"

# Function to get folder sizes
function Get-FolderSizes {
    param (
        [string]$FolderPath
    )

    $folders = Get-ChildItem -Path $FolderPath -Directory -ErrorAction SilentlyContinue
    $folderSizes = @()

    foreach ($folder in $folders) {
        try {
            $folderSize = (Get-ChildItem -Path $folder.FullName -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
            $folderSizes += [PSCustomObject]@{
                Folder = $folder.FullName
                SizeGB = [math]::round($folderSize / 1GB)
            }
        } catch {
            Write-Warning "Access to the path '$($folder.FullName)' is denied."
        }
    }

    return $folderSizes
}

# Function to get the largest file
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

# Start the analysis loop
while ($true) {
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

    # Find the largest folder and continue the loop
    $largestFolder = $folderSizes | Sort-Object -Property SizeGB -Descending | Select-Object -First 1
    Write-Output "Descending into largest folder: $($largestFolder.Folder), Size: $($largestFolder.SizeGB) GB"
    $currentPath = $largestFolder.Folder
}