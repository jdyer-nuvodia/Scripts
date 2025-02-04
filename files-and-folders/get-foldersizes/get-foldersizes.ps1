# get-foldersizes.ps1

param (
    [string]$Path = "C:\",
    [int]$MaxDepth = 10
)

Write-Host "Analyzing folders in: $Path"

function Get-FolderSizes {
    param (
        [string]$FolderPath,
        [int]$CurrentDepth = 0
    )

    if ($CurrentDepth -ge $MaxDepth) {
        return
    }

    $folders = Get-ChildItem -Path $FolderPath -Directory -ErrorAction SilentlyContinue

    foreach ($folder in $folders) {
        try {
            $files = [System.IO.Directory]::EnumerateFiles($folder.FullName, '*', [System.IO.SearchOption]::AllDirectories)
            $subfolders = [System.IO.Directory]::EnumerateDirectories($folder.FullName, '*', [System.IO.SearchOption]::AllDirectories)
            $folderSize = 0
            foreach ($file in $files) {
                $folderSize += (Get-Item $file).Length
            }
            [PSCustomObject]@{
                Folder = $folder.FullName
                SizeGB = [math]::round($folderSize / 1GB, 2)
                TotalSubfolders = $subfolders.Count
                TotalFiles = $files.Count
            }
        } catch {
            Write-Warning "Access to the path '$folder' is denied."
        }
    }
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

while ($true) {
    $folderSizes = Get-FolderSizes -FolderPath $currentPath

    if ($folderSizes.Count -eq 0) {
        $largestFile = Get-LargestFile -FolderPath $currentPath
        if ($null -ne $largestFile) {
            Write-Output "Largest file: $($largestFile.FullName), Size: $([math]::round($largestFile.Length / 1GB, 2)) GB"
        } else {
            Write-Output "No files found in $currentPath"
        }
        break
    }

    # Display the top 3 largest folders
    $topFolders = $folderSizes | Sort-Object -Property SizeGB -Descending | Select-Object -First 3
    $topFolders | ForEach-Object {
        Write-Output "Folder: $($_.Folder), Size: $($_.SizeGB) GB, Total Subfolders: $($_.TotalSubfolders), Total Files: $($_.TotalFiles)"
    }

    # Descend into the largest folder
    $largestFolder = $topFolders | Select-Object -First 1
    Write-Output "Descending into largest folder: $($largestFolder.Folder), Size: $($largestFolder.SizeGB) GB, Total Subfolders: $($largestFolder.TotalSubfolders), Total Files: $($largestFolder.TotalFiles)"
    $currentPath = $largestFolder.Folder
}