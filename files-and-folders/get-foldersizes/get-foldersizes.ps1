# get-foldersizes.ps1

param (
    [string]$Path = "C:\",
    [int]$MaxDepth = 10
)

Write-Host "Analyzing folders in: $Path"

function Get-FolderSizes {
    [CmdletBinding()]
    param (
        [string]$FolderPath,
        [int]$CurrentDepth = 0
    )

    if ($CurrentDepth -ge $MaxDepth) {
        return $null
    }

    $folders = Get-ChildItem -Path $FolderPath -Directory -ErrorAction SilentlyContinue | Out-Null
    $results = New-Object System.Collections.Generic.List[PSObject]

    foreach ($folder in $folders) {
        try {
            $null = $files = @([System.IO.Directory]::EnumerateFiles($folder.FullName, '*', [System.IO.SearchOption]::AllDirectories))
            $null = $subfolders = @([System.IO.Directory]::EnumerateDirectories($folder.FullName, '*', [System.IO.SearchOption]::AllDirectories))
            $folderSize = 0
            foreach ($file in $files) {
                $folderSize += (Get-Item $file -ErrorAction SilentlyContinue).Length
            }
            $null = $results.Add([PSCustomObject]@{
                Folder = $folder.FullName
                SizeGB = [math]::round($folderSize / 1GB, 2)
                TotalSubfolders = $subfolders.Count
                TotalFiles = $files.Count
            })
        } catch {
            Write-Warning "Access to the path '$($folder.FullName)' is denied."
        }
    }

    Write-Output $results
}

function Get-LargestFile {
    [CmdletBinding()]
    param (
        [string]$FolderPath
    )

    try {
        $largestFile = Get-ChildItem -Path $FolderPath -Recurse -File -ErrorAction SilentlyContinue | 
            Sort-Object -Property Length -Descending | 
            Select-Object -First 1
        return $largestFile
    } catch {
        Write-Warning "Access to the path '$FolderPath' is denied."
        return $null
    }
}

$currentPath = $Path

while ($true) {
    $folderSizes = Get-FolderSizes -FolderPath $currentPath | Out-Null

    if ($null -eq $folderSizes -or $folderSizes.Count -eq 0) {
        $largestFile = Get-LargestFile -FolderPath $currentPath
        if ($null -ne $largestFile) {
            Write-Host "Largest file: $($largestFile.FullName), Size: $([math]::round($largestFile.Length / 1GB, 2)) GB"
        } else {
            Write-Host "No files found in $currentPath"
        }
        break
    }

    # Display the top 3 largest folders
    $topFolders = $folderSizes | Sort-Object -Property SizeGB -Descending | Select-Object -First 3
    foreach ($folder in $topFolders) {
        Write-Host "Folder: $($folder.Folder), Size: $($folder.SizeGB) GB, Total Subfolders: $($folder.TotalSubfolders), Total Files: $($folder.TotalFiles)"
    }

    # Descend into the largest folder
    $largestFolder = $topFolders | Select-Object -First 1
    Write-Host "Descending into largest folder: $($largestFolder.Folder), Size: $($largestFolder.SizeGB) GB, Total Subfolders: $($largestFolder.TotalSubfolders), Total Files: $($largestFolder.TotalFiles)"
    $currentPath = $largestFolder.Folder
}