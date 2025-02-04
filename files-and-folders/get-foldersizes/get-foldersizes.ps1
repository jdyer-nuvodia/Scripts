param (
    [string]$Path = "C:\",
    [int]$MaxDepth = 10  # Increased depth to allow deeper recursion
)

Write-Host "Analyzing folders in: $Path"

function Get-FolderSizes {
    param (
        [string]$FolderPath,
        [int]$CurrentDepth = 0
    )

    if ($CurrentDepth -ge $MaxDepth) {
        return @()
    }

    $folders = Get-ChildItem -Path $FolderPath -Directory -ErrorAction SilentlyContinue
    $folderSizes = @()

    $jobs = @()
    foreach ($folder in $folders) {
        $job = Start-Job -ScriptBlock {
            param ($folder)
            try {
                $files = [System.IO.Directory]::EnumerateFiles($folder, '*', [System.IO.SearchOption]::AllDirectories)
                $subfolders = [System.IO.Directory]::EnumerateDirectories($folder, '*', [System.IO.SearchOption]::AllDirectories)
                $folderSize = 0
                foreach ($file in $files) {
                    $folderSize += (Get-Item $file).Length
                }
                return [PSCustomObject]@{
                    Folder = $folder
                    SizeGB = [math]::round($folderSize / 1GB, 2)  # Rounded to 2 decimal places
                    TotalSubfolders = $subfolders.Count
                    TotalFiles = $files.Count
                }
            } catch {
                Write-Warning "Access to the path '$folder' is denied."
                return $null
            }
        } -ArgumentList $folder.FullName
        $jobs += $job
    }

    $jobs | ForEach-Object {
        $result = Receive-Job -Job $_ -Wait
        if ($result -ne $null) {
            $folderSizes += $result
        }
        Remove-Job -Job $_
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
    $topFolders | Format-Table -Property Folder, SizeGB, TotalSubfolders, TotalFiles -AutoSize

    # Descend into the largest folder
    $largestFolder = $topFolders | Select-Object -First 1
    Write-Output "Descending into largest folder: $($largestFolder.Folder), Size: $($largestFolder.SizeGB) GB, Total Subfolders: $($largestFolder.TotalSubfolders), Total Files: $($largestFolder.TotalFiles)"
    $currentPath = $largestFolder.Folder
}