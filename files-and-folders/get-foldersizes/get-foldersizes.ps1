param (
    [string]$Path = "C:\",
    [int]$MaxDepth = 2  # Limit the depth of recursion
)

Write-Host "Analyzing folders in: $Path"

function Get-FolderSizes {
    param (
        [string]$FolderPath,
        [int]$Depth
    )

    $folders = Get-ChildItem -Path $FolderPath -Directory -ErrorAction SilentlyContinue
    $folderSizes = @()

    foreach ($folder in $folders) {
        try {
            Start-Job -ScriptBlock {
                param ($folder)
                $folderSize = [long]0
                $items = Get-ChildItem -Path $folder.FullName -Recurse -File -ErrorAction SilentlyContinue
                foreach ($item in $items) {
                    $folderSize += $item.Length
                }
                [PSCustomObject]@{
                    Folder = $folder.FullName
                    SizeGB = [math]::round($folderSize / 1GB)
                }
            } -ArgumentList $folder
        } catch {
            Write-Warning "Access to the path '$($folder.FullName)' is denied."
        }
    }

    $jobs = Get-Job | Wait-Job | Receive-Job
    $folderSizes += $jobs

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
    $folderSizes = Get-FolderSizes -FolderPath $currentPath -Depth $currentDepth

    if ($folderSizes.Count -eq 0) {
        $largestFile = Get-LargestFile -FolderPath $currentPath
        if ($null -ne $largestFile) {
            Write-Output "Largest file: $($largestFile.FullName), Size: $([math]::round($largestFile.Length / 1GB)) GB"
        } else {
            Write-Output "No files found in $currentPath"
        }
        break
    }

    $largestFolder = $folderSizes | Sort-Object -Property SizeGB -Descending | Select-Object -First 1
    Write-Output "Descending into largest folder: $($largestFolder.Folder), Size: $($largestFolder.SizeGB) GB"
    $currentPath = $largestFolder.Folder
    $currentDepth++
}