# get-foldersizes.ps1

param (
    [string]$Path = "C:\",
    [int]$MaxDepth = 10
)

# Set global error action preference
$ErrorActionPreference = 'SilentlyContinue'

Write-Host "Analyzing folders in: $Path"

function Get-FolderSizes {
    param (
        [string]$FolderPath,
        [int]$CurrentDepth = 0
    )

    if ($CurrentDepth -ge $MaxDepth) {
        return $null
    }

    $folders = Get-ChildItem -Path $FolderPath -Directory
    $folderSizes = @()
    $totalItems = ($folders | Measure-Object).Count
    Write-Host "Found $totalItems folders to process..."
    $processedCount = 0
    
    foreach ($folder in $folders) {
        try {
            # Get immediate files in the current directory (non-recursive)
            $currentFiles = Get-ChildItem -Path $folder.FullName -File
            $largestCurrentFile = $currentFiles | Sort-Object -Property Length -Descending | Select-Object -First 1

            # Get all files recursively for total size calculation
            $allFiles = @([System.IO.Directory]::EnumerateFiles($folder.FullName, '*', [System.IO.SearchOption]::AllDirectories))
            $subfolders = @([System.IO.Directory]::EnumerateDirectories($folder.FullName, '*', [System.IO.SearchOption]::AllDirectories))
            $folderSize = ($allFiles | ForEach-Object { (Get-Item $_).Length } | Measure-Object -Sum).Sum
            
            $folderSizes += [PSCustomObject]@{
                Folder = $folder.FullName
                SizeGB = [math]::round($folderSize / 1GB, 2)
                TotalSubfolders = ($subfolders | Measure-Object).Count
                TotalFiles = ($allFiles | Measure-Object).Count
                LargestFile = if ($largestCurrentFile) {
                    [PSCustomObject]@{
                        Name = $largestCurrentFile.Name
                        Path = $largestCurrentFile.FullName
                        SizeGB = [math]::round($largestCurrentFile.Length / 1GB, 2)
                        SizeMB = [math]::round($largestCurrentFile.Length / 1MB, 2)
                    }
                } else { $null }
            }
            $processedCount++
            Write-Host "`rProcessed $processedCount of $totalItems folders..." -NoNewline
        } catch {
            Write-Warning "Access to the path '$($folder.FullName)' is denied."
        }
    }
    Write-Host "`nCompleted processing $processedCount folders."
    return $folderSizes
}

function Write-TableLine {
    param([int]$Length = 150)
    Write-Host ("-" * $Length)
}

$currentPath = $Path

while ($true) {
    $folderSizes = Get-FolderSizes -FolderPath $currentPath

    if ($null -eq $folderSizes -or $folderSizes.Count -eq 0) {
        $currentFiles = Get-ChildItem -Path $currentPath -File
        $largestFile = $currentFiles | Sort-Object -Property Length -Descending | Select-Object -First 1
        if ($null -ne $largestFile) {
            Write-Output "Largest file in current directory: $($largestFile.Name), Size: $([math]::round($largestFile.Length / 1GB, 2)) GB ($([math]::round($largestFile.Length / 1MB, 2)) MB)"
        } else {
            Write-Output "No files found in $currentPath"
        }
        break
    }

    # Display the top 3 largest folders in a table format
    Write-Host "`nTop 3 Largest Folders in: $currentPath`n"
    Write-TableLine
    $format = "{0,-50} | {1,10} | {2,15} | {3,12} | {4,-50}"
    Write-Host ($format -f "Folder Path", "Size (GB)", "Subfolders", "Files", "Largest File (in this directory)")
    Write-TableLine

    $topFolders = $folderSizes | Sort-Object -Property SizeGB -Descending | Select-Object -First 3
    foreach ($folder in $topFolders) {
        $largestFileInfo = if ($folder.LargestFile) {
            if ($folder.LargestFile.SizeGB -ge 1) {
                "$($folder.LargestFile.Name) ($($folder.LargestFile.SizeGB) GB)"
            } else {
                "$($folder.LargestFile.Name) ($($folder.LargestFile.SizeMB) MB)"
            }
        } else {
            "No files"
        }
        
        Write-Host ($format -f 
            ($folder.Folder.Length -gt 47 ? "..." + $folder.Folder.Substring($folder.Folder.Length - 44) : $folder.Folder),
            $folder.SizeGB,
            $folder.TotalSubfolders,
            $folder.TotalFiles,
            ($largestFileInfo.Length -gt 47 ? "..." + $largestFileInfo.Substring($largestFileInfo.Length - 44) : $largestFileInfo)
        )
    }
    Write-TableLine

    # Descend into the largest folder
    $largestFolder = $topFolders | Select-Object -First 1
    Write-Host "`nDescending into: $($largestFolder.Folder)`n"
    $currentPath = $largestFolder.Folder
}