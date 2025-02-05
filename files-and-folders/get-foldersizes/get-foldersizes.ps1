# get-foldersizes.ps1

param (
    [string]$Path = "C:\",
    [int]$MaxDepth = 10
)

$ErrorActionPreference = 'SilentlyContinue'

# Create a global variable to track the largest file
$script:largestFileInfo = $null

Write-Host "Analyzing folders in: $Path"

function Get-FolderSizes {
    param (
        [string]$FolderPath,
        [int]$CurrentDepth = 0
    )

    if ($CurrentDepth -ge $MaxDepth) {
        return $null
    }

    # Try to get folders with basic access first
    $folders = Get-ChildItem -Path $FolderPath -Directory -Force
    $folderSizes = @()
    $totalItems = ($folders | Measure-Object).Count
    Write-Host "Found $totalItems folders to process..."
    $processedCount = 0
    
    foreach ($folder in $folders) {
        try {
            # Use robocopy to get file sizes (more reliable with permissions)
            $robocopyOutput = robocopy $folder.FullName NULL /L /XJ /R:0 /W:0 /NP /E /BYTES /NFL /NDL /NJH /MT:64
            $folderSize = 0
            $fileCount = 0
            $largestFileSize = 0
            $largestFilePath = $null

            # Process robocopy output
            $robocopyOutput | ForEach-Object {
                if ($_ -match '^\s*(\d+)\s+(.*)$') {
                    $fileSize = [long]$matches[1]
                    $filePath = $matches[2].Trim()
                    $folderSize += $fileSize
                    $fileCount++

                    # Track largest file
                    if ($fileSize -gt $largestFileSize) {
                        $largestFileSize = $fileSize
                        $largestFilePath = Join-Path $folder.FullName $filePath
                    }
                }
            }

            # Update global largest file if needed
            if ($largestFileSize -gt 0 -and ($null -eq $script:largestFileInfo -or $largestFileSize -gt $script:largestFileInfo.Length)) {
                $script:largestFileInfo = [PSCustomObject]@{
                    FullName = $largestFilePath
                    Length = $largestFileSize
                }
            }

            # Get subfolder count using alternative method
            $subFolderCount = (Get-ChildItem -Path $folder.FullName -Directory -Force -Recurse | Measure-Object).Count

            $folderSizes += [PSCustomObject]@{
                Folder = $folder.FullName
                SizeGB = [math]::round($folderSize / 1GB, 2)
                TotalSubfolders = $subFolderCount
                TotalFiles = $fileCount
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
    param([int]$Length = 100)
    Write-Host ("-" * $Length)
}

$currentPath = $Path

while ($true) {
    $folderSizes = Get-FolderSizes -FolderPath $currentPath

    if ($null -eq $folderSizes -or $folderSizes.Count -eq 0) {
        if ($null -ne $script:largestFileInfo) {
            Write-Output "`nLargest file found:"
            Write-Output "Path: $($script:largestFileInfo.FullName)"
            Write-Output "Size: $([math]::round($script:largestFileInfo.Length / 1GB, 2)) GB"
        } else {
            Write-Output "No files found in any of the scanned directories"
        }
        break
    }

    # Display the top 3 largest folders in a table format
    Write-Host "`nTop 3 Largest Folders in: $currentPath`n"
    Write-TableLine
    $format = "{0,-50} | {1,10} | {2,15} | {3,12}"
    Write-Host ($format -f "Folder Path", "Size (GB)", "Subfolders", "Files")
    Write-TableLine

    $topFolders = $folderSizes | Sort-Object -Property SizeGB -Descending | Select-Object -First 3
    foreach ($folder in $topFolders) {
        Write-Host ($format -f 
            ($folder.Folder.Length -gt 47 ? "..." + $folder.Folder.Substring($folder.Folder.Length - 44) : $folder.Folder),
            $folder.SizeGB,
            $folder.TotalSubfolders,
            $folder.TotalFiles
        )
    }
    Write-TableLine

    # If we've found any files, display the current largest file
    if ($null -ne $script:largestFileInfo) {
        Write-Host "`nCurrent largest file:"
        Write-Host "Path: $($script:largestFileInfo.FullName)"
        Write-Host "Size: $([math]::round($script:largestFileInfo.Length / 1GB, 2)) GB`n"
    }

    # Descend into the largest folder
    $largestFolder = $topFolders | Select-Object -First 1
    Write-Host "`nDescending into: $($largestFolder.Folder)`n"
    $currentPath = $largestFolder.Folder
}