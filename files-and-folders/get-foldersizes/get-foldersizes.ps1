# Define the starting path
$startingPath = 'C:\'

function Get-LargestFolders {
    param (
        [string]$path
    )

    $folders = Get-ChildItem -Path $path -Directory -Force -ErrorAction SilentlyContinue
    if ($folders.Count -eq 0) {
        return
    }

    $largestFolders = $folders | ForEach-Object {
        $size = (Get-ChildItem $_.FullName -Recurse -File -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
        [PSCustomObject]@{
            FolderName = $_.FullName
            SizeGB = [Math]::Round($size / 1GB, 2)
            IsHidden = $_.Attributes.HasFlag([System.IO.FileAttributes]::Hidden)
        }
    } | Sort-Object SizeGB -Descending | Select-Object -First 3

    Write-Output $largestFolders

    $largestFolder = $largestFolders | Select-Object -First 1
    if ($largestFolder) {
        Process-LargestFolder -path $largestFolder.FolderName
    }
}

function Process-LargestFolder {
    param (
        [string]$path
    )

    Get-LargestFolders -path $path
}

# Start from the specified starting path
Get-LargestFolders -path $startingPath