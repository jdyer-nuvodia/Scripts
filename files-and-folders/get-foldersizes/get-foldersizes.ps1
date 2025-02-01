# Global variables
$startingDirectory = 'C:\'

# Define the function to get folder sizes
function Get-FolderSizes {
    param (
        [string]$startingPath
    )

    # Display the path being analyzed
    Write-Host "`nAnalyzing folders in: $startingPath`n"

    # Get the size of directories under the starting path
    $folders = Get-ChildItem -Path $startingPath -Directory -Force | ForEach-Object {
        $size = (Get-ChildItem $_.FullName -Recurse -File -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
        [PSCustomObject]@{
            FolderName = $_.FullName
            SizeGB = [Math]::Round($size / 1GB, 2)
            IsHidden = $_.Attributes.HasFlag([System.IO.FileAttributes]::Hidden)
        }
    } | Sort-Object SizeGB -Descending | Select-Object -First 3

    # Display the results
    $folders | Format-Table -AutoSize

    return $folders
}

# Define the function to analyze the largest folder
function AnalyzeLargestFolder {
    param (
        [string]$startingPath
    )

    # Get the sizes of folders under the starting path
    $folders = Get-FolderSizes -startingPath $startingPath

    # Get the largest folder
    $largestFolder = $folders | Select-Object -First 1

    if ($largestFolder -ne $null -and (Get-ChildItem -Path $largestFolder.FolderName -Directory -Force | Measure-Object).Count -gt 0) {
        return $largestFolder.FolderName
    } else {
        Write-Host "No more subdirectories to analyze in the path: $startingPath"
        return $null
    }
}

# Call the function to analyze the largest folder
$path2 = AnalyzeLargestFolder -startingPath $startingDirectory
if ($path2) {
    $path3 = AnalyzeLargestFolder -startingPath $path2
    if ($path3) {
        AnalyzeLargestFolder -startingPath $path3
    }
}