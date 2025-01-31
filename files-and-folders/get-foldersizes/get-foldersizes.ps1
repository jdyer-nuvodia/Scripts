# Define the function to get folder sizes
function Get-FolderSizes {
    param (
        [string]$startingPath
    )

    # Get the size of directories under the starting path
    Get-ChildItem -Path $startingPath -Directory -Force | ForEach-Object {
        $size = (Get-ChildItem $_.FullName -Recurse -File -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
        [PSCustomObject]@{
            FolderName = $_.FullName
            SizeGB = [Math]::Round($size / 1GB, 2)
            IsHidden = $_.Attributes.HasFlag([System.IO.FileAttributes]::Hidden)
        }
    } | Sort-Object SizeGB -Descending | Select-Object -First 3
}

# Define the function to get the largest folder and analyze it
function AnalyzeLargestFolder {
    param (
        [string]$startingPath
    )

    # Get the largest folder
    $largestFolder = Get-FolderSizes -startingPath $startingPath | Select-Object -First 1

    if ($largestFolder -ne $null) {
        Write-Host "Analyzing largest folder: $($largestFolder.FolderName)"
        # Run the Get-FolderSizes function on the largest folder
        Get-FolderSizes -startingPath $largestFolder.FolderName
    } else {
        Write-Host "No folders found in the specified path."
    }
}

# Define the starting path
$startingPath = 'C:\'

# Call the function to analyze the largest folder
AnalyzeLargestFolder -startingPath $startingPath