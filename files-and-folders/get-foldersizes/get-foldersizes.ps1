$driveLetter = "C:\"
$outputFile = "output.txt"

# Function to display progress
function Show-Progress {
    param (
        [string]$message
    )
    Write-Host $message
    Write-Output $message | Out-File -Append -FilePath $outputFile
}

# Function to get the largest folders in a given path
function Get-LargestFolders {
    param (
        [string]$path,
        [int]$limit = 3
    )
    Show-Progress "Analyzing folders in path: $path"
    
    $folders = Get-ChildItem -Path $path -Directory -Force | ForEach-Object {
        $size = (Get-ChildItem $_.FullName -Recurse -File -Force | Measure-Object -Property Length -Sum).Sum
        [PSCustomObject]@{
            FolderName = $_.FullName
            SizeGB = [Math]::Round($size / 1GB, 2)
            IsHidden = $_.Attributes.HasFlag([System.IO.FileAttributes]::Hidden)
        }
    } | Sort-Object SizeGB -Descending | Select-Object -First $limit

    foreach ($folder in $folders) {
        Show-Progress "Folder: $($folder.FolderName), Size: $($folder.SizeGB) GB"
    }

    return $folders
}

# Function to get the largest file in a given path
function Get-LargestFile {
    param (
        [string]$path
    )
    Show-Progress "Analyzing files in path: $path"
    
    $largestFile = Get-ChildItem -Path $path -File -Recurse -Force | ForEach-Object {
        [PSCustomObject]@{
            FileName = $_.FullName
            SizeGB = [Math]::Round($_.Length / 1GB, 2)
            IsHidden = $_.Attributes.HasFlag([System.IO.FileAttributes]::Hidden)
        }
    } | Sort-Object SizeGB -Descending | Select-Object -First 1

    if ($largestFile -ne $null) {
        Show-Progress "Largest file found: $($largestFile.FileName) with size $($largestFile.SizeGB) GB"
    }
    return $largestFile
}

# Initial path
$currentPath = $driveLetter
$deepestFolder = $null

# Start of the script execution
Show-Progress "Script execution started at $(Get-Date)"

# Drill down through the largest subfolder until there are no more folders
while ($true) {
    try {
        $largestFolders = Get-LargestFolders -path $currentPath -limit 3
        if ($null -eq $largestFolders -or $largestFolders.Count -eq 0) {
            break
        }
        $deepestFolder = $largestFolders[0]  # Select the first (largest) subfolder
        $currentPath = $deepestFolder.FolderName
    } catch {
        Write-Error "An error occurred while processing the path: $currentPath"
        break
    }
}

if ($deepestFolder -ne $null) {
    Show-Progress "Deepest folder path: $($deepestFolder.FolderName) with size $($deepestFolder.SizeGB) GB"
    try {
        $largestFile = Get-LargestFile -path $deepestFolder.FolderName
        if ($largestFile -ne $null) {
            Show-Progress "Largest file in the deepest folder: $($largestFile.FileName) with size $($largestFile.SizeGB) GB"
        } else {
            Show-Progress "No files found in the deepest folder."
        }
    } catch {
        Write-Error "An error occurred while processing the path: $($deepestFolder.FolderName)"
    }
} else {
    Show-Progress "No folders found in the drive."
}

# End of the script execution
Show-Progress "Script execution ended at $(Get-Date)"

# Output the location of the log file
Write-Host "The log can be found at $outputFile."