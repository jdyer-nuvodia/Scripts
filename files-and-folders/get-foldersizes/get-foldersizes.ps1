$driveLetter = "C:\"
$outputFile = "C:\temp\output.txt"

# Create the output directory if it doesn't exist
if (-not (Test-Path -Path (Split-Path -Path $outputFile))) {
    New-Item -ItemType Directory -Path (Split-Path -Path $outputFile) | Out-Null
}

# Function to check if PowerShell is running as administrator
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

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
    
    $folders = Get-ChildItem -Path $path -Directory -Force -ErrorAction SilentlyContinue | ForEach-Object {
        $size = (Get-ChildItem $_.FullName -Recurse -File -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
        [PSCustomObject]@{
            FolderName = $_.FullName
            SizeGB = [Math]::Round($size / 1GB, 2)
            IsHidden = $_.Attributes.HasFlag([System.IO.FileAttributes]::Hidden)
        }
    } | Sort-Object SizeGB -Descending | Select-Object -First $limit

    if ($folders.Count -gt 0) {
        $folders | Format-Table FolderName, SizeGB, IsHidden -AutoSize | Out-File -Append -FilePath $outputFile
        $folders | Format-Table FolderName, SizeGB, IsHidden -AutoSize
    } else {
        Show-Progress "No folders found in the path: $path"
    }
    
    return $folders
}

# Function to get the largest file in a given path
function Get-LargestFile {
    param (
        [string]$path
    )
    Show-Progress "Analyzing files in path: $path"
    
    $largestFile = Get-ChildItem -Path $path -File -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
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

# Check if the script is running with administrator privileges
if (-not (Test-Administrator)) {
    Show-Progress "Restarting PowerShell as administrator..."
    Start-Process powershell -ArgumentList "-File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Initial path
$currentPath = $driveLetter
$deepestFolder = $null
$visitedPaths = @()

# Start of the script execution
Show-Progress "Script execution started at $(Get-Date)"

# Drill down through the largest subfolder until there are no more folders
while ($true) {
    if ($visitedPaths -contains $currentPath) {
        Show-Progress "Loop detected! Stopping further analysis."
        break
    }

    $visitedPaths += $currentPath

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

if ($deepestFolder -ne $null -and $deepestFolder.FolderName -ne "") {
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

# Display the results in table format
$results = @()

if ($deepestFolder -ne $null -and $deepestFolder.FolderName -ne "") {
    $results += $deepestFolder
    if ($largestFile -ne $null) {
        $results += $largestFile
    }
}

# Output the results in table format to the log file and console
if ($results.Count -gt 0) {
    $results | Format-Table FolderName, SizeGB, IsHidden -AutoSize | Out-File -FilePath $outputFile -Append
    $results | Format-Table FolderName, SizeGB, IsHidden -AutoSize
} else {
    Show-Progress "No results to display."
}

# Output the location of the log file
Write-Host "The log can be found at $outputFile."

# Pause at the end to prevent the window from closing
Write-Host "Press any key to exit..."
[System.Console]::ReadKey() | Out-Null