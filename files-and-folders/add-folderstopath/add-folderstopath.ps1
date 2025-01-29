# Directory you want to add to the PATH
$rootPath = "C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts"

# Get all subdirectories recursively
$subDirs = Get-ChildItem -Path $rootPath -Recurse -Directory

# Get the current PATH
$currentPath = [System.Environment]::GetEnvironmentVariable("Path", [System.EnvironmentVariableTarget]::Machine)
$currentPathArray = $currentPath -split ';'

# Function to add a path if it doesn't exist
function Add-UniquePathItem($path) {
    if ($currentPathArray -notcontains $path) {
        return $path
    }
    return $null
}

# Add the root directory and its subdirectories to the PATH if they don't exist
$newPaths = @()
$newPaths += Add-UniquePathItem $rootPath
foreach ($dir in $subDirs.FullName) {
    $newPaths += Add-UniquePathItem $dir
}

# Filter out null values and join the new paths
$newPaths = $newPaths | Where-Object { $_ -ne $null }

if ($newPaths.Count -gt 0) {
    # Join new paths into a single string for updating PATH
    $newPathsString = $newPaths -join ";"
    $newPath = $currentPath + ";" + $newPathsString

    # Update the PATH environment variable
    [System.Environment]::SetEnvironmentVariable("Path", $newPath, [System.EnvironmentVariableTarget]::Machine)

    # Echo the paths that were added
    Write-Host "The following directories were added to PATH:"
    foreach ($path in $newPaths) {
        Write-Host "- $path"
    }
} else {
    Write-Host "No new directories added to PATH."
}
