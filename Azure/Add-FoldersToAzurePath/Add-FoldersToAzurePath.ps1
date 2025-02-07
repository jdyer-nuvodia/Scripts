# Path to clouddrive/scripts
$clouddriveScriptsPath = "$HOME/clouddrive/scripts"

# Check if clouddrive/scripts exists
if (Test-Path -Path $clouddriveScriptsPath) {
    # Get all directories under clouddrive/scripts, including subdirectories
    $directories = Get-ChildItem -Path $clouddriveScriptsPath -Directory -Recurse

    # Add the root scripts directory
    if ($env:PATH -notmatch [regex]::Escape($clouddriveScriptsPath)) {
        $env:PATH += ":$clouddriveScriptsPath"
        Write-Host "Added $clouddriveScriptsPath to PATH"
    } else {
        Write-Host "$clouddriveScriptsPath is already in PATH"
    }

    foreach ($dir in $directories) {
        $fullPath = $dir.FullName

        # Check if the directory is already in PATH
        if ($env:PATH -notmatch [regex]::Escape($fullPath)) {
            # Add the directory to PATH
            $env:PATH += ":$fullPath"
            Write-Host "Added $fullPath to PATH"
        } else {
            Write-Host "$fullPath is already in PATH"
        }
    }

    # Persist PATH update
    $updatedPath = $env:PATH -replace '^:', ''  # Remove leading colon if present
    echo "export PATH=$updatedPath" >> $HOME/.profile
    Write-Host "PATH has been updated and changes have been saved."
} else {
    Write-Error "The directory $clouddriveScriptsPath does not exist."
}
