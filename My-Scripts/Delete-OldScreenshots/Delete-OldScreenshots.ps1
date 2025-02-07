# Set the folder path
$folderPath = "C:\Users\jdyer\OneDrive - Nuvodia\Pictures\Screenshots"

# Set the number of days old for files to be deleted
$daysOld = 30

# Get the current date
$currentDate = Get-Date

# Calculate the cutoff date
$cutoffDate = $currentDate.AddDays(-$daysOld)

# Get all files in the folder older than the cutoff date
$oldFiles = Get-ChildItem -Path $folderPath -File | Where-Object { $_.LastWriteTime -lt $cutoffDate }

# Delete the old files silently
foreach ($file in $oldFiles) {
    Remove-Item $file.FullName -Force -ErrorAction SilentlyContinue
}

# Exit silently
exit