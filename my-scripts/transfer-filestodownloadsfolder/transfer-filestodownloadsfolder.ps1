# Specify source and destination folders
$sourceFolder = "C:\Users\jdyer\OneDrive - Nuvodia\Documents\File Transfer"
$destinationFolder = "C:\Users\jdyer\Downlaods"

# Suppress all output
$ErrorActionPreference = 'Continue'
$ProgressPreference = 'Continue'

# Move all files from source to destination
Get-ChildItem -Path $sourceFolder -Recurse -File | Move-Item -Destination $destinationFolder

# Exit silently
