# Define variables
$downloadUrl = "https://wiztree.co.uk/download/wiztreeportable.zip"  # URL for WizTree Portable
$zipFilePath = "C:\temp\wiztreeportable.zip"  # Path to save the downloaded ZIP file
$extractPath = "C:\temp\WizTree"  # Path to extract files
$exePath = "$extractPath\WizTree64.exe"  # Path to the executable

# Create temp directory if it doesn't exist
if (-Not (Test-Path -Path "C:\temp")) {
    New-Item -ItemType Directory -Path "C:\temp"
}

# Download WizTree Portable
Invoke-WebRequest -Uri $downloadUrl -OutFile $zipFilePath

# Extract the ZIP file
Expand-Archive -Path $zipFilePath -DestinationPath $extractPath -Force

# Run WizTree as Administrator
Start-Process -FilePath $exePath -Verb RunAs
