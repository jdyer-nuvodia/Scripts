# Define the path where the redistributable installers will be saved
$downloadPath = "$env:TEMP\Redistributables"

# Create the download directory if it doesn't exist
if (!(Test-Path -Path $downloadPath)) {
    New-Item -ItemType Directory -Path $downloadPath
}

# URLs for the latest redistributables
$urls = @(
    "https://aka.ms/vs/17/release/vc_redist.x86.exe",
    "https://aka.ms/vs/17/release/vc_redist.x64.exe"
)

# Filenames for the redistributables
$filenames = @(
    "vc_redist.x86.exe",
    "vc_redist.x64.exe"
)

# Download the redistributables
for ($i = 0; $i -lt $urls.Count; $i++) {
    Write-Host "Downloading $($filenames[$i])..."
    Invoke-WebRequest -Uri $urls[$i] -OutFile "$downloadPath\$($filenames[$i])"
}

# Install the redistributables
foreach ($filename in $filenames) {
    Write-Host "Installing $filename..."
    Start-Process -FilePath "$downloadPath\$filename" -ArgumentList "/install /passive /norestart" -Wait
}

Write-Host "Installation complete."
