# Define paths for installed software
$paths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
)

# Create an array to hold the software objects
$softwareList = @()

# Loop through each path and retrieve software details
foreach ($path in $paths) {
    $installedSoftware = Get-ItemProperty -Path $path\*
    
    foreach ($obj in $installedSoftware) {
        if ($obj.DisplayName) {
            # Create a custom object for each software
            $software = [PSCustomObject]@{
                DisplayName = $obj.DisplayName
                DisplayVersion = $obj.DisplayVersion
            }
            
            # Add the software object to the list
            $softwareList += $software
        }
    }
}

# Sort the software list alphabetically by DisplayName
$sortedSoftwareList = $softwareList | Sort-Object -Property DisplayName

# Output sorted list to console
foreach ($software in $sortedSoftwareList) {
    Write-Host "$($software.DisplayName) - $($software.DisplayVersion)"
}

# Get the FQDN of the local computer
$fqdn = [System.Net.Dns]::GetHostEntry($env:computerName).HostName

# Export sorted list to a CSV file with FQDN in the filename
$sortedSoftwareList | Export-Csv -Path "C:\Temp\InstalledSoftware_$fqdn.csv" -NoTypeInformation
