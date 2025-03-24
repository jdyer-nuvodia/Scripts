# =============================================================================
# Script: Get-InstalledSoftware.ps1
# Created: 2024-01-18 15:30:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-02-27 15:45:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0
# Additional Info: Initial script creation for software inventory
# =============================================================================

<#
.SYNOPSIS
    Retrieves installed software information from Windows registry and exports to CSV.
.DESCRIPTION
    This script performs the following actions:
    - Queries multiple registry paths for installed software information
    - Retrieves DisplayName and DisplayVersion for each installed application
    - Sorts the results alphabetically by DisplayName
    - Exports the results to a CSV file named with the computer's FQDN
    - Displays the results in the console
    
    Dependencies:
    - PowerShell 5.1 or higher
    - Write access to C:\Temp directory
    - Registry read access
.EXAMPLE
    .\Get-InstalledSoftware.ps1
    Retrieves all installed software and exports to C:\Temp\InstalledSoftware_<FQDN>.csv
.NOTES
    Security Level: Low
    Required Permissions: Registry read access, filesystem write access
    Validation Requirements: Verify CSV output contains expected software entries
#>

# Define paths for installed software
$StartPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
)

# Create an array to hold the software objects
$softwareList = @()

# Loop through each path and retrieve software details
foreach ($StartPath in $StartPaths) {
    $installedSoftware = Get-ItemProperty -Path $StartPath\*
    
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
