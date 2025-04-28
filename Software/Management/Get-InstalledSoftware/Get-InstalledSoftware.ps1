# =============================================================================
# Script: Get-InstalledSoftware.ps1
# Created: 2024-01-18 15:30:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-28 23:05:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1.1
# Additional Info: Modified to export CSV to script directory instead of C:\Temp
# =============================================================================

<#
.SYNOPSIS
    Retrieves installed software information from Windows registry and exports to CSV.
.DESCRIPTION
    This script performs the following actions:
    - Queries multiple registry paths for installed software information
    - Retrieves DisplayName and DisplayVersion for each installed application
    - Sorts the results alphabetically by DisplayName
    - Filters results by keyword if specified
    - Exports the results to a CSV file named with the computer's FQDN
    - Displays the results in the console
      Dependencies:
    - PowerShell 5.1 or higher
    - Write access to script directory
    - Registry read access
.PARAMETER Keyword
    Optional. Filter results to only include software with names containing this keyword.
    Filtering is case-insensitive.
.EXAMPLE
    .\Get-InstalledSoftware.ps1
    Retrieves all installed software and exports to InstalledSoftware_<FQDN>.csv in the script directory
.EXAMPLE
    .\Get-InstalledSoftware.ps1 -Keyword "Microsoft"
    Retrieves only software containing "Microsoft" in the name and exports to InstalledSoftware_<FQDN>.csv
.NOTES
    Security Level: Low
    Required Permissions: Registry read access, filesystem write access
    Validation Requirements: Verify CSV output contains expected software entries
#>

param (
    [Parameter(Mandatory=$false)]
    [string]$Keyword
)

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

# Filter by keyword if provided
if ($Keyword) {
    Write-Host "Filtering results for software containing: $Keyword" -ForegroundColor Cyan
    $filteredSoftwareList = $sortedSoftwareList | Where-Object { $_.DisplayName -like "*$Keyword*" }
    
    # Check if any results were found
    if ($filteredSoftwareList.Count -eq 0) {
        Write-Host "No software found containing the keyword: $Keyword" -ForegroundColor Yellow
        exit
    }
    
    # Use filtered list for output and export
    $outputList = $filteredSoftwareList
} else {
    # Use full list if no keyword provided
    $outputList = $sortedSoftwareList
}

# Output results to console
Write-Host "Found $($outputList.Count) software item(s)" -ForegroundColor Green
foreach ($software in $outputList) {
    Write-Host "$($software.DisplayName) - $($software.DisplayVersion)"
}

# Get the FQDN of the local computer
$fqdn = [System.Net.Dns]::GetHostEntry($env:computerName).HostName

# Get the script directory path
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path

# Export list to a CSV file with FQDN in the filename in the script directory
$outputFileName = Join-Path -Path $scriptDirectory -ChildPath "InstalledSoftware_$fqdn.csv"
$outputList | Export-Csv -Path $outputFileName -NoTypeInformation

# Provide export confirmation
Write-Host "Results exported to: $outputFileName" -ForegroundColor Green
