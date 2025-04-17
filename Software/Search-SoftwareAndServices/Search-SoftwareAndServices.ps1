# =============================================================================
# Script: Search-SoftwareAndServices.ps1
# Created: 2025-04-17 19:54:45 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-17 19:54:45 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.0
# Additional Info: Initial script creation for searching software and services
# =============================================================================

<#
.SYNOPSIS
    Searches all installed software and services for a specified keyword.

.DESCRIPTION
    This script performs the following actions:
    - Searches Windows registry for installed software matching the specified keyword
    - Searches Windows services for names or descriptions matching the specified keyword
    - Displays matches with color-coded output based on result type
    - Optionally exports results to a CSV file
    
    Dependencies:
    - PowerShell 5.1 or higher
    - Registry read access
    - Service enumeration permissions
    
.PARAMETER Keyword
    The keyword to search for in software names, descriptions, and services.
    This parameter is mandatory.

.PARAMETER ExportPath
    Optional. The path where to export the CSV file with results.
    If not specified, results are only displayed in the console.

.PARAMETER IncludeServices
    Optional. Switch to include services in the search.
    Default is to search both software and services.

.PARAMETER IncludeSoftware
    Optional. Switch to include software in the search.
    Default is to search both software and services.

.PARAMETER WhatIf
    Shows what would happen if the script runs without executing any actions that would modify the system.

.EXAMPLE
    .\Search-SoftwareAndServices.ps1 -Keyword "Adobe"
    Searches for "Adobe" in both installed software and services and displays the results.

.EXAMPLE
    .\Search-SoftwareAndServices.ps1 -Keyword "SQL" -ExportPath "C:\Temp\SQLComponents.csv" -IncludeSoftware
    Searches for "SQL" only in installed software and exports the results to the specified CSV file.

.EXAMPLE
    .\Search-SoftwareAndServices.ps1 -Keyword "Print" -IncludeServices
    Searches for "Print" only in services and displays the results.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Keyword,
    
    [Parameter(Mandatory = $false)]
    [string]$ExportPath,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeServices,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeSoftware
)

# If neither switch is specified, search both by default
if (-not $IncludeServices -and -not $IncludeSoftware) {
    $IncludeServices = $true
    $IncludeSoftware = $true
}

# Create results array
$results = @()

function Write-ColorOutput {
    param (
        [string]$Message,
        [string]$ForegroundColor = "White"
    )
    
    Write-Host $Message -ForegroundColor $ForegroundColor
}

# Function to search installed software
function Search-InstalledSoftware {
    param (
        [string]$Keyword
    )
    
    Write-ColorOutput "Searching for installed software matching keyword: $Keyword..." -ForegroundColor Cyan
    
    # Define paths for installed software
    $registryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    
    $softwareResults = @()
    
    # Loop through each registry path and retrieve software details
    foreach ($path in $registryPaths) {
        if (Test-Path $path) {
            $installedSoftware = Get-ItemProperty -Path "$path\*" -ErrorAction SilentlyContinue
            
            foreach ($software in $installedSoftware) {
                if ($software.DisplayName -and ($software.DisplayName -like "*$Keyword*" -or $software.Publisher -like "*$Keyword*")) {
                    $softwareObj = [PSCustomObject]@{
                        Type = "Software"
                        Name = $software.DisplayName
                        Version = $software.DisplayVersion
                        Publisher = $software.Publisher
                        InstallDate = $software.InstallDate
                        InstallLocation = $software.InstallLocation
                        UninstallString = $software.UninstallString
                    }
                    
                    $softwareResults += $softwareObj
                }
            }
        }
    }
    
    return $softwareResults
}

# Function to search Windows services
function Search-WindowsServices {
    param (
        [string]$Keyword
    )
    
    Write-ColorOutput "Searching for services matching keyword: $Keyword..." -ForegroundColor Cyan
    
    $serviceResults = @()
    
    # Get services matching the keyword
    $services = Get-Service | Where-Object { 
        $_.DisplayName -like "*$Keyword*" -or 
        $_.Name -like "*$Keyword*" -or 
        $_.Description -like "*$Keyword*"
    }
    
    foreach ($service in $services) {
        $serviceDetails = Get-WmiObject -Class Win32_Service -Filter "Name='$($service.Name)'" -ErrorAction SilentlyContinue
        
        $serviceObj = [PSCustomObject]@{
            Type = "Service"
            Name = $service.DisplayName
            Status = $service.Status
            StartType = $service.StartType
            ServiceName = $service.Name
            Description = $serviceDetails.Description
            PathName = $serviceDetails.PathName
            StartName = $serviceDetails.StartName
        }
        
        $serviceResults += $serviceObj
    }
    
    return $serviceResults
}

# Start script execution
Write-ColorOutput "Starting search for '$Keyword' in installed software and services..." -ForegroundColor White

# Search installed software if specified
if ($IncludeSoftware) {
    if ($PSCmdlet.ShouldProcess("System", "Search installed software for keyword: $Keyword")) {
        $softwareResults = Search-InstalledSoftware -Keyword $Keyword
        $results += $softwareResults
        
        Write-ColorOutput "Found $($softwareResults.Count) software item(s) matching '$Keyword'" -ForegroundColor Green
        
        # Display software results
        foreach ($item in $softwareResults) {
            Write-ColorOutput "`nSoftware: $($item.Name)" -ForegroundColor Green
            Write-ColorOutput "  Version: $($item.Version)" -ForegroundColor White
            Write-ColorOutput "  Publisher: $($item.Publisher)" -ForegroundColor White
            if ($item.InstallLocation) {
                Write-ColorOutput "  Install Location: $($item.InstallLocation)" -ForegroundColor DarkGray
            }
            if ($item.InstallDate) {
                Write-ColorOutput "  Install Date: $($item.InstallDate)" -ForegroundColor DarkGray
            }
        }
    }
}

# Search services if specified
if ($IncludeServices) {
    if ($PSCmdlet.ShouldProcess("System", "Search services for keyword: $Keyword")) {
        $serviceResults = Search-WindowsServices -Keyword $Keyword
        $results += $serviceResults
        
        Write-ColorOutput "`nFound $($serviceResults.Count) service(s) matching '$Keyword'" -ForegroundColor Green
        
        # Display service results
        foreach ($item in $serviceResults) {
            # Color based on service status
            $statusColor = switch ($item.Status) {
                "Running" { "Green" }
                "Stopped" { "Yellow" }
                default { "White" }
            }
            
            Write-ColorOutput "`nService: $($item.Name) [$($item.ServiceName)]" -ForegroundColor Cyan
            Write-ColorOutput "  Status: $($item.Status)" -ForegroundColor $statusColor
            Write-ColorOutput "  Start Type: $($item.StartType)" -ForegroundColor White
            if ($item.Description) {
                Write-ColorOutput "  Description: $($item.Description)" -ForegroundColor DarkGray
            }
            Write-ColorOutput "  Path: $($item.PathName)" -ForegroundColor DarkGray
        }
    }
}

# If no results found
if ($results.Count -eq 0) {
    Write-ColorOutput "No items found matching keyword: $Keyword" -ForegroundColor Yellow
}

# Export to CSV if path provided
if ($ExportPath -and $results.Count -gt 0) {
    if ($PSCmdlet.ShouldProcess("Export results", "Export search results to CSV file: $ExportPath")) {
        try {
            $results | Export-Csv -Path $ExportPath -NoTypeInformation
            Write-ColorOutput "`nExported $($results.Count) results to $ExportPath" -ForegroundColor Green
        }
        catch {
            Write-ColorOutput "Error exporting results to CSV: $_" -ForegroundColor Red
        }
    }
}

Write-ColorOutput "`nSearch completed." -ForegroundColor Cyan
