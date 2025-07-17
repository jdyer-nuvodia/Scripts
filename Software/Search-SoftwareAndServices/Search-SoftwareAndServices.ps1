# =============================================================================
# Script: Search-SoftwareAndServices.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1.2
# Additional Info: Fixed header metadata for workflow validation
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

# Color support variables and Write-ColorOutput function
$Script:UseAnsiColors = $PSVersionTable.PSVersion.Major -ge 7
$Script:Colors = if ($Script:UseAnsiColors) {
    @{
        'White'    = "`e[37m"
        'Cyan'     = "`e[36m"
        'Green'    = "`e[32m"
        'Yellow'   = "`e[33m"
        'Red'      = "`e[31m"
        'Magenta'  = "`e[35m"
        'DarkGray' = "`e[90m"
        'Reset'    = "`e[0m"
    }
} else {
    @{
        'White'    = [ConsoleColor]::White
        'Cyan'     = [ConsoleColor]::Cyan
        'Green'    = [ConsoleColor]::Green
        'Yellow'   = [ConsoleColor]::Yellow
        'Red'      = [ConsoleColor]::Red
        'Magenta'  = [ConsoleColor]::Magenta
        'DarkGray' = [ConsoleColor]::DarkGray
    }
}

function Write-ColorOutput {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$Color = "White"
    )

    if ($Script:UseAnsiColors) {
        # PowerShell 7+ with ANSI escape codes
        $colorCode = $Script:Colors[$Color]
        $resetCode = $Script:Colors.Reset
        Write-Output "${colorCode}${Message}${resetCode}"
    } else {
        # PowerShell 5.1 - Change console color, write output, then reset
        $originalColor = $Host.UI.RawUI.ForegroundColor
        try {
            if ($Script:Colors[$Color] -and $Script:Colors[$Color] -ne "") {
                $Host.UI.RawUI.ForegroundColor = $Script:Colors[$Color]
            }
            Write-Output $Message
        } finally {
            $Host.UI.RawUI.ForegroundColor = $originalColor
        }
    }
}

# Function to search installed software
function Search-InstalledSoftware {
    param (
        [string]$Keyword
    )

    Write-ColorOutput -Message "Searching for installed software matching keyword: $Keyword..." -Color Cyan

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
                        Type            = "Software"
                        Name            = $software.DisplayName
                        Version         = $software.DisplayVersion
                        Publisher       = $software.Publisher
                        InstallDate     = $software.InstallDate
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
function Search-WindowsService {
    param (
        [string]$Keyword
    )

    Write-ColorOutput -Message "Searching for services matching keyword: $Keyword..." -Color Cyan

    $serviceResults = @()

    # Get services matching the keyword with error handling
    try {
        # Capture non-terminating errors using ErrorVariable
        $services = Get-Service -ErrorAction SilentlyContinue -ErrorVariable serviceErrors |
            Where-Object {
                $_.DisplayName -like "*$Keyword*" -or
                $_.Name -like "*$Keyword*" -or
                ($null -ne $_.Description -and $_.Description -like "*$Keyword*")
            }

        # Log permission errors if verbose
        if ($serviceErrors) {
            $permissionDeniedCount = ($serviceErrors | Where-Object { $_.Exception.Message -like "*PermissionDenied*" }).Count
            if ($permissionDeniedCount -gt 0) {
                Write-ColorOutput -Message "Note: Unable to query $permissionDeniedCount service(s) due to permission restrictions." -Color Yellow
                Write-Verbose "Some services could not be accessed due to permission restrictions. This is normal behavior when not running as administrator."
            }
        }
    } catch {
        Write-ColorOutput -Message "Error retrieving services: $_" -Color Red
        return $serviceResults
    }

    foreach ($service in $services) {
        try {
            $serviceDetails = Get-CimInstance -ClassName Win32_Service -Filter "Name = '$($service.Name)'" -ErrorAction SilentlyContinue

            $serviceObj = [PSCustomObject]@{
                Type        = "Service"
                Name        = $service.DisplayName
                Status      = $service.Status
                StartType   = $service.StartType
                ServiceName = $service.Name
                Description = $serviceDetails.Description
                PathName    = $serviceDetails.PathName
                StartName   = $serviceDetails.StartName
            }
            $serviceResults += $serviceObj
        } catch {
            Write-Verbose "Error processing service $($service.Name): $_"
        }
    }

    return $serviceResults
}

# Start script execution
Write-ColorOutput -Message "Starting search for '$Keyword' in installed software and services..." -Color White

# Search installed software if specified
if ($IncludeSoftware) {
    if ($PSCmdlet.ShouldProcess("System", "Search installed software for keyword: $Keyword")) {
        $softwareResults = Search-InstalledSoftware -Keyword $Keyword
        $results += $softwareResults

        Write-ColorOutput -Message "Found $($softwareResults.Count) software item(s) matching '$Keyword'" -Color Green

        # Display software results
        foreach ($item in $softwareResults) {
            Write-ColorOutput -Message "`nSoftware: $($item.Name)" -Color Green
            Write-ColorOutput -Message "  Version: $($item.Version)" -Color White
            Write-ColorOutput -Message "  Publisher: $($item.Publisher)" -Color White
            if ($item.InstallLocation) {
                Write-ColorOutput -Message "  Install Location: $($item.InstallLocation)" -Color DarkGray
            }
            if ($item.InstallDate) {
                Write-ColorOutput -Message "  Install Date: $($item.InstallDate)" -Color DarkGray
            }
        }
    }
}

# Search services if specified
if ($IncludeServices) {
    if ($PSCmdlet.ShouldProcess("System", "Search services for keyword: $Keyword")) {
        $serviceResults = Search-WindowsService -Keyword $Keyword
        $results += $serviceResults

        Write-ColorOutput -Message "`nFound $($serviceResults.Count) service(s) matching '$Keyword'" -Color Green

        # Display service results
        foreach ($item in $serviceResults) {
            # Color based on service status
            $statusColor = switch ($item.Status) {
                "Running" { "Green" }
                "Stopped" { "Yellow" }
                default { "White" }
            }

            Write-ColorOutput -Message "`nService: $($item.Name) [$($item.ServiceName)]" -Color Cyan
            Write-ColorOutput -Message "  Status: $($item.Status)" -Color $statusColor
            Write-ColorOutput -Message "  Start Type: $($item.StartType)" -Color White
            if ($item.Description) {
                Write-ColorOutput -Message "  Description: $($item.Description)" -Color DarkGray
            }
            Write-ColorOutput -Message "  Path: $($item.PathName)" -Color DarkGray
        }
    }
}

# If no results found
if ($results.Count -eq 0) {
    Write-ColorOutput -Message "No items found matching keyword: $Keyword" -Color Yellow
}

# Export to CSV if path provided
if ($ExportPath -and $results.Count -gt 0) {
    if ($PSCmdlet.ShouldProcess("Export results", "Export search results to CSV file: $ExportPath")) {
        try {
            $results | Export-Csv -Path $ExportPath -NoTypeInformation
            Write-ColorOutput -Message "`nExported $($results.Count) results to $ExportPath" -Color Green
        } catch {
            Write-ColorOutput -Message "Error exporting results to CSV: $_" -Color Red
        }
    }
}

Write-ColorOutput -Message "`nSearch completed." -Color Cyan
