# =============================================================================
# Script: Get-InstalledSoftware.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 2.1.2
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
# =============================================================================

<#
.SYNOPSIS
    Retrieves installed software information from Windows registry and exports to CSV.
.DESCRIPTION
    This script performs the following actions:
    - Uses multiple discovery methods to find installed software:
      * Windows Registry (multiple paths)
      * WMI/CIM Product Information
      * AppX/Modern Apps (optional)
      * Running Processes
      * Common installation directories
      * Start Menu shortcuts (optional)
    - Performs cross-matching to avoid duplicate entries
    - Retrieves DisplayName and DisplayVersion for each installed application
    - Sorts the results alphabetically by DisplayName
    - Filters results by keyword if specified
    - Exports the results to a CSV file named with the computer's FQDN and timestamp
    - Displays the results in the console

    Dependencies:
    - PowerShell 5.1 or higher
    - Write access to script directory
    - Admin privileges recommended for full discovery
.PARAMETER Keyword
    Optional. Filter results to only include software with names containing this keyword.
    Filtering is case-insensitive.
.PARAMETER IncludeModernApps
    Optional switch. When specified, includes modern applications (AppX/UWP) in the results.
    Default behavior is to exclude these apps if the switch is not used.
.PARAMETER IncludeShortcuts
    Optional switch. When specified, includes applications discovered through Start Menu shortcuts.
    Default behavior is to exclude shortcut-discovered apps if the switch is not used.
.EXAMPLE
    .\Get-InstalledSoftware.ps1
    Retrieves installed software excluding modern apps and shortcut-discovered apps.
    Exports results to InstalledSoftware_<FQDN>_<TIMESTAMP>.csv
.EXAMPLE
    .\Get-InstalledSoftware.ps1 -IncludeModernApps -IncludeShortcuts
    Retrieves all installed software using all available discovery methods.
    Exports results to InstalledSoftware_<FQDN>_<TIMESTAMP>.csv
.EXAMPLE
    .\Get-InstalledSoftware.ps1 -Keyword "Microsoft" -IncludeModernApps
    Retrieves software containing "Microsoft" in the name, including modern apps.
.NOTES
    Security Level: Low
    Required Permissions: Registry read access, filesystem write access
    Validation Requirements: Verify CSV output contains expected software entries
#>

param (
    [Parameter(Mandatory = $false)]
    [string]$Keyword,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeModernApps,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeShortcuts
)

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
        'Reset'    = ''
    }
}

function Write-ColorOutput {
    <#
    .SYNOPSIS
    Outputs colored text in a way that's compatible with PSScriptAnalyzer requirements.

    .DESCRIPTION
    This function provides colored output while maintaining compatibility with PSScriptAnalyzer
    by using only Write-Output and standard PowerShell cmdlets.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$Color = "White"
    )

    # Always use Write-Output to satisfy PSScriptAnalyzer
    # For PowerShell 7+, include ANSI color codes in the output
    if ($Script:UseAnsiColors) {
        $colorCode = $Script:Colors[$Color]
        $resetCode = $Script:Colors.Reset
        Write-Output "${colorCode}${Message}${resetCode}"
    } else {
        # For PowerShell 5.1, just output the message
        # Color formatting will be handled by the terminal/host if supported
        Write-Output $Message
    }
}


# Function to normalize software names for comparison
function Get-NormalizedName {
    param (
        [string]$Name
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return ""
    }

    # Remove special characters, multiple spaces, and convert to lowercase
    $normalized = $Name -replace '[^\w\s]', ' ' -replace '\s+', ' '
    return $normalized.ToLower().Trim()
}

# Function to determine similarity between names
function Test-NameSimilarity {
    param (
        [string]$Name1,
        [string]$Name2
    )

    $normalized1 = Get-NormalizedName -Name $Name1
    $normalized2 = Get-NormalizedName -Name $Name2

    # Return true if names are similar enough (name1 contains name2 or vice versa)
    return ($normalized1 -and $normalized2) -and ($normalized1.Contains($normalized2) -or $normalized2.Contains($normalized1))
}

Write-ColorOutput -Message "Gathering installed software information..." -Color 'Cyan'

# Create an array to hold the software objects
$softwareList = @()

# Track unique software for deduplication
$uniqueSoftware = @{}

#region Registry-based software discovery
Write-ColorOutput -Message "Scanning registry locations..." -Color 'DarkGray'

# Define paths for installed software
$StartPaths = @(
    # Standard application registration paths
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",

    # Microsoft Store applications
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModel\StateRepository\Cache\Package",

    # Modern app specific locations
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths",

    # Click-to-Run applications
    "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",

    # UWP/Modern applications
    "HKLM:\SOFTWARE\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\PackageRepository\Packages"
)

# Loop through each path and retrieve software details
foreach ($StartPath in $StartPaths) {
    # Check if the registry path exists before attempting to get items
    if (Test-Path -Path $StartPath) {
        try {
            $installedSoftware = Get-ItemProperty -Path "$StartPath\*" -ErrorAction SilentlyContinue

            foreach ($obj in $installedSoftware) {
                if ($obj.DisplayName) {
                    $normalizedName = Get-NormalizedName -Name $obj.DisplayName

                    # Skip if we've already added this software (by normalized name)
                    if (-not $uniqueSoftware.ContainsKey($normalizedName)) {
                        # Create a custom object for each software
                        $software = [PSCustomObject]@{
                            DisplayName    = $obj.DisplayName
                            DisplayVersion = $obj.DisplayVersion
                            Source         = "Registry"
                        }

                        # Add to our tracking dictionary and software list
                        $uniqueSoftware[$normalizedName] = $true
                        $softwareList += $software
                    }
                }
            }
        } catch {
            Write-ColorOutput -Message "Error accessing path: $StartPath" -Color 'DarkGray'
            # Continue to the next path without halting execution
            continue
        }
    }
}
#endregion

#region WMI/CIM Product discovery
Write-ColorOutput -Message "Scanning WMI/CIM product information..." -Color 'DarkGray'
try {
    $cimProducts = Get-CimInstance -ClassName Win32_Product -ErrorAction SilentlyContinue

    foreach ($product in $cimProducts) {
        if ($product.Name) {
            $normalizedName = Get-NormalizedName -Name $product.Name

            # Skip if already added, otherwise add to our list
            if (-not $uniqueSoftware.ContainsKey($normalizedName)) {
                $software = [PSCustomObject]@{
                    DisplayName    = $product.Name
                    DisplayVersion = $product.Version
                    Source         = "WMI"
                }

                $uniqueSoftware[$normalizedName] = $true
                $softwareList += $software
            }
        }
    }
} catch {
    Write-ColorOutput -Message "Error accessing WMI product information: $_" -Color 'DarkGray'
}
#endregion

#region AppX Package discovery
if ($IncludeModernApps.IsPresent) {
    Write-ColorOutput -Message "Scanning modern apps (AppX packages)..." -Color 'DarkGray'
    try {
        # Get all AppX packages for the current user
        $appxPackages = Get-AppxPackage

        foreach ($package in $appxPackages) {
            $displayName = $package.Name
            if ($package.DisplayName) { $displayName = $package.DisplayName }

            if ($displayName) {
                $normalizedName = Get-NormalizedName -Name $displayName

                # Add each non-duplicate app to our list
                $isDuplicate = $false
                foreach ($key in $uniqueSoftware.Keys) {
                    if (Test-NameSimilarity -Name1 $normalizedName -Name2 $key) {
                        $isDuplicate = $true
                        break
                    }
                }

                if (-not $isDuplicate) {
                    $software = [PSCustomObject]@{
                        DisplayName    = $displayName
                        DisplayVersion = $package.Version
                        Source         = "AppX"
                    }

                    $uniqueSoftware[$normalizedName] = $true
                    $softwareList += $software
                }
            }
        }
    } catch {
        Write-ColorOutput -Message "Error accessing AppX package information: $_" -Color 'DarkGray'
    }
}
#endregion

#region Process-based discovery
Write-ColorOutput -Message "Scanning running processes for software..." -Color 'DarkGray'

try {
    # Get all running processes with file paths
    $processes = Get-Process -ErrorAction SilentlyContinue |
        Where-Object { $_.Path -ne $null -and $_.Path -ne "" } |
        Select-Object Name, Path -Unique

    foreach ($process in $processes) {
        try {
            # Skip system processes and known utilities
            if ($process.Name -match "^(svchost|conhost|cmd|powershell|explorer|notepad|calculator)$") {
                continue
            }

            if (Test-Path -Path $process.Path) {
                $fileInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($process.Path)

                # Use the product name if available, otherwise use process name
                $appName = $fileInfo.ProductName
                if ([string]::IsNullOrWhiteSpace($appName)) {
                    $appName = $process.Name
                }

                if (-not [string]::IsNullOrWhiteSpace($appName)) {
                    $normalizedName = Get-NormalizedName -Name $appName

                    # Check for duplicates
                    $isDuplicate = $false
                    foreach ($key in $uniqueSoftware.Keys) {
                        if (Test-NameSimilarity -Name1 $normalizedName -Name2 $key) {
                            $isDuplicate = $true
                            break
                        }
                    }

                    if (-not $isDuplicate) {
                        $software = [PSCustomObject]@{
                            DisplayName    = $appName
                            DisplayVersion = $fileInfo.ProductVersion
                            Source         = "Process"
                        }

                        $uniqueSoftware[$normalizedName] = $true
                        $softwareList += $software
                    }
                }
            }
        } catch {
            # Skip processes that can't be analyzed
            continue
        }
    }
} catch {
    Write-ColorOutput -Message "Error accessing process information: $_" -Color 'DarkGray'
}
#endregion

#region Common installation directories
Write-ColorOutput -Message "Scanning common installation directories..." -Color 'DarkGray'

# Define common application installation directories
$appDirectories = @(
    "$env:ProgramFiles",
    "${ env:ProgramFiles(x86)}",
    "$env:LOCALAPPDATA\Programs",
    "$env:APPDATA\Programs",
    "$env:ProgramData\Microsoft\Windows\Start Menu\Programs",
    "$env:LOCALAPPDATA\Microsoft",
    "$env:PROGRAMFILES\WindowsApps"
)

foreach ($directory in $appDirectories) {
    if (Test-Path $directory) {
        try {
            # Look for executable files (.exe) in first-level subdirectories
            $subdirectories = Get-ChildItem -Path $directory -Directory -ErrorAction SilentlyContinue

            foreach ($subdir in $subdirectories) {
                # Skip Windows and system directories
                if ($subdir.Name -match "^(Windows|Microsoft|Common|Internet Explorer|MSBuild|Reference|dotnet|WindowsPowerShell)$") {
                    continue
                }

                # Look for executable files
                $exeFiles = Get-ChildItem -Path $subdir.FullName -Filter "*.exe" -File -Recurse -Depth 2 -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -notlike "*uninstall*" -and $_.Name -notlike "*setup*" -and $_.Name -notlike "*installer*" } |
                    Select-Object -First 1

                foreach ($exeFile in $exeFiles) {
                    try {
                        $fileInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($exeFile.FullName)

                        # Use product name if available, or directory name as fallback
                        $appName = $fileInfo.ProductName
                        if ([string]::IsNullOrWhiteSpace($appName)) {
                            $appName = $subdir.Name
                        }

                        if (-not [string]::IsNullOrWhiteSpace($appName)) {
                            $normalizedName = Get-NormalizedName -Name $appName

                            # Check for duplicates
                            $isDuplicate = $false
                            foreach ($key in $uniqueSoftware.Keys) {
                                if (Test-NameSimilarity -Name1 $normalizedName -Name2 $key) {
                                    $isDuplicate = $true
                                    break
                                }
                            }

                            if (-not $isDuplicate) {
                                $software = [PSCustomObject]@{
                                    DisplayName    = $appName
                                    DisplayVersion = $fileInfo.ProductVersion
                                    Source         = "Installation Directory"
                                }

                                $uniqueSoftware[$normalizedName] = $true
                                $softwareList += $software
                            }
                        }
                    } catch {
                        # Skip files that can't be analyzed
                        continue
                    }
                }
            }
        } catch {
            Write-ColorOutput -Message "Error accessing directory: $directory" -Color 'DarkGray'
            continue
        }
    }
}
#endregion

#region Shortcut-based discovery
if ($IncludeShortcuts.IsPresent) {
    Write-ColorOutput -Message "Scanning Start Menu shortcuts..." -Color 'DarkGray'
    $startMenuPaths = @(
        "$env:ProgramData\Microsoft\Windows\Start Menu\Programs",
        "$env:APPDATA\Microsoft\Windows\Start Menu\Programs"
    )

    foreach ($startMenuPath in $startMenuPaths) {
        try {
            if (Test-Path $startMenuPath) {
                $shortcuts = Get-ChildItem -Path $startMenuPath -Filter "*.lnk" -Recurse -ErrorAction SilentlyContinue

                foreach ($shortcut in $shortcuts) {
                    try {
                        # Skip desktop.ini files
                        if ($shortcut.Name -like "desktop.ini") { continue }

                        # Get shortcut name without extension
                        $appName = [System.IO.Path]::GetFileNameWithoutExtension($shortcut.Name)

                        # Skip if this appears to be an installer/uninstaller
                        if ($appName -match "uninstall|setup|install|remove|update|readme|help|support") {
                            continue
                        }

                        $normalizedName = Get-NormalizedName -Name $appName
                        $isDuplicate = $false
                        foreach ($key in $uniqueSoftware.Keys) {
                            if (Test-NameSimilarity -Name1 $normalizedName -Name2 $key) {
                                $isDuplicate = $true
                                break
                            }
                        }

                        if (-not $isDuplicate) {
                            # Try to get version from target file if possible
                            $version = "Unknown"
                            $shell = New-Object -ComObject WScript.Shell
                            $targetPath = $shell.CreateShortcut($shortcut.FullName).TargetPath

                            if (Test-Path $targetPath) {
                                try {
                                    $fileInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($targetPath)
                                    if ($fileInfo.FileVersion) {
                                        $version = $fileInfo.FileVersion
                                    }
                                } catch {
                                    # Continue if we can't get version information
                                    Write-ColorOutput -Message "Unable to retrieve version information for $targetPath" -Color 'DarkGray'
                                }
                            }

                            $software = [PSCustomObject]@{
                                DisplayName    = $appName
                                DisplayVersion = $version
                                Source         = "Shortcut"
                            }

                            $uniqueSoftware[$normalizedName] = $true
                            $softwareList += $software
                        }
                    } catch {
                        # Skip shortcuts that can't be processed
                        continue
                    }
                }
            }
        } catch {
            Write-ColorOutput -Message "Error accessing shortcuts in $startMenuPath" -Color 'DarkGray'
        }
    }
}
#endregion

# Sort the software list alphabetically by DisplayName
$sortedSoftwareList = $softwareList | Sort-Object -Property DisplayName

# Filter by keyword if provided
if ($Keyword) {
    Write-ColorOutput -Message "Filtering results for software containing: $Keyword" -Color 'Cyan'
    $filteredSoftwareList = $sortedSoftwareList | Where-Object { $_.DisplayName -like "*$Keyword*" }

    # Check if any results were found
    if ($filteredSoftwareList.Count -eq 0) {
        Write-ColorOutput -Message "No software found containing the keyword: $Keyword" -Color 'Yellow'
        exit
    }

    # Use filtered list for output and export
    $outputList = $filteredSoftwareList
} else {
    # Use full list if no keyword provided
    $outputList = $sortedSoftwareList
}

# Get the FQDN of the local computer
$fqdn = [System.Net.Dns]::GetHostEntry($env:computerName).HostName

# Determine script location - handle $PSScriptRoot or $MyInvocation appropriately
$scriptDirectory = if ($PSScriptRoot) {
    # Use $PSScriptRoot if available (PowerShell 3.0 and later)
    $PSScriptRoot
} elseif ($MyInvocation.MyCommand.Path) {
    # Use $MyInvocation if path exists
    Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    # Fall back to current directory if neither is available
    $PWD.Path
}

# Output results to console
Write-ColorOutput -Message "`nFound $($outputList.Count) software item(s)" -Color 'Green'
foreach ($software in $outputList) {
    $sourceInfo = if ($software.Source) { " [$($software.Source)]" } else { "" }
    Write-ColorOutput -Message "$($software.DisplayName) - $($software.DisplayVersion)$sourceInfo" -Color 'White'
}

# Get current timestamp for the filename
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Export list to a CSV file with FQDN and timestamp in the filename in the script directory
$outputFileName = Join-Path -Path $scriptDirectory -ChildPath "InstalledSoftware_$($fqdn)_$($timestamp).csv"

# Create directory if it doesn't exist
$outputDirectory = Split-Path -Parent $outputFileName
if (-not (Test-Path -Path $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}

# Export to CSV (only Name and Version columns)
$outputList | Select-Object DisplayName, DisplayVersion | Export-Csv -Path $outputFileName -NoTypeInformation

# Provide export confirmation
Write-ColorOutput -Message "`nResults exported to: $outputFileName" -Color 'Green'
