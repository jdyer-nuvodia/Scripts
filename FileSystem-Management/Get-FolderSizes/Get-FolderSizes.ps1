# =============================================================================
# Script: Get-FolderSizes.ps1
# Created: 2025-02-05 00:55:03 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-04 17:37:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.5.8
# Additional Info: Suppressed return value output in console
# =============================================================================

# Requires -Version 5.1

<#
.SYNOPSIS
    Ultra-fast directory scanner that analyzes folder sizes and identifies largest files.

.DESCRIPTION
    This script performs a high-performance recursive directory scan to identify the largest
    folders and files in a given directory path. It uses multi-threading when available
    and optimized .NET methods for maximum performance, even when scanning system directories.
    
    All dependencies are installed automatically without user interaction.

    Features:
    - Multi-threaded scanning when ThreadJob module is available
    - Silent installation of required modules without user interaction
    - Fallback to single-threaded mode when ThreadJob is not available
    - Handles access-denied errors gracefully
    - Identifies largest files in each directory
    - Creates detailed log file of the scan
    - Continues with limited functionality if admin rights unavailable
    - Supports custom depth limitation
    - Properly handles symbolic links and junction points

    Dependencies:
    - Windows PowerShell 5.1 or later
    - ThreadJob module (optional - will be installed automatically if not present)
    - Administrative privileges recommended but not required
    - Minimum 4GB RAM recommended

    Performance Impact:
    - CPU: Medium to High during scan
    - Memory: Medium (4GB+ recommended)
    - Disk I/O: Low to Medium
    - Network: Low (unless scanning network paths)

.PARAMETER Path
    The root directory path to start scanning from. Defaults to "C:\"

.PARAMETER MaxDepth
    Maximum depth of recursion for the directory scan. Defaults to 10 levels deep.

.PARAMETER Top
    Number of largest folders to display at each level. Defaults to 3. Range: 1-50.

.EXAMPLE
    .\Get-FolderSizes.ps1
    Scans the C:\ drive with default settings

.EXAMPLE
    .\Get-FolderSizes.ps1 -Path "D:\Users" -MaxDepth 5
    Scans the D:\Users directory with a maximum depth of 5 levels

.EXAMPLE
    .\Get-FolderSizes.ps1 -Path "\\server\share"
    Scans a network share starting from the root

.EXAMPLE
    .\Get-FolderSizes.ps1 -Top 10
    Scans the C:\ drive and shows the 10 largest folders at each level

.NOTES
    Security Level: Medium
    Required Permissions: 
    - Administrative access (recommended but not required)
    - Read access to scanned directories
    - Write access to C:\temp for logging
    
    Validation Requirements:
    - Check available memory (4GB+)
    - Validate write access to log directory
    - Test ThreadJob module availability

    Author:  jdyer-nuvodia
    Created: 2025-02-05 00:55:03 UTC
    Updated: 2025-06-04 17:25:00 UTC

    Requirements:
    - Windows PowerShell 5.1 or later
    - Administrative privileges recommended
    - Minimum 4GB RAM recommended for large directory structures

    Version History:
    1.0.0 - Initial release
    1.0.1 - Fixed compatibility issues with older PowerShell versions
    1.0.2 - Added ThreadJob module handling and fallback mechanism
    1.0.8 - Fixed handling of special characters in ThreadJobs processing
    1.1.0 - Modified for silent non-interactive operation with automatic dependency installation
    1.2.0 - Updated output formatting to display results in tabular format with progress indicators
    1.4.0 - Modified to only descend into the largest folder at each directory level
    1.5.0 - Added proper support for symbolic links and junction points
    1.5.1 - Fixed 'findstr' command not found errors by using PowerShell native commands
    1.5.2 - Added special handling for OneDrive reparse points
    1.5.3 - Fixed redundant completion messages in recursive processing
    1.5.4 - Eliminated redundant completion messages in recursive processing
    1.5.5 - Completely redesigned recursive processing to prevent redundant messages
    1.5.6 - Fixed Script Analyzer warnings for unused variables
    1.5.7 - Fixed recursive processing of completion messages with completion state tracking
    1.5.8 - Suppressed return value output in console
#>

param (
    [string]$Path = "C:\",
    [int]$MaxDepth = 10,
    [ValidateRange(1, 50)]
    [int]$Top = 3
)

#region Helper Functions

# New function to detect symbolic links and junction points
function Get-PathType {
    param (
        [string]$Path
    )
    
    try {
        # Special handling for OneDrive paths
        if ($Path -match "OneDrive -") {
            $dirInfo = New-Object System.IO.DirectoryInfo $Path
            
            if ($dirInfo.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
                # This is an OneDrive reparse point - special handling
                return @{
                    Type = "OneDriveFolder"
                    Target = "Cloud Storage"
                    IsReparsePoint = $true
                    IsOneDrive = $true
                }
            }
        }
        
        $dirInfo = New-Object System.IO.DirectoryInfo $Path
        if ($dirInfo.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
            # This is a reparse point (symbolic link, junction, etc.)
            $target = $null
            $type = "ReparsePoint"
            
            # Method 1: Try fsutil for most accurate results
            try {
                $fsutil = & fsutil reparsepoint query "$Path" 2>&1
                
                if ($fsutil -match "Symbolic Link") {
                    $type = "SymbolicLink"
                    # Improved parsing logic for symbolic links
                    $printNameLine = $fsutil | Where-Object { $_ -match "Print Name:" }
                    if ($printNameLine) {
                        $target = ($printNameLine -replace "^.*?Print Name:\s*", "").Trim()
                    }
                }
                elseif ($fsutil -match "Mount Point") {
                    $type = "MountPoint" 
                    $printNameLine = $fsutil | Where-Object { $_ -match "Print Name:" }
                    if ($printNameLine) {
                        $target = ($printNameLine -replace "^.*?Print Name:\s*", "").Trim()
                    }
                }
                elseif ($fsutil -match "Junction") {
                    $type = "Junction"
                    $printNameLine = $fsutil | Where-Object { $_ -match "Print Name:" }
                    if ($printNameLine) {
                        $target = ($printNameLine -replace "^.*?Print Name:\s*", "").Trim()
                    }
                }
                # Check for OneDrive specific patterns in fsutil output
                elseif ($fsutil -match "OneDrive" -or $Path -match "OneDrive -") {
                    $type = "OneDriveFolder"
                    $target = "Cloud Storage"
                }
            }
            catch {
                Write-Verbose "fsutil method failed: $($_.Exception.Message)"
                # If path contains OneDrive, treat as OneDrive folder
                if ($Path -match "OneDrive -") {
                    $type = "OneDriveFolder"
                    $target = "Cloud Storage"
                }
            }
            
            # Method 2: Try .NET method if fsutil didn't work or target is empty
            if ([string]::IsNullOrEmpty($target)) {
                try {
                    # For Windows 10/Server 2016+
                    if ($PSVersionTable.PSVersion.Major -ge 5) {
                        # Use reflection to access the Target property if available
                        $targetProperty = [System.IO.DirectoryInfo].GetProperty("Target", [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::Public)
                        
                        if ($null -ne $targetProperty) {
                            $target = $targetProperty.GetValue($dirInfo)
                            if ($target -is [array] -and $target.Length -gt 0) {
                                $target = $target[0]  # Take first element if array
                            }
                        }
                    }
                }
                catch {
                    Write-Verbose ".NET target method failed: $($_.Exception.Message)"
                    # If path contains OneDrive, treat as OneDrive folder
                    if ($Path -match "OneDrive -") {
                        $type = "OneDriveFolder"
                        $target = "Cloud Storage"
                    }
                }
            }
            
            # Method 3: Use PowerShell native commands instead of findstr
            if ([string]::IsNullOrEmpty($target)) {
                try {
                    # Use Get-Item with -Force parameter to get link information
                    $item = Get-Item -Path $Path -Force -ErrorAction Stop
                    
                    # Check for LinkType property (PowerShell 5.1+)
                    if ($item.PSObject.Properties.Name -contains "LinkType") {
                        if ($item.LinkType -eq "Junction") {
                            $type = "Junction"
                            if ($item.PSObject.Properties.Name -contains "Target") {
                                $target = $item.Target
                                if ($target -is [array] -and $target.Length -gt 0) {
                                    $target = $target[0]
                                }
                            }
                        }
                        elseif ($item.LinkType -eq "SymbolicLink") {
                            $type = "SymbolicLink"
                            if ($item.PSObject.Properties.Name -contains "Target") {
                                $target = $item.Target
                                if ($target -is [array] -and $target.Length -gt 0) {
                                    $target = $target[0]
                                }
                            }
                        }
                    }
                }
                catch {
                    Write-Verbose "PowerShell Get-Item method failed: $($_.Exception.Message)"
                }
            }
            
            # Final check - if we still have an Unknown Target and path has OneDrive, mark as OneDrive
            if (([string]::IsNullOrEmpty($target) -or $target -eq "Unknown Target") -and $Path -match "OneDrive -") {
                $type = "OneDriveFolder"
                $target = "Cloud Storage"
            }
            
            # Return results with either found target or "Unknown Target"
            return @{
                Type = $type
                Target = if ([string]::IsNullOrEmpty($target)) { "Unknown Target" } else { $target }
                IsReparsePoint = $true
                IsOneDrive = ($type -eq "OneDriveFolder")
            }
        }
        else {
            # Regular directory
            return @{
                Type = "Directory"
                Target = $null
                IsReparsePoint = $false
                IsOneDrive = $false
            }
        }
    }
    catch {
        Write-Warning "Error determining path type for '$Path': $($_.Exception.Message)"
        # Check if it might be an OneDrive path
        if ($Path -match "OneDrive -") {
            return @{
                Type = "OneDriveFolder"
                Target = "Cloud Storage"
                IsReparsePoint = $true
                IsOneDrive = $true
            }
        }
        
        return @{
            Type = "Unknown"
            Target = $null
            IsReparsePoint = $false
            IsOneDrive = $false
        }
    }
}

function Initialize-NuGetProvider {
    try {
        $nugetProvider = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
        $minimumVersion = [Version]"2.8.5.201"
        
        if (-not $nugetProvider -or $nugetProvider.Version -lt $minimumVersion) {tinue
            Write-Host "Installing NuGet provider..." -ForegroundColor Cyan
            
            # Check internet connectivity firstder.Version -lt $minimumVersion) {
            try {-Host "Installing NuGet provider..." -ForegroundColor Cyan
                $testConnection = Test-NetConnection -ComputerName "www.powershellgallery.com" -Port 443 -InformationLevel Quiet -ErrorAction Stop
                if (-not $testConnection) {irst
                    Write-Host "ERROR: Cannot connect to PowerShell Gallery. Internet connection appears to be down." -ForegroundColor Red
                    return $false Test-NetConnection -ComputerName "www.powershellgallery.com" -Port 443 -InformationLevel Quiet -ErrorAction Stop
                }f (-not $testConnection) {
            } catch {rite-Host "ERROR: Cannot connect to PowerShell Gallery. Internet connection appears to be down." -ForegroundColor Red
                Write-Host "ERROR: Failed to check internet connectivity: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "This could prevent module installation from external repositories." -ForegroundColor Yellow
            } catch {
                Write-Host "ERROR: Failed to check internet connectivity: $($_.Exception.Message)" -ForegroundColor Red
            # Check execution policyld prevent module installation from external repositories." -ForegroundColor Yellow
            $executionPolicy = Get-ExecutionPolicy
            Write-Host "Current PowerShell execution policy: $executionPolicy" -ForegroundColor Cyan
            if ($executionPolicy -in @("Restricted", "AllSigned")) {
                Write-Host "WARNING: Current execution policy ($executionPolicy) may prevent module installation." -ForegroundColor Yellow
                Write-Host "Consider changing to RemoteSigned with: Set-ExecutionPolicy RemoteSigned -Scope Process" -ForegroundColor Yellow
            }f ($executionPolicy -in @("Restricted", "AllSigned")) {
                Write-Host "WARNING: Current execution policy ($executionPolicy) may prevent module installation." -ForegroundColor Yellow
            # Attempt to install using Find-PackageProvidered with: Set-ExecutionPolicy RemoteSigned -Scope Process" -ForegroundColor Yellow
            try {
                Write-Host "Attempting installation via Find-PackageProvider method..." -ForegroundColor Cyan
                Find-PackageProvider -Name NuGet -RequiredVersion 2.8.5.201 -ErrorAction Stop | Install-PackageProvider -Force -ErrorAction Stop
                Write-Host "NuGet provider installed successfully using Find-PackageProvider." -ForegroundColor Green
                return $trueAttempting installation via Find-PackageProvider method..." -ForegroundColor Cyan
            }   Find-PackageProvider -Name NuGet -RequiredVersion 2.8.5.201 -ErrorAction Stop | Install-PackageProvider -Force -AllowClobber -SkipPublisherCheck -ErrorAction Stop
            catch {te-Host "NuGet provider installed successfully using Find-PackageProvider." -ForegroundColor Green
                Write-Host "Find-PackageProvider method failed with error details:" -ForegroundColor Red
                Write-Host "  - Error type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
                Write-Host "  - Error message: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "Find-PackageProvider method failed with error details:" -ForegroundColor Red
                Write-Host "  - Error type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
                Write-Host "  - Error message: $($_.Exception.Message)" -ForegroundColor Red
                if ($_.Exception.InnerException) {
                    Write-Host "  - Inner error: $($_.Exception.InnerException.Message)" -ForegroundColor Red
                }
                Write-Host "  - Command attempted: Find-PackageProvider -Name NuGet -RequiredVersion 2.8.5.201" -ForegroundColor Yellow
                
                Write-Host "Attempting direct Install-PackageProvider method..." -ForegroundColor Yellow
                # Attempt direct install with Install-PackageProvider
                try {
                    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -AllowClobber -SkipPublisherCheck -ErrorAction Stop
                    Write-Host "NuGet provider installed successfully using Install-PackageProvider." -ForegroundColor Green
                    return $true
                }
                catch {
                    Write-Host "Install-PackageProvider method failed with error details:" -ForegroundColor Red
                    Write-Host "  - Error type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
                    Write-Host "  - Error message: $($_.Exception.Message)" -ForegroundColor Red
                    if ($_.Exception.InnerException) {
                        Write-Host "  - Inner error: $($_.Exception.InnerException.Message)" -ForegroundColor Red
                    }
                    
                    # Check for common issues
                    if ($_.Exception.Message -match "proxy") {
                        Write-Host "DIAGNOSIS: Error may be related to proxy configuration." -ForegroundColor Yellow
                        Write-Host "Try setting proxy with: [System.Net.WebRequest]::DefaultWebProxy = New-Object System.Net.WebProxy('http://proxyserver:port')" -ForegroundColor Yellow
                    }
                    elseif ($_.Exception.Message -match "trust|certificate") {
                        Write-Host "DIAGNOSIS: Error may be related to certificate/trust issues." -ForegroundColor Yellow
                        Write-Host "Try: [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12" -ForegroundColor Yellow
                    }
                    
                    Write-Host "MANUAL INSTALLATION: Please install NuGet provider manually with:" -ForegroundColor Yellow
                    Write-Host "Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -AllowClobber -SkipPublisherCheck" -ForegroundColor Yellow
                    return $false
                }
            }
        }
        return $true
    }
    catch {
        Write-Host "Failed to initialize NuGet provider: $($_.Exception.Message)" -ForegroundColor Yellow
        return $false
    }
}


to suppress prompts
function Initialize-ThreadJobModule {
    try {
        if (-not (Initialize-NuGetProvider)) {f (-not (Initialize-NuGetProvider)) {
            Write-Warning "Could not initialize NuGet provider. ThreadJob installation may fail."            Write-Warning "Could not initialize NuGet provider. ThreadJob installation may fail."
            return $false
        }

        # Check if ThreadJob module is already available
        if (Get-Module -ListAvailable -Name ThreadJob) {ListAvailable -Name ThreadJob) {
            Import-Module ThreadJob -ErrorAction Stop   Import-Module ThreadJob -ErrorAction Stop
            Write-Host "ThreadJob module already installed and imported successfully." -ForegroundColor Green            Write-Host "ThreadJob module already installed and imported successfully." -ForegroundColor Green
            return $true
        }}

        Write-Host "ThreadJob module not found. Attempting to install..." -ForegroundColor Cyan-Host "ThreadJob module not found. Attempting to install..." -ForegroundColor Cyan
        
        # Check PSGallery availability
        try {ry {
            $psGallery = Get-PSRepository -Name PSGallery -ErrorAction StopGallery = Get-PSRepository -Name PSGallery -ErrorAction Stop
            Write-Host "PSGallery repository status: $($psGallery.InstallationPolicy)" -ForegroundColor CyanForegroundColor Cyan
        }
        catch {
            Write-Host "ERROR: Cannot access PSGallery repository." -ForegroundColor Red   Write-Host "ERROR: Cannot access PSGallery repository." -ForegroundColor Red
            Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red    Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Attempting to register PSGallery..." -ForegroundColor Yellow to register PSGallery..." -ForegroundColor Yellow
        }
        
        # Set PSGallery as trusted
        if (-not (Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue)) {
            try {ry {
                Register-PSRepository -Default -InstallationPolicy Trusted -ErrorAction Stopister-PSRepository -Default -InstallationPolicy Trusted -ErrorAction Stop
                Write-Host "Successfully registered PSGallery repository." -ForegroundColor Green
            }
            catch {
                Write-Host "ERROR: Failed to register PSGallery repository:" -ForegroundColor Red
                Write-Host "  - Error message: $($_.Exception.Message)" -ForegroundColor Redrite-Host "  - Error message: $($_.Exception.Message)" -ForegroundColor Red
                if ($_.Exception.InnerException) {
                    Write-Host "  - Inner error: $($_.Exception.InnerException.Message)" -ForegroundColor Redt "  - Inner error: $($_.Exception.InnerException.Message)" -ForegroundColor Red
                }   }
                Write-Host "Please register PSGallery manually: Register-PSRepository -Default" -ForegroundColor YellowWrite-Host "Please register PSGallery manually: Register-PSRepository -Default" -ForegroundColor Yellow
                return $falseeturn $false
            }
        } else {
            try {ry {
                Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
                Write-Host "Successfully set PSGallery repository as trusted." -ForegroundColor Greenlor Green
            }
            catch {
                Write-Host "ERROR: Failed to set PSGallery as trusted:" -ForegroundColor Red   Write-Host "ERROR: Failed to set PSGallery as trusted:" -ForegroundColor Red
                Write-Host "  - Error message: $($_.Exception.Message)" -ForegroundColor Red       Write-Host "  - Error message: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "This may prevent automatic module installation." -ForegroundColor Yellow                Write-Host "This may prevent automatic module installation." -ForegroundColor Yellow
            }
        }

        # Install ThreadJob module with diagnostics
        try {
            Write-Host "Attempting to install ThreadJob module..." -ForegroundColor Cyan   Write-Host "Attempting to install ThreadJob module..." -ForegroundColor Cyan
            Install-Module -Name ThreadJob -Repository PSGallery -Scope AllUsers -Force -AllowClobber -ErrorAction Stop -Verbosetall-Module -Name ThreadJob -Repository PSGallery -Scope AllUsers -Force -AllowClobber -SkipPublisherCheck -ErrorAction Stop -Verbose
            Write-Host "ThreadJob module installed successfully." -ForegroundColor Green
        }
        catch {
            Write-Host "ERROR: Failed to install ThreadJob module:" -ForegroundColor RedWrite-Host "ERROR: Failed to install ThreadJob module:" -ForegroundColor Red
            Write-Host "  - Error type: $($_.Exception.GetType().FullName)" -ForegroundColor Red.GetType().FullName)" -ForegroundColor Red
            Write-Host "  - Error message: $($_.Exception.Message)" -ForegroundColor RedregroundColor Red
            
            # Additional diagnostics for common issues
            if ($_.Exception.Message -match "administrator|elevated") {f ($_.Exception.Message -match "administrator|elevated") {
                Write-Host "DIAGNOSIS: Installation requires administrator privileges." -ForegroundColor Yellowistrator privileges." -ForegroundColor Yellow
                Write-Host "Please restart PowerShell as Administrator and try again." -ForegroundColor Yellow
            }
            elseif ($_.Exception.Message -match "access|denied") {lseif ($_.Exception.Message -match "access|denied") {
                Write-Host "DIAGNOSIS: Access denied error. Check folder permissions." -ForegroundColor Yellow    Write-Host "DIAGNOSIS: Access denied error. Check folder permissions." -ForegroundColor Yellow
                Write-Host "Try using -Scope CurrentUser instead of -Scope AllUsers" -ForegroundColor Yellow
            }
            
            Write-Host "MANUAL INSTALLATION: Please install ThreadJob module manually with:" -ForegroundColor YellowANUAL INSTALLATION: Please install ThreadJob module manually with:" -ForegroundColor Yellow
            Write-Host "Install-Module -Name ThreadJob -Repository PSGallery -Scope CurrentUser -Force" -ForegroundColor Yellow   Write-Host "Install-Module -Name ThreadJob -Repository PSGallery -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck" -ForegroundColor Yellow
                        
            return $false
        }

        # Import the module with diagnostics
        try {
            Import-Module ThreadJob -ErrorAction Stop   Import-Module ThreadJob -ErrorAction Stop
            Write-Host "ThreadJob module imported successfully." -ForegroundColor Greente-Host "ThreadJob module imported successfully." -ForegroundColor Green
            return $true
        }
        catch {
            Write-Host "ERROR: Failed to import ThreadJob module:" -ForegroundColor RedRROR: Failed to import ThreadJob module:" -ForegroundColor Red
            Write-Host "  - Error message: $($_.Exception.Message)" -ForegroundColor Red   Write-Host "  - Error message: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Module may have installed but cannot be loaded." -ForegroundColor Yellow       Write-Host "Module may have installed but cannot be loaded." -ForegroundColor Yellow
            return $false return $false
        }
    }
    catch {
        Write-Warning "Could not install/import ThreadJob module: $($_.Exception.Message)"   Write-Warning "Could not install/import ThreadJob module: $($_.Exception.Message)"
        Write-Host "Falling back to single-threaded operation mode." -ForegroundColor Yellow       Write-Host "Falling back to single-threaded operation mode." -ForegroundColor Yellow
        return $false        return $false
    }    }
}}



function Format-SizeWithPadding {{
    param (
        [double]$Size,   [double]$Size,
        [int]$DecimalPlaces = 2,    [int]$DecimalPlaces = 2,
        [string]$Unit = "GB"t = "GB"
    )
    
    switch ($Unit) {
        "GB" { $divider = 1GB }
        "MB" { $divider = 1MB }   "MB" { $divider = 1MB }
        "KB" { $divider = 1KB }    "KB" { $divider = 1KB }
        default { $divider = 1GB }
    }   }
        
    return "{0:F$DecimalPlaces}" -f ($Size / $divider)alPlaces}" -f ($Size / $divider)
}

function Format-Path {ion Format-Path {
    param ( (
        [string]$Path
    )
    try {ry {
        $fullPath = [System.IO.Path]::GetFullPath($Path.Trim())llPath = [System.IO.Path]::GetFullPath($Path.Trim())
        return $fullPath
    }
    catch {atch {
        Write-Warning "Error formatting path '$Path': $($_.Exception.Message)"       Write-Warning "Error formatting path '$Path': $($_.Exception.Message)"
        return $Path        return $Path
    }
}

function Write-TableHeader {
    param([int]$Width = 150)
    
    Write-Host ("-" * $Width)
    Write-Host ("Folder Path".PadRight(50) + " | " +  | " + 
                "Size (GB)".PadLeft(11) + " | " + 
                "Subfolders".PadLeft(15) + " | " + PadLeft(15) + " | " + 
                "Files".PadLeft(12) + " | " +                "Files".PadLeft(12) + " | " + 
                "Largest File (in this directory)")                "Largest File (in this directory)")
    Write-Host ("-" * $Width)dth)
}

function Write-TableRow {Row {
    param(
        [string]$FolderPath,ath,
        [long]$Size,
        [int]$SubfolderCount,   [int]$SubfolderCount,
        [int]$FileCount,    [int]$FileCount,
        [object]$LargestFile
    )
    
    $sizeGB = Format-SizeWithPadding -Size $Size -DecimalPlaces 2 -Unit "GB"cimalPlaces 2 -Unit "GB"
    $largestFileInfo = if ($LargestFile) {FileInfo = if ($LargestFile) {
        $largestFileSize = Format-SizeWithPadding -Size $LargestFile.Size -DecimalPlaces 2 -Unit "MB" = Format-SizeWithPadding -Size $LargestFile.Size -DecimalPlaces 2 -Unit "MB"
        "$($LargestFile.Name) ($largestFileSize MB)"   "$($LargestFile.Name) ($largestFileSize MB)"
    } else {} else {
        "No files found"
    }
    
    Write-Host ($FolderPath.PadRight(50) + " | " + 
                $sizeGB.PadLeft(11) + " | " + 1) + " | " + 
                $SubfolderCount.ToString().PadLeft(15) + " | " +                $SubfolderCount.ToString().PadLeft(15) + " | " + 
                $FileCount.ToString().PadLeft(12) + " | " +                 $FileCount.ToString().PadLeft(12) + " | " + 
                $largestFileInfo)Info)
}

function Write-ProgressBar {essBar {
    param (
        [int]$Completed,   [int]$Completed,
        [int]$Total,    [int]$Total,
        [int]$Width = 50
    )
    
    $percentComplete = [math]::Min(100, [math]::Floor(($Completed / $Total) * 100))$percentComplete = [math]::Min(100, [math]::Floor(($Completed / $Total) * 100))
    $filledWidth = [math]::Floor($Width * ($percentComplete / 100))Width * ($percentComplete / 100))
    $bar = "[" + ("=" * $filledWidth).PadRight($Width) + "] $percentComplete% | Completed processing $Completed of $Total folders"idth).PadRight($Width) + "] $percentComplete% | Completed processing $Completed of $Total folders"
    
    Write-Host "`r$bar" -NoNewlinerite-Host "`r$bar" -NoNewline
    if ($Completed -eq $Total) {   if ($Completed -eq $Total) {
        Write-Host ""  # Add a newline when complete        Write-Host ""  # Add a newline when complete
    }
}}

#endregion#endregion

#region Setup

# Check for elevated privileges but don't prompt user - continue with limited functionality
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {if (-not $isAdmin) {
    Write-Host "Running with limited privileges. Some directories may be inaccessible." -ForegroundColor Yellowirectories may be inaccessible." -ForegroundColor Yellow
}

# PowerShell Version Check and ThreadJob Initialization
$script:isLegacyPowerShell = $PSVersionTable.PSVersion.Major -lt 5rsionTable.PSVersion.Major -lt 5
if ($script:isLegacyPowerShell) {ipt:isLegacyPowerShell) {
    Write-Warning "Running in PowerShell 4.0 compatibility mode. Some features may be limited."lity mode. Some features may be limited."
    $global:useThreadJobs = $false   $global:useThreadJobs = $false
} else {} else {
    $global:useThreadJobs = Initialize-ThreadJobModule= Initialize-ThreadJobModule
}

# Transcript Logging Setup
$transcriptPath = "C:\temp"
try {
    if (-not (Test-Path $transcriptPath)) {if (-not (Test-Path $transcriptPath)) {
        New-Item -ItemType Directory -Path $transcriptPath -Force -ErrorAction SilentlyContinue | Out-Null -Path $transcriptPath -Force -ErrorAction SilentlyContinue | Out-Null
    }
    
    if (Test-Path $transcriptPath) {-Path $transcriptPath) {
        $transcriptFile = Join-Path $transcriptPath "FolderScan_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').txt").txt"
        Start-Transcript -Path $transcriptFile -Force -ErrorAction SilentlyContinue
    } else {
        $transcriptFile = Join-Path $env:TEMP "FolderScan_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').txt"   $transcriptFile = Join-Path $env:TEMP "FolderScan_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').txt"
        Start-Transcript -Path $transcriptFile -Force -ErrorAction SilentlyContinuetart-Transcript -Path $transcriptFile -Force -ErrorAction SilentlyContinue
        Write-Warning "Could not create transcript in C:\temp, using $transcriptFile instead" in C:\temp, using $transcriptFile instead"
    }   }
} catch {} catch {
    Write-Warning "Failed to start transcript: $_"start transcript: $_"
}

# Script Header in Transcript
Write-Host "======================================================"
Write-Host "Folder Size Scanner - Execution Log"- Execution Log"
Write-Host "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"'yyyy-MM-dd HH:mm:ss')"
Write-Host "Started (UTC): $((Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss'))"et-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "User: $env:USERNAME"
Write-Host "Computer: $env:COMPUTERNAME"
Write-Host "Target Path: $Path"
Write-Host "Admin Privileges: $isAdmin"dmin Privileges: $isAdmin"
Write-Host "Threading Mode: $(if ($global:useThreadJobs) { 'Multi-threaded' } else { 'Single-threaded' })"Write-Host "Threading Mode: $(if ($global:useThreadJobs) { 'Multi-threaded' } else { 'Single-threaded' })"
Write-Host "======================================================"============================================"
Write-Host ""

# .NET Type Definition# .NET Type Definition
Remove-TypeData -TypeName "FastFileScanner" -ErrorAction SilentlyContinueScanner" -ErrorAction SilentlyContinue
Remove-TypeData -TypeName "FolderSizeHelper" -ErrorAction SilentlyContinueFolderSizeHelper" -ErrorAction SilentlyContinue

# Helper Type for Folder Processingr Folder Processing
Add-Type -TypeDefinition @"nition @"
using System;
using System.IO;
using System.Linq;
using System.Security;using System.Security;
using System.Collections.Generic;
using System.Runtime.InteropServices;sing System.Runtime.InteropServices;

public static class FolderSizeHelper
{
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]rror = true, CharSet = CharSet.Unicode)]
    static extern bool GetDiskFreeSpaceEx(string lpDirectoryName,ring lpDirectoryName,
        out ulong lpFreeBytesAvailable,out ulong lpFreeBytesAvailable,
        out ulong lpTotalNumberOfBytes,
        out ulong lpTotalNumberOfFreeBytes);   out ulong lpTotalNumberOfFreeBytes);
        
    public static long GetDirectorySize(string path)string path)
    {
        long size = 0;        long size = 0;
        var stack = new Stack<string>();tring>();
        stack.Push(path);tack.Push(path);

        while (stack.Count > 0)stack.Count > 0)
        {
            string dir = stack.Pop();
            try
            {
                foreach (string file in Directory.GetFiles(dir))ch (string file in Directory.GetFiles(dir))
                {
                    tryry
                    {
                        size += new FileInfo(file).Length;       size += new FileInfo(file).Length;
                    }                    }
                    catch (Exception) { }
                }

                foreach (string subDir in Directory.GetDirectories(dir))oreach (string subDir in Directory.GetDirectories(dir))
                {   {
                    stack.Push(subDir);
                }
            }
            catch (UnauthorizedAccessException) { }cessException) { }
            catch (SecurityException) { }   catch (SecurityException) { }
            catch (IOException) { }OException) { }
            catch (Exception) { }       catch (Exception) { }
        }        }
        return size;
    }

    public static Tuple<int, int> GetDirectoryCounts(string path)int, int> GetDirectoryCounts(string path)
    {
        int files = 0;
        int folders = 0;        int folders = 0;
        var stack = new Stack<string>();tring>();
        stack.Push(path);tack.Push(path);

        while (stack.Count > 0)stack.Count > 0)
        {
            string dir = stack.Pop();
            try
            {
                files += Directory.GetFiles(dir).Length;Length;
                var subDirs = Directory.GetDirectories(dir);.GetDirectories(dir);
                folders += subDirs.Length;olders += subDirs.Length;
                foreach (var subDir in subDirs) {   foreach (var subDir in subDirs) {
                    stack.Push(subDir);
                }
            }
            catch (UnauthorizedAccessException) { }cessException) { }
            catch (SecurityException) { }   catch (SecurityException) { }
            catch (IOException) { }
            catch (Exception) { }       catch (Exception) { }
        }        }
        return new Tuple<int, int>(files, folders);
    }

    public static FileDetails GetLargestFile(string path)c static FileDetails GetLargestFile(string path)
    {
        try
        {
            var fileInfo = new DirectoryInfo(path)ectoryInfo(path)
                .GetFiles("*.*", SearchOption.TopDirectoryOnly).GetFiles("*.*", SearchOption.TopDirectoryOnly)
                .OrderByDescending(f => f.Length)g(f => f.Length)
                .FirstOrDefault();ult();
                
            if (fileInfo == null)
                return null;   return null;
                
            return new FileDetails
            {
                Name = fileInfo.Name,  Name = fileInfo.Name,
                Path = fileInfo.FullName,       Path = fileInfo.FullName,
                Size = fileInfo.Length   Size = fileInfo.Length
            };   };
        }
        catchatch
        {   {
            return null;        return null;
        }
    }
    
    public class FileDetails
    {
        public string Name { get; set; }   public string Name { get; set; }
        public string Path { get; set; }       public string Path { get; set; }
        public long Size { get; set; }set; }
    }    }
}
"@ -ErrorAction SilentlyContinue"@ -ErrorAction SilentlyContinue

$ErrorActionPreference = 'SilentlyContinue'

Write-Host "Ultra-fast folder analysis starting at: $Path" "Ultra-fast folder analysis starting at: $Path"
Write-Host "Script started by: $env:USERNAME at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"Write-Host "Script started by: $env:USERNAME at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

#endregion#endregion

#region Folder Scanning Logicder Scanning Logic

function Get-FolderSize {
    param (
        [string]$FolderPath,FolderPath,
        [int]$CurrentDepth,   [int]$CurrentDepth,
        [int]$MaxDepth,        [int]$MaxDepth,
        [int]$Topint]$Top
    )

    try {
        if ($CurrentDepth -gt $MaxDepth) {) {
            return @{ 
                ProcessedFolders = $false;    ProcessedFolders = $false; 
                HasSubfolders = $false;       HasSubfolders = $false;
                CompletionMessageShown = $false                CompletionMessageShown = $false
            }
        }

        $folderPath = Format-Path $FolderPathFormat-Path $FolderPath
        if (-not (Test-Path -Path $folderPath -PathType Container)) {th -PathType Container)) {
            Write-Warning "Path '$FolderPath' does not exist or is not a directory."rPath' does not exist or is not a directory."
            return @{ 
                ProcessedFolders = $false;    ProcessedFolders = $false; 
                HasSubfolders = $false;       HasSubfolders = $false;
                CompletionMessageShown = $false                CompletionMessageShown = $false
            }
        }

        # Check if this path is a symbolic link, junction, or mount point, junction, or mount point
        $pathType = Get-PathType -Path $folderPathath $folderPath
        
        # Special handling for OneDrive folders
        if ($pathType.IsOneDrive) {
            Write-Host "`nDetected $($pathType.Type) at '$folderPath'" -ForegroundColor Cyan   Write-Host "`nDetected $($pathType.Type) at '$folderPath'" -ForegroundColor Cyan
            Write-Host "OneDrive folder - continuing with scan (files may be cloud-based)" -ForegroundColor Cyanbased)" -ForegroundColor Cyan
            # No need to follow the link for OneDrive folders - continue with current path
        }
        elseif ($pathType.Type -ne "Directory" -and $pathType.Type -ne "Unknown") {pe -ne "Unknown") {
            Write-Host "`nDetected $($pathType.Type) at '$folderPath'" -ForegroundColor Yellow
            
            # If it's a link, try to use the target path instead it's a link, try to use the target path instead
            if ($pathType.Target -and $pathType.Target -ne "Unknown Target" -and $pathType.Target -ne "Cloud Storage") {rget -ne "Unknown Target" -and $pathType.Target -ne "Cloud Storage") {
                Write-Host "Following link to target: $($pathType.Target)" -ForegroundColor YellowForegroundColor Yellow
                
                # Handle relative paths in targets relative paths in targets
                if (-not [System.IO.Path]::IsPathRooted($pathType.Target)) {ooted($pathType.Target)) {
                    $targetPath = Join-Path (Split-Path $folderPath -Parent) $pathType.Target   $targetPath = Join-Path (Split-Path $folderPath -Parent) $pathType.Target
                } else {} else {
                    $targetPath = $pathType.TargetTarget
                }
                
                # Check if the target existsif the target exists
                if (Test-Path -Path $targetPath -PathType Container) {
                    $folderPath = $targetPath
                } else { else {
                    Write-Warning "Target path '$targetPath' does not exist or is not accessible."       Write-Warning "Target path '$targetPath' does not exist or is not accessible."
                    # Continue with the original path           # Continue with the original path
                }                }
            }
        }

        Write-Host "`nTop $Top Largest Folders in: $folderPath" -ForegroundColor CyanderPath" -ForegroundColor Cyan
        Write-Host ""

        # Get Folder Size and Counts using .NET methods        # Get Folder Size and Counts using .NET methods
        $counts = [FolderSizeHelper]::GetDirectoryCounts($folderPath)izeHelper]::GetDirectoryCounts($folderPath)
        $folderCount = $counts.Item2

        # Get Largest File
        $largestFile = [FolderSizeHelper]::GetLargestFile($folderPath)derSizeHelper]::GetLargestFile($folderPath)

        # Display largest file information
        if ($largestFile) {
            Write-Host "`nLargest file in $folderPath :" -ForegroundColor GreenundColor Green
            Write-Host "Name: $($largestFile.Name)""
            $fileSize = if ($largestFile.Size -gt 1MB) {
                "$([Math]::Round($largestFile.Size / 1MB, 2)) MB"Math]::Round($largestFile.Size / 1MB, 2)) MB"
            } elseif ($largestFile.Size -gt 1KB) {1KB) {
                "$([Math]::Round($largestFile.Size / 1KB, 2)) KB"   "$([Math]::Round($largestFile.Size / 1KB, 2)) KB"
            } else {
                "$($largestFile.Size) bytes"       "$($largestFile.Size) bytes"
            }            }
            Write-Host "Size: $fileSize"ize"
        }

        # Get Subfolders and Process
        $subFolders = try { Get-ChildItem -Path $folderPath -Directory -ErrorAction Stop } catch { Write-Warning "Error getting subfolders in '$folderPath': $($_.Exception.Message)"; @() }ath $folderPath -Directory -ErrorAction Stop } catch { Write-Warning "Error getting subfolders in '$folderPath': $($_.Exception.Message)"; @() }

        if ($subFolders -and $subFolders.Count -gt 0) {$subFolders -and $subFolders.Count -gt 0) {
            $folderCount = $subFolders.Count
            Write-Host "`nFound $folderCount subfolders to process..." -ForegroundColor Cyan$folderCount subfolders to process..." -ForegroundColor Cyan
            
            # Calculate folder sizes for sortinging
            $sortedFolders = @()$sortedFolders = @()
            $currentIndex = 0
            $totalFolders = $subFolders.CountbFolders.Count
            
            foreach ($folder in $subFolders) {
                $currentIndex++currentIndex++
                if ($currentIndex % 10 -eq 0 -and $totalFolders -gt 50) {if ($currentIndex % 10 -eq 0 -and $totalFolders -gt 50) {
                    Write-Progress -Activity "Calculating folder sizes" -Status "$currentIndex of $totalFolders" -PercentComplete (($currentIndex / $totalFolders) * 100)culating folder sizes" -Status "$currentIndex of $totalFolders" -PercentComplete (($currentIndex / $totalFolders) * 100)
                }}
                
                $subFolderPath = $folder.FullName
                
                # Check if folder is a symbolic link or junction pointtion point
                $subPathType = Get-PathType -Path $subFolderPathath $subFolderPath
                
                # Use a different color for OneDrive folders Use a different color for OneDrive folders
                if ($subPathType.IsOneDrive) {
                    Write-Host "  - $($subFolderPath): $($subPathType.Type) - OneDrive cloud storage" -ForegroundColor Cyan
                }
                elseif ($subPathType.Type -ne "Directory" -and $subPathType.Type -ne "Unknown") {elseif ($subPathType.Type -ne "Directory" -and $subPathType.Type -ne "Unknown") {
                    Write-Host "  - $($subFolderPath): $($subPathType.Type) pointing to $($subPathType.Target)" -ForegroundColor Yellowrget)" -ForegroundColor Yellow
                }
                
                $subFolderSize = try { [FolderSizeHelper]::GetDirectorySize($subFolderPath) } catch { 0 }$subFolderSize = try { [FolderSizeHelper]::GetDirectorySize($subFolderPath) } catch { 0 }
                $subFolderCounts = try { [FolderSizeHelper]::GetDirectoryCounts($subFolderPath) } catch { New-Object -TypeName 'System.Tuple[int,int]'(0, 0) }Helper]::GetDirectoryCounts($subFolderPath) } catch { New-Object -TypeName 'System.Tuple[int,int]'(0, 0) }
                $subFolderLargestFile = try { [FolderSizeHelper]::GetLargestFile($subFolderPath) } catch { $null }ry { [FolderSizeHelper]::GetLargestFile($subFolderPath) } catch { $null }
                
                $sortedFolders += [PSCustomObject]@{
                    Path = $subFolderPath
                    Size = $subFolderSize
                    FileCount = $subFolderCounts.Item1.Item1
                    FolderCount = $subFolderCounts.Item2ts.Item2
                    LargestFile = $subFolderLargestFile   LargestFile = $subFolderLargestFile
                    PathType = $subPathType.Type       PathType = $subPathType.Type
                    Target = $subPathType.Target        Target = $subPathType.Target
                }
            }}
            
            Write-Progress -Activity "Calculating folder sizes" -Completed
            
            # Sort folders by size in descending ordering order
            $sortedFolders = $sortedFolders | Sort-Object -Property Size -Descending$sortedFolders | Sort-Object -Property Size -Descending
            
            # Always display the table header
            Write-TableHeader
            
            # Get top folders but ensure we don't exceed available folders# Get top folders but ensure we don't exceed available folders
            $topFoldersCount = [Math]::Min($Top, $sortedFolders.Count)$sortedFolders.Count)
            $topFolders = $sortedFolders | Select-Object -First $topFoldersCountect-Object -First $topFoldersCount
            
            # Display each folder in table formatsplay each folder in table format
            foreach ($folder in $topFolders) {
                Write-TableRow -FolderPath $folder.Path -Size $folder.Size -SubfolderCount $folder.FolderCount -FileCount $folder.FileCount -LargestFile $folder.LargestFile$folder.FolderCount -FileCount $folder.FileCount -LargestFile $folder.LargestFile
                
                # If it's a special path type, add an info line
                if ($folder.PathType -ne "Directory" -and $folder.PathType -ne "Unknown") {PathType -ne "Directory" -and $folder.PathType -ne "Unknown") {
                    if ($folder.PathType -eq "OneDriveFolder") {
                        Write-Host "  ^ OneDrive cloud-based storage" -ForegroundColor Cyan   Write-Host "  ^ OneDrive cloud-based storage" -ForegroundColor Cyan
                    } else {   } else {
                        Write-Host "  ^ $($folder.PathType) pointing to: $($folder.Target)" -ForegroundColor Yellow           Write-Host "  ^ $($folder.PathType) pointing to: $($folder.Target)" -ForegroundColor Yellow
                    }        }
                }
            }
            
            Write-Host ("-" * 150)
            Write-Host ""
            
            # Process only the largest subfolder if within depth limit
            $completionMessageShown = $falsepletionMessageShown = $false
            if ($CurrentDepth + 1 -le $MaxDepth -and $sortedFolders.Count -gt 0) {
                $largestFolder = $sortedFolders[0] # Get the single largest folder$largestFolder = $sortedFolders[0] # Get the single largest folder
                
                Write-Host "`nDescending into largest subfolder: $($largestFolder.Path)" -ForegroundColor Cyan
                
                # Call recursively and capture the structured return valuetructured return value
                $result = Get-FolderSize -FolderPath $largestFolder.Path -CurrentDepth ($CurrentDepth + 1) -MaxDepth $MaxDepth -Top $Top -FolderPath $largestFolder.Path -CurrentDepth ($CurrentDepth + 1) -MaxDepth $MaxDepth -Top $Top
                
                # Only show a completion message if:
                # 1. Processing happened
                # 2. The child had subfolders
                # 3. No completion message has been shown in this branch yets branch yet
                if ($result.ProcessedFolders -eq $true -and 
                    $result.HasSubfolders -eq $true -and  -and 
                    $result.CompletionMessageShown -eq $false) {ult.CompletionMessageShown -eq $false) {
                    Write-Host "`nCompleted processing the largest subfolder." -ForegroundColor GreenoregroundColor Green
                    $completionMessageShown = $true
                } else { else {
                    # Propagate the completion message state from child to parent       # Propagate the completion message state from child to parent
                    $completionMessageShown = $result.CompletionMessageShown        $completionMessageShown = $result.CompletionMessageShown
                }
            }
            
            # Return structured information about this level's processing
            return @{ 
                ProcessedFolders = $true;           # This level processed folders   ProcessedFolders = $true;           # This level processed folders
                HasSubfolders = $true;              # This level had subfoldersHasSubfolders = $true;              # This level had subfolders
                CompletionMessageShown = $completionMessageShown  # Track if any completion message was shown completion message was shown
            }
        } else {
            Write-Host "No subfolders found to process." -ForegroundColor YellowYellow
            # Return structured information - processed but had no subfoldersolders
            return @{ 
                ProcessedFolders = $true;    # We did process this folder    ProcessedFolders = $true;    # We did process this folder 
                HasSubfolders = $false;      # But it had no subfolders       HasSubfolders = $false;      # But it had no subfolders
                CompletionMessageShown = $false  # No completion message needed for leaf nodes           CompletionMessageShown = $false  # No completion message needed for leaf nodes
            } }
        }
    }
    catch {
        Write-Warning "Error processing folder '$FolderPath': $($_.Exception.Message)"sing folder '$FolderPath': $($_.Exception.Message)"
        return @{ 
            ProcessedFolders = $false;    ProcessedFolders = $false; 
            HasSubfolders = $false;       HasSubfolders = $false;
            CompletionMessageShown = $false           CompletionMessageShown = $false
        }        }
    }
}

# Start the Recursive Scane Recursive Scan
Get-FolderSize -FolderPath $Path -CurrentDepth 1 -MaxDepth $MaxDepth -Top $Top | Out-NullGet-FolderSize -FolderPath $Path -CurrentDepth 1 -MaxDepth $MaxDepth -Top $Top | Out-Null

#endregion

#region Drive Information Display
function Show-DriveInfo { {
    param (aram (
        [Parameter(Mandatory=$true)]    [Parameter(Mandatory=$true)]
        [object]$Volume
    )
    
    Write-Host "`nDrive Volume Details:" -ForegroundColor Green
    Write-Host "------------------------" -ForegroundColor Green
    Write-Host "Drive Letter: $($Volume.DriveLetter)" -ForegroundColor CyanCyan
    Write-Host "Drive Label: $($Volume.FileSystemLabel)" -ForegroundColor Cyan
    Write-Host "File System: $($Volume.FileSystem)" -ForegroundColor Cyan
    Write-Host "Drive Type: $($Volume.DriveType)" -ForegroundColor Cyan
    Write-Host "Size: $([math]::Round($Volume.Size/1GB, 2)) GB" -ForegroundColor Cyan   Write-Host "Size: $([math]::Round($Volume.Size/1GB, 2)) GB" -ForegroundColor Cyan
    Write-Host "Free Space: $([math]::Round($Volume.SizeRemaining/1GB, 2)) GB" -ForegroundColor Cyan    Write-Host "Free Space: $([math]::Round($Volume.SizeRemaining/1GB, 2)) GB" -ForegroundColor Cyan
    Write-Host "Health Status: $($Volume.HealthStatus)" -ForegroundColor Cyanrite-Host "Health Status: $($Volume.HealthStatus)" -ForegroundColor Cyan
}

try {
    # Get all available volumes with drive letters and sort them with drive letters and sort them
    $volumes = Get-Volume |     $volumes = Get-Volume | 
        Where-Object { $_.DriveLetter } | Letter } | 
        Sort-Object DriveLetter

    if ($volumes.Count -eq 0) {f ($volumes.Count -eq 0) {
        Write-Error "No drives with letters found on the system."        Write-Error "No drives with letters found on the system."
        exit
    }

    # Select the volume with lowest drive letter
    $lowestVolume = $volumes[0]
       
    Write-Host "Found lowest drive letter: $($lowestVolume.DriveLetter)" -ForegroundColor Yellowte-Host "Found lowest drive letter: $($lowestVolume.DriveLetter)" -ForegroundColor Yellow
    Show-DriveInfo -Volume $lowestVolume
}
catch {
    Write-Error "Error accessing drive information. Error: $_"    Write-Error "Error accessing drive information. Error: $_"
}
#endregionegion

# Stop Transcript
try {
    Stop-Transcript
    Write-Host "Transcript stopped. Log file: $transcriptFile"   Write-Host "Transcript stopped. Log file: $transcriptFile"
} catch {} catch {
    Write-Warning "Failed to stop transcript: $_"
}}



Write-Host "`nScript finished at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Green
Write-Host "`nScript finished at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Green
