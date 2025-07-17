# =============================================================================
# Script: Install-DattoFromUSB.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.1
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
# =============================================================================

<#
.SYNOPSIS
    Automates the installation of Datto Agent software from a USB drive.

.DESCRIPTION
    This script automatically detects a connected USB drive, locates the Datto Agent
    installation files in a predefined folder structure, and performs a silent installation.

    The script handles:
    - USB drive detection
    - Verification of installation files
    - Silent installation with appropriate parameters
    - Error handling for missing drives or installation files

.EXAMPLE
    .\Install-DattoFromUSB.ps1
    Automatically detects a USB drive and installs the Datto Agent silently.
#>

# Function to find the USB drive letter
function Find-USBDrive {
    $drives = Get-Disk | Where-Object { $_.BusType -eq "USB" }
    if ($drives) {
        # Assuming the first USB drive is the one you want
        $usbDrive = $drives | Select-Object -First 1
        $partition = $usbDrive | Get-Partition | Select-Object -First 1
        if ($partition) {
            $volume = $partition | Get-Volume | Select-Object -First 1
            if ($volume) {
                return $volume.DriveLetter
            }
        }
    }
    # Return null if no USB drive is found
    return $null
}

# Define the path to the DattoAgent folder (relative to the USB drive's root)
$DattoAgentFolderRelative = "DattoAgent"

# Find the USB drive letter
$usbDriveLetter = Find-USBDrive

if ($usbDriveLetter) {
    # Construct the full path to the DattoAgent folder on the USB drive
    $DattoAgentFolder = "${ usbDriveLetter}:\$DattoAgentFolderRelative"

    # Check if the DattoAgent folder exists
    if (Test-Path $DattoAgentFolder -PathType Container) {

        # Find the executable file within the DattoAgent folder
        # This assumes only one executable exists in the folder
        $installerFile = Get-ChildItem -Path $DattoAgentFolder -Filter *.exe | Select-Object -First 1

        if ($installerFile) {
            # Construct the full path to the installer
            $installerPath = $installerFile.FullName

            # Run the installer with silent installation arguments
            # IMPORTANT: Replace "/S /qn" with the actual silent install arguments for your installer
            Start-Process -FilePath $installerPath -ArgumentList "/S /qn" -Wait -NoNewWindow
        } else {
            Write-Error "No executable file found in the DattoAgent folder: $DattoAgentFolder"
        }
    } else {
        Write-Error "DattoAgent folder not found on USB drive: $DattoAgentFolder"
    }
} else {
    Write-Error "USB drive not found."
}
