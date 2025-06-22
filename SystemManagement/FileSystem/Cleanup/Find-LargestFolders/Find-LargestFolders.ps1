# =============================================================================
# Script: Find-LargestFolders.ps1
# Created: 2025-06-21 00:50:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-06-21 03:28:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2.3
# Additional Info: Fixed table formatting in PowerShell 5.1 by replacing Format-Table with manual formatting to prevent terminal resize issues
# =============================================================================

<#
.SYNOPSIS
Efficiently finds the largest folders by recursively drilling down into only the largest subdirectories.

.DESCRIPTION
This script scans a directory and identifies the largest subdirectories and files without scanning the entire drive.
It now includes comprehensive system overhead detection including:
- NTFS file system overhead (MFT, reserved clusters)
- Volume Shadow Copy (VSS) storage overhead
- System files (pagefile.sys, hiberfil.sys, System Volume Information, Recycle Bin)
- Hidden and system files/directories using the -Force parameter

The script works by:
1. Scanning the top-level directories in the specified path (including hidden/system directories)
2. Identifying the largest folders and files (including hidden/system files)
3. Recursively drilling down into only the largest folder
4. Repeating this process until reaching the specified depth or finding no more subdirectories
5. Displaying system overhead information when analyzing drive roots

This approach is much faster than scanning entire drives and focuses on finding the largest space consumers while accounting for all system overhead.

.PARAMETER StartPath
The root directory path to begin analysis. Default is "C:\"

.PARAMETER MaxDepth
Maximum recursion depth for drilling down. Default is 15 levels.

.PARAMETER TopCount
Number of largest folders/files to display at each level. Default is 3.

.PARAMETER MinSizeGB
Minimum size in GB to display a folder/file. Default is 0.1 GB (100 MB).

.PARAMETER WhatIf
Shows what would be analyzed without performing the actual scan.

.EXAMPLE
.\Find-LargestFolders.ps1
Analyzes C:\ and drills down into the largest folders.

.EXAMPLE
.\Find-LargestFolders.ps1 -StartPath "D:\Data" -MaxDepth 5 -TopCount 5
Analyzes D:\Data with max depth of 5 levels, showing top 5 items at each level.

.EXAMPLE
.\Find-LargestFolders.ps1 -StartPath "C:\Users" -MinSizeGB 1.0 -Debug
Analyzes C:\Users, only shows items larger than 1 GB, and enables debug output for detailed system overhead analysis.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $false, Position = 0)]
    [ValidateScript({
        if (-not (Test-Path -Path $_ -PathType Container)) {
            throw "Path '$_' does not exist or is not a directory."
        }
        return $true
    })]
    [string]$StartPath = "C:\",

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 100)]
    [int]$MaxDepth = 15,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 100)]
    [int]$TopCount = 3,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, [double]::MaxValue)]
    [double]$MinSizeGB = 0.1
)

begin {
    # Script variables
    $Script:StartTime = Get-Date
    $Script:ProcessedFolders = 0
    $Script:CurrentDepth = 0
    $Script:LargestFileFound = $null

    # Color codes for different PowerShell versions
if ($PSVersionTable.PSVersion.Major -ge 7) {
    $Script:Colors = @{
        Reset      = "`e[0m"
        White      = "`e[37m"
        Cyan       = "`e[36m"
        Green      = "`e[32m"
        Yellow     = "`e[33m"
        Red        = "`e[31m"
        Magenta    = "`e[35m"
        DarkGray   = "`e[90m"
        Bold       = "`e[1m"
    }
} else {
    $Script:Colors = @{
        Reset      = ""
        White      = ""
        Cyan       = ""
        Green      = ""
        Yellow     = ""
        Red        = ""
        Magenta    = ""
        DarkGray   = ""
        Bold       = ""
    }
}

function Format-FileSize {
    param(
        [Parameter(Mandatory = $true)]
        [long]$SizeInBytes
    )

    if ($SizeInBytes -ge 1TB) {
        return "{0:N2} TB" -f ($SizeInBytes / 1TB)
    } elseif ($SizeInBytes -ge 1GB) {
        return "{0:N2} GB" -f ($SizeInBytes / 1GB)
    } elseif ($SizeInBytes -ge 1MB) {
        return "{0:N2} MB" -f ($SizeInBytes / 1MB)
    } elseif ($SizeInBytes -ge 1KB) {
        return "{0:N2} KB" -f ($SizeInBytes / 1KB)
    } else {
        return "$SizeInBytes bytes"
    }
}

function Get-DirectorySize {
    param(
        [string]$Path
    )

    try {
        # Get all files in the directory (not subdirectories)
        $files = Get-ChildItem -Path $Path -File -Force -ErrorAction SilentlyContinue
        $totalSize = ($files | Measure-Object -Property Length -Sum).Sum

        if ($null -eq $totalSize) {
            $totalSize = 0
        }

        return [PSCustomObject]@{
            Path = $Path
            SizeBytes = $totalSize
            FileCount = $files.Count
            IsAccessible = $true
            Error = $null
        }
    } catch {
        return [PSCustomObject]@{
            Path = $Path
            SizeBytes = 0
            FileCount = 0
            IsAccessible = $false
            Error = $_.Exception.Message
        }
    }
}

function Get-SubdirectoryTotalSize {
    param(
        [string]$Path
    )

    try {
        # Get total size including all subdirectories, hidden files, and system files using Get-ChildItem -Recurse -Force
        $files = Get-ChildItem -Path $Path -File -Recurse -Force -ErrorAction SilentlyContinue
        $totalSize = ($files | Measure-Object -Property Length -Sum).Sum

        if ($null -eq $totalSize) {
            $totalSize = 0
        }

        Write-DebugInfo -Message "Directory '$Path' total size: $(Format-FileSize -SizeInBytes $totalSize)" -Category "SIZE"
        return $totalSize
    } catch {
        Write-Warning "Cannot access directory: $Path - $($_.Exception.Message)"
        return 0
    }
}

function Get-LargestItem {
    param(
        [string]$Path,
        [int]$TopCount,
        [double]$MinSizeGB
    )

    $results = @()
    $minSizeBytes = $MinSizeGB * 1GB

    try {
        Write-Output "$($Script:Colors.Cyan)Scanning: $Path$($Script:Colors.Reset)"
        # Get subdirectories (including hidden and system directories)
        $subdirs = Get-ChildItem -Path $Path -Directory -Force -ErrorAction SilentlyContinue

        foreach ($subdir in $subdirs) {
            $Script:ProcessedFolders++
            Write-Progress -Activity "Analyzing Folders" -Status "Processing: $($subdir.Name)" -PercentComplete -1

            $totalSize = Get-SubdirectoryTotalSize -Path $subdir.FullName

            if ($totalSize -ge $minSizeBytes) {
                $results += [PSCustomObject]@{
                    Type = "Folder"
                    Name = $subdir.Name
                    Path = $subdir.FullName
                    SizeBytes = $totalSize
                    SizeFormatted = Format-FileSize -SizeInBytes $totalSize
                }
            }
        }

        # Get files in current directory (including hidden and system files)
        $files = Get-ChildItem -Path $Path -File -Force -ErrorAction SilentlyContinue
        foreach ($file in $files) {
            if ($file.Length -ge $minSizeBytes) {
                $fileItem = [PSCustomObject]@{
                    Type = "File"
                    Name = $file.Name
                    Path = $file.FullName
                    SizeBytes = $file.Length
                    SizeFormatted = Format-FileSize -SizeInBytes $file.Length
                }
                $results += $fileItem

                # Track the largest file found globally
                if ($null -eq $Script:LargestFileFound -or $file.Length -gt $Script:LargestFileFound.SizeBytes) {
                    $Script:LargestFileFound = $fileItem
                }
            }
        }

        # Sort by size descending and take top items
        $topItems = $results | Sort-Object SizeBytes -Descending | Select-Object -First $TopCount

        return $topItems

    } catch {
        Write-Output "$($Script:Colors.Red)Error accessing path: $Path - $($_.Exception.Message)$($Script:Colors.Reset)"
        return @()
    }
}

function Show-LargestItemsTable {
    param(
        [array]$Items,
        [string]$Path,
        [int]$Depth
    )

    if ($Items.Count -eq 0) {
        Write-Output "$($Script:Colors.Yellow)No items found meeting the minimum size criteria.$($Script:Colors.Reset)"
        return
    }
    $indent = "  " * $Depth
    Write-Output ""
    Write-Output "$($Script:Colors.Bold)$($Script:Colors.Green)$indent=== LEVEL $($Depth + 1): $Path ===$($Script:Colors.Reset)"
    Write-Output ""
      # Sort by actual size in bytes (not formatted string) and filter out invalid items
    $sortedItems = $Items | Where-Object { 
        $_.Type -and $_.Name -and $_.SizeFormatted -and $_.SizeBytes -gt 0 
    } | Sort-Object @{Expression={$_.Type}; Descending=$true}, @{Expression={$_.SizeBytes}; Descending=$true}
    
    # Create simple table header
    Write-Output "Type   Name                        Size"
    Write-Output "----   ----                        ----"
      # Display each item with proper formatting
    foreach ($item in $sortedItems) {
        # Null safety for all properties
        $type = if ($item.Type) { $item.Type } else { "Unknown" }
        $name = if ($item.Name) { $item.Name } else { "Unknown" }
        $sizeFormatted = if ($item.SizeFormatted) { $item.SizeFormatted } else { "0 B" }
        $typeFormatted = $type.PadRight(6)
        $nameFormatted = if ($name.Length -gt 25) {
            $name.Substring(0, 22) + "..."
        } else {
            $name.PadRight(25)
        }
        $sizeFormattedPadded = $sizeFormatted.PadLeft(12)
        
        Write-Output "$typeFormatted $nameFormatted $sizeFormattedPadded"
    }
}

function Find-LargestFoldersRecursive {
    param(
        [string]$Path,
        [int]$CurrentDepth,
        [int]$MaxDepth,
        [int]$TopCount,
        [double]$MinSizeGB
    )

    if ($CurrentDepth -ge $MaxDepth) {
        Write-Output "$($Script:Colors.Yellow)Maximum depth ($MaxDepth) reached.$($Script:Colors.Reset)"
        return
    }
    # Get largest items at current level
    $largestItems = Get-LargestItem -Path $Path -TopCount $TopCount -MinSizeGB $MinSizeGB

    if ($largestItems.Count -eq 0) {
        Write-Output "$($Script:Colors.Yellow)No items found at current level meeting size criteria.$($Script:Colors.Reset)"
        return
    }

    # Show table for current level
    Show-LargestItemsTable -Items $largestItems -Path $Path -Depth $CurrentDepth

    # Find the largest folder (not file) to drill down into
    $largestFolder = $largestItems | Where-Object { $_.Type -eq "Folder" } | Select-Object -First 1
    if ($largestFolder) {
        # Recursively analyze the largest folder
        Find-LargestFoldersRecursive -Path $largestFolder.Path -CurrentDepth ($CurrentDepth + 1) -MaxDepth $MaxDepth -TopCount $TopCount -MinSizeGB $MinSizeGB
    } else {
        Write-Output "$($Script:Colors.Green)No more folders to drill down into at this level.$($Script:Colors.Reset)"
    }
}

function Start-AdvancedTranscript {
    [CmdletBinding(SupportsShouldProcess)]
    param([string]$LogPath)

    try {
        $computerName = $env:COMPUTERNAME
        $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $logFileName = "Find-LargestFolders_${computerName}_${timestamp}.log"
        $fullLogPath = Join-Path -Path $LogPath -ChildPath $logFileName

        if ($PSCmdlet.ShouldProcess($fullLogPath, "Start transcript log")) {
            # Ensure log directory exists
            if (-not (Test-Path -Path $LogPath)) {
                New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
            }
            # Create header
            $headerText = @"
===============================================
FIND LARGEST FOLDERS ANALYZER v1.2.3
===============================================
Log started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC')
Computer: $computerName
PowerShell Version: $($PSVersionTable.PSVersion)
Process ID: $PID
===============================================
"@
            Set-Content -Path $fullLogPath -Value $headerText -Encoding UTF8

            Start-Transcript -Path $fullLogPath -Append -Force | Out-Null
            Write-Output "$($Script:Colors.Green)Transcript started: $fullLogPath$($Script:Colors.Reset)"
            return $fullLogPath
        }
        return $null
    } catch {
        Write-Warning "Could not start transcript: $($_.Exception.Message)"
        return $null
    }
}

function Stop-AdvancedTranscript {
    [CmdletBinding(SupportsShouldProcess)]
    param()

    try {
        if ($PSCmdlet.ShouldProcess("transcript", "Stop transcript log")) {
            Stop-Transcript | Out-Null
            Write-Output "$($Script:Colors.Green)Transcript stopped successfully$($Script:Colors.Reset)"
        }
    } catch {
        Write-Warning "Error stopping transcript: $($_.Exception.Message)"
    }
}

function Write-DebugInfo {
    [CmdletBinding()]
    param(
        [string]$Message,
        [string]$Category = "DEBUG"
    )
    if ($DebugPreference -ne 'SilentlyContinue') {
        $timestamp = Get-Date -Format "HH:mm:ss.fff"
        $formattedMessage = "[$timestamp] [$Category] $Message"
        Write-Output "$($Script:Colors.Magenta)$formattedMessage$($Script:Colors.Reset)"
    }
}

function ConvertFrom-SizeString {
    <#
    .SYNOPSIS
    Converts size strings with units (GB, MB, KB, B) and returns size in bytes.
    .DESCRIPTION
    Helper function to convert size strings like "5.23 GB", "1,024 MB", etc. to bytes.
    .PARAMETER SizeText
    The size string to convert (e.g., "5.23 GB", "1,024 MB").
    .EXAMPLE
    ConvertFrom-SizeString -SizeText "5.23 GB"
    Returns the size in bytes equivalent to 5.23 GB.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$SizeText
    )
    try {
        # Remove common formatting characters and normalize
        $cleanText = $SizeText -replace ',', '' -replace '\s+', ' '
        # Match number followed by optional unit
        if ($cleanText -match '(\d+\.?\d*)\s*(GB|MB|KB|B|BYTES)?') {
            $value = [double]$matches[1]
            $unit = if ($matches[2]) { $matches[2].ToUpper() } else { "B" }
            switch ($unit) {
                "GB" { return [int64]($value * 1GB) }
                "MB" { return [int64]($value * 1MB) }
                "KB" { return [int64]($value * 1KB) }
                "B"  { return [int64]$value }
                default { return [int64]$value }
            }
        }
        return 0
    } catch {
        Write-DebugInfo -Message "Failed to convert size string '$SizeText': $($_.Exception.Message)" -Category "SIZE_CONVERT"
        return 0
    }
}

function Get-NTFSOverhead {
    <#
    .SYNOPSIS
    Retrieves comprehensive NTFS file system overhead information using fsutil fsinfo ntfsinfo.
    .DESCRIPTION
    Uses fsutil fsinfo ntfsinfo to extract detailed NTFS metadata including MFT size, reserved clusters,
    and other file system overhead that contributes to used space on the drive but is not accounted
    for in standard file enumeration.
    .PARAMETER DriveLetter
    The drive letter (without colon) to analyze for NTFS overhead information.
    .EXAMPLE
    Get-NTFSOverhead -DriveLetter "C"
    Returns NTFS overhead information for the C: drive.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$DriveLetter
    )
    try {
        $overhead = [PSCustomObject]@{
            MFTSize = 0
            TotalReservedClusters = 0
            StorageReservedClusters = 0
            MFTZoneSize = 0
            BytesPerCluster = 0
            TotalOverhead = 0
            EstimationMethod = "Unknown"
            RawNTFSInfo = @{
            }
        }
        # Execute fsutil fsinfo ntfsinfo to get comprehensive NTFS information
        try {
            Write-DebugInfo -Message "Executing fsutil fsinfo ntfsinfo ${DriveLetter}:" -Category "NTFS"
            $fsutilOutput = & fsutil fsinfo ntfsinfo "${DriveLetter}:" 2>$null
            if ($fsutilOutput -and $fsutilOutput.Count -gt 0) {
                Write-DebugInfo -Message "Successfully retrieved fsutil output with $($fsutilOutput.Count) lines" -Category "NTFS"
                foreach ($line in $fsutilOutput) {
                    $line = $line.Trim()
                    # Parse MFT Valid Data Length (actual MFT size in use)
                    if ($line -match "Mft Valid Data Length\s*:\s*(.+)") {
                        $mftSizeText = $matches[1].Trim()
                        Write-DebugInfo -Message "Found MFT Valid Data Length: '$mftSizeText'" -Category "NTFS"
                        # Parse size with unit (e.g., "1.01 GB")
                        if ($mftSizeText -match "(\d+\.?\d*)\s*(GB|MB|KB|B)") {
                            $mftValue = [double]$matches[1]
                            $mftUnit = $matches[2]
                            switch ($mftUnit) {
                                "GB" { $overhead.MFTSize = [int64]($mftValue * 1GB) }
                                "MB" { $overhead.MFTSize = [int64]($mftValue * 1MB) }
                                "KB" { $overhead.MFTSize = [int64]($mftValue * 1KB) }
                                "B"  { $overhead.MFTSize = [int64]$mftValue }
                            }
                            Write-DebugInfo -Message "Parsed MFT size: $($overhead.MFTSize) bytes" -Category "NTFS"
                        }
                    }
                    # Parse Total Reserved Clusters
                    elseif ($line -match "Total Reserved Clusters\s*:\s*([0-9,]+)\s*\(\s*(.+?)\s*\)") {
                        $reservedClustersText = $matches[1] -replace ',', ''
                        $reservedSizeText = $matches[2].Trim()
                        Write-DebugInfo -Message "Found Total Reserved Clusters: '$reservedClustersText' ($reservedSizeText)" -Category "NTFS"
                        $overhead.TotalReservedClusters = [int64]$reservedClustersText
                        # Parse the size in parentheses
                        if ($reservedSizeText -match "(\d+\.?\d*)\s*(GB|MB|KB|B)") {
                            $reservedValue = [double]$matches[1]
                            $reservedUnit = $matches[2]
                            switch ($reservedUnit) {
                                "GB" { $overhead.RawNTFSInfo['TotalReservedSize'] = [int64]($reservedValue * 1GB) }
                                "MB" { $overhead.RawNTFSInfo['TotalReservedSize'] = [int64]($reservedValue * 1MB) }
                                "KB" { $overhead.RawNTFSInfo['TotalReservedSize'] = [int64]($reservedValue * 1KB) }
                                "B"  { $overhead.RawNTFSInfo['TotalReservedSize'] = [int64]$reservedValue }
                            }
                        }
                    }
                    # Parse Bytes Per Cluster for calculations
                    elseif ($line -match "Bytes Per Cluster\s*:\s*([0-9,]+)") {
                        $bytesPerClusterText = $matches[1] -replace ',', ''
                        $overhead.BytesPerCluster = [int64]$bytesPerClusterText
                        Write-DebugInfo -Message "Found Bytes Per Cluster: $($overhead.BytesPerCluster)" -Category "NTFS"
                    }
                }
                # Calculate total NTFS overhead
                $overhead.TotalOverhead = $overhead.MFTSize
                if ($overhead.RawNTFSInfo['TotalReservedSize'] -gt 0) {
                    $overhead.TotalOverhead += $overhead.RawNTFSInfo['TotalReservedSize']
                }
                $overhead.EstimationMethod = "fsutil fsinfo ntfsinfo"
                Write-DebugInfo -Message "Total NTFS overhead calculated: $($overhead.TotalOverhead) bytes" -Category "NTFS"
            } else {
                Write-DebugInfo -Message "No output from fsutil or command failed" -Category "NTFS"
                $overhead.EstimationMethod = "No Data Available"
            }
        } catch {
            Write-DebugInfo -Message "Error executing fsutil: $($_.Exception.Message)" -Category "NTFS"
            $overhead.EstimationMethod = "Error: $($_.Exception.Message)"
        }
        return $overhead
    } catch {
        return [PSCustomObject]@{
            MFTSize = 0
            TotalReservedClusters = 0
            StorageReservedClusters = 0
            MFTZoneSize = 0
            BytesPerCluster = 0
            TotalOverhead = 0
            EstimationMethod = "Error: $($_.Exception.Message)"
            RawNTFSInfo = @{
            }
        }
    }
}

function Get-VSSOverhead {
    <#
    .SYNOPSIS
    Retrieves Volume Shadow Copy storage overhead information using vssadmin.
    .DESCRIPTION
    Uses vssadmin list shadowstorage to extract VSS storage allocation and usage information
    that contributes to used space on the drive but is not accounted for in standard file enumeration.
    This includes space reserved for shadow copies and currently used shadow copy storage.
    .PARAMETER DriveLetter
    The drive letter (without colon) to analyze for VSS overhead information.
    .EXAMPLE
    Get-VSSOverhead -DriveLetter "C"
    Returns VSS overhead information for the C: drive.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$DriveLetter
    )
    try {
        $vssInfo = [PSCustomObject]@{
            AllocatedSpace = 0
            UsedSpace = 0
            MaxSpace = 0
            ShadowCopyCount = 0
            TotalOverhead = 0
            EstimationMethod = "Unknown"
            RawVSSInfo = @{
            }
        }
        # Execute vssadmin list shadowstorage to get VSS storage information
        try {
            Write-DebugInfo -Message "Executing vssadmin list shadowstorage /for=${DriveLetter}:" -Category "VSS"
            $vssOutput = & vssadmin list shadowstorage /for="${DriveLetter}:" 2>$null
            if ($vssOutput -and $vssOutput.Count -gt 0) {
                Write-DebugInfo -Message "Successfully retrieved VSS output with $($vssOutput.Count) lines" -Category "VSS"
                $foundValidStorage = $false
                foreach ($line in $vssOutput) {
                    $line = $line.Trim()
                    # Parse Volume Shadow Copy Storage usage
                    if ($line -match "Used Shadow Copy Storage space:\s*(.+)") {
                        $usedSpaceText = $matches[1].Trim()
                        Write-DebugInfo -Message "Found Used Shadow Copy Storage: '$usedSpaceText'" -Category "VSS"
                        $parsedSize = ConvertFrom-SizeString -SizeText $usedSpaceText
                        if ($parsedSize -gt 0) {
                            $vssInfo.UsedSpace = $parsedSize
                            $foundValidStorage = $true
                        }
                    }
                    # Parse Allocated Shadow Copy Storage space
                    elseif ($line -match "Allocated Shadow Copy Storage space:\s*(.+)") {
                        $allocatedSpaceText = $matches[1].Trim()
                        Write-DebugInfo -Message "Found Allocated Shadow Copy Storage: '$allocatedSpaceText'" -Category "VSS"
                        $parsedSize = ConvertFrom-SizeString -SizeText $allocatedSpaceText
                        if ($parsedSize -gt 0) {
                            $vssInfo.AllocatedSpace = $parsedSize
                            $foundValidStorage = $true
                        }
                    }
                    # Parse Maximum Shadow Copy Storage space
                    elseif ($line -match "Maximum Shadow Copy Storage space:\s*(.+)") {
                        $maxSpaceText = $matches[1].Trim()
                        Write-DebugInfo -Message "Found Maximum Shadow Copy Storage: '$maxSpaceText'" -Category "VSS"
                        # Handle special cases like "UNBOUNDED"
                        if ($maxSpaceText -notmatch "UNBOUNDED|UNLIMITED") {
                            $parsedSize = ConvertFrom-SizeString -SizeText $maxSpaceText
                            if ($parsedSize -gt 0) {
                                $vssInfo.MaxSpace = $parsedSize
                            }
                        } else {
                            $vssInfo.RawVSSInfo['MaxSpaceUnbounded'] = $true
                        }
                    }
                }
                # Calculate total VSS overhead (use the larger of allocated or used space)
                $vssInfo.TotalOverhead = [Math]::Max($vssInfo.AllocatedSpace, $vssInfo.UsedSpace)
                if ($foundValidStorage) {
                    $vssInfo.EstimationMethod = "vssadmin list shadowstorage"
                    Write-DebugInfo -Message "Total VSS overhead calculated: $($vssInfo.TotalOverhead) bytes" -Category "VSS"
                } else {
                    $vssInfo.EstimationMethod = "No VSS Storage Found"
                }
            } else {
                Write-DebugInfo -Message "No VSS output or command failed" -Category "VSS"
                $vssInfo.EstimationMethod = "No Data Available"
            }
        } catch {
            Write-DebugInfo -Message "Error executing vssadmin: $($_.Exception.Message)" -Category "VSS"
            $vssInfo.EstimationMethod = "Error: $($_.Exception.Message)"
        }
        return $vssInfo
    } catch {
        return [PSCustomObject]@{
            AllocatedSpace = 0
            UsedSpace = 0
            MaxSpace = 0
            ShadowCopyCount = 0
            TotalOverhead = 0
            EstimationMethod = "Error: $($_.Exception.Message)"
            RawVSSInfo = @{
            }
        }
    }
}

function Get-SystemFilesSize {
    <#
    .SYNOPSIS
    Calculates the size of critical system files and directories that may not be included in regular scans.
    .DESCRIPTION
    Identifies and calculates the size of system files, hidden files, and special directories
    that contribute to drive usage but may be missed in standard directory enumeration.
    .PARAMETER DriveLetter
    The drive letter (without colon) to analyze for system files.
    .EXAMPLE
    Get-SystemFilesSize -DriveLetter "C"
    Returns system files size information for the C: drive.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$DriveLetter
    )
    try {
        $systemInfo = [PSCustomObject]@{
            PageFileSize = 0
            HibernationFileSize = 0
            SystemVolumeInfoSize = 0
            TempFilesSize = 0
            RecycleBinSize = 0
            TotalSystemSize = 0
            EstimationMethod = "Get-ChildItem with -Force"
            Details = @{
            }
        }

        $drivePath = "${DriveLetter}:"

        # Check for pagefile.sys
        $pageFilePath = Join-Path -Path $drivePath -ChildPath "pagefile.sys"
        if (Test-Path -Path $pageFilePath) {
            try {
                $pageFile = Get-Item -Path $pageFilePath -Force -ErrorAction SilentlyContinue
                if ($pageFile) {
                    $systemInfo.PageFileSize = $pageFile.Length
                    $systemInfo.Details['PageFile'] = $pageFile.Length
                    Write-DebugInfo -Message "Found pagefile.sys: $(Format-FileSize -SizeInBytes $pageFile.Length)" -Category "SYSTEM"
                }
            } catch {
                Write-DebugInfo -Message "Cannot access pagefile.sys: $($_.Exception.Message)" -Category "SYSTEM"
            }
        }

        # Check for hiberfil.sys
        $hiberFilePath = Join-Path -Path $drivePath -ChildPath "hiberfil.sys"
        if (Test-Path -Path $hiberFilePath) {
            try {
                $hiberFile = Get-Item -Path $hiberFilePath -Force -ErrorAction SilentlyContinue
                if ($hiberFile) {
                    $systemInfo.HibernationFileSize = $hiberFile.Length
                    $systemInfo.Details['HibernationFile'] = $hiberFile.Length
                    Write-DebugInfo -Message "Found hiberfil.sys: $(Format-FileSize -SizeInBytes $hiberFile.Length)" -Category "SYSTEM"
                }
            } catch {
                Write-DebugInfo -Message "Cannot access hiberfil.sys: $($_.Exception.Message)" -Category "SYSTEM"
            }
        }

        # Check System Volume Information (VSS snapshots location)
        $sviPath = Join-Path -Path $drivePath -ChildPath "System Volume Information"
        if (Test-Path -Path $sviPath) {
            try {
                $sviFiles = Get-ChildItem -Path $sviPath -File -Recurse -Force -ErrorAction SilentlyContinue
                if ($sviFiles) {
                    $sviSize = ($sviFiles | Measure-Object -Property Length -Sum).Sum
                    if ($sviSize -gt 0) {
                        $systemInfo.SystemVolumeInfoSize = $sviSize
                        $systemInfo.Details['SystemVolumeInfo'] = $sviSize
                        Write-DebugInfo -Message "Found System Volume Information: $(Format-FileSize -SizeInBytes $sviSize)" -Category "SYSTEM"
                    }
                }
            } catch {
                Write-DebugInfo -Message "Cannot access System Volume Information: $($_.Exception.Message)" -Category "SYSTEM"
            }
        }

        # Check $Recycle.Bin
        $recycleBinPath = Join-Path -Path $drivePath -ChildPath '$Recycle.Bin'
        if (Test-Path -Path $recycleBinPath) {
            try {
                $recycleFiles = Get-ChildItem -Path $recycleBinPath -File -Recurse -Force -ErrorAction SilentlyContinue
                if ($recycleFiles) {
                    $recycleSize = ($recycleFiles | Measure-Object -Property Length -Sum).Sum
                    if ($recycleSize -gt 0) {
                        $systemInfo.RecycleBinSize = $recycleSize
                        $systemInfo.Details['RecycleBin'] = $recycleSize
                        Write-DebugInfo -Message "Found Recycle Bin contents: $(Format-FileSize -SizeInBytes $recycleSize)" -Category "SYSTEM"
                    }
                }
            } catch {
                Write-DebugInfo -Message "Cannot access Recycle Bin: $($_.Exception.Message)" -Category "SYSTEM"
            }
        }

        # Calculate total system files size
        $systemInfo.TotalSystemSize = $systemInfo.PageFileSize + $systemInfo.HibernationFileSize +
                                     $systemInfo.SystemVolumeInfoSize + $systemInfo.RecycleBinSize

        Write-DebugInfo -Message "Total system files size: $(Format-FileSize -SizeInBytes $systemInfo.TotalSystemSize)" -Category "SYSTEM"

        return $systemInfo
    } catch {
        return [PSCustomObject]@{
            PageFileSize = 0
            HibernationFileSize = 0
            SystemVolumeInfoSize = 0
            TempFilesSize = 0
            RecycleBinSize = 0
            TotalSystemSize = 0
            EstimationMethod = "Error: $($_.Exception.Message)"
            Details = @{
            }
        }
    }
}

function Show-SystemOverhead {
    param(
        [string]$DriveLetter
    )

    Write-Output ""
    Write-Output "$($Script:Colors.Bold)$($Script:Colors.Green)System Overhead Analysis for Drive ${DriveLetter}:$($Script:Colors.Reset)"
    Write-Output "$($Script:Colors.Green)============================================$($Script:Colors.Reset)"

    # Get NTFS overhead
    $ntfsOverhead = Get-NTFSOverhead -DriveLetter $DriveLetter
    Write-Output "$($Script:Colors.White)NTFS Overhead:$($Script:Colors.Reset)"
    Write-Output "$($Script:Colors.White)  MFT Size: $($Script:Colors.Cyan)$(Format-FileSize -SizeInBytes $ntfsOverhead.MFTSize)$($Script:Colors.Reset)"
    if ($ntfsOverhead.RawNTFSInfo['TotalReservedSize'] -gt 0) {
        Write-Output "$($Script:Colors.White)  Reserved Clusters: $($Script:Colors.Cyan)$(Format-FileSize -SizeInBytes $ntfsOverhead.RawNTFSInfo['TotalReservedSize'])$($Script:Colors.Reset)"
    }
    Write-Output "$($Script:Colors.White)  Total NTFS Overhead: $($Script:Colors.Cyan)$(Format-FileSize -SizeInBytes $ntfsOverhead.TotalOverhead)$($Script:Colors.Reset)"

    # Get VSS overhead
    $vssOverhead = Get-VSSOverhead -DriveLetter $DriveLetter
    Write-Output ""
    Write-Output "$($Script:Colors.White)Volume Shadow Copy (VSS) Overhead:$($Script:Colors.Reset)"
    if ($vssOverhead.TotalOverhead -gt 0) {
        Write-Output "$($Script:Colors.White)  Used VSS Space: $($Script:Colors.Cyan)$(Format-FileSize -SizeInBytes $vssOverhead.UsedSpace)$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)  Allocated VSS Space: $($Script:Colors.Cyan)$(Format-FileSize -SizeInBytes $vssOverhead.AllocatedSpace)$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)  Total VSS Overhead: $($Script:Colors.Cyan)$(Format-FileSize -SizeInBytes $vssOverhead.TotalOverhead)$($Script:Colors.Reset)"
    } else {
        Write-Output "$($Script:Colors.White)  No VSS storage detected$($Script:Colors.Reset)"
    }

    # Get system files
    $systemFiles = Get-SystemFilesSize -DriveLetter $DriveLetter
    Write-Output ""
    Write-Output "$($Script:Colors.White)System Files:$($Script:Colors.Reset)"
    if ($systemFiles.PageFileSize -gt 0) {
        Write-Output "$($Script:Colors.White)  Page File: $($Script:Colors.Cyan)$(Format-FileSize -SizeInBytes $systemFiles.PageFileSize)$($Script:Colors.Reset)"
    }
    if ($systemFiles.HibernationFileSize -gt 0) {
        Write-Output "$($Script:Colors.White)  Hibernation File: $($Script:Colors.Cyan)$(Format-FileSize -SizeInBytes $systemFiles.HibernationFileSize)$($Script:Colors.Reset)"
    }
    if ($systemFiles.SystemVolumeInfoSize -gt 0) {
        Write-Output "$($Script:Colors.White)  System Volume Information: $($Script:Colors.Cyan)$(Format-FileSize -SizeInBytes $systemFiles.SystemVolumeInfoSize)$($Script:Colors.Reset)"
    }
    if ($systemFiles.RecycleBinSize -gt 0) {
        Write-Output "$($Script:Colors.White)  Recycle Bin: $($Script:Colors.Cyan)$(Format-FileSize -SizeInBytes $systemFiles.RecycleBinSize)$($Script:Colors.Reset)"
    }
    Write-Output "$($Script:Colors.White)  Total System Files: $($Script:Colors.Cyan)$(Format-FileSize -SizeInBytes $systemFiles.TotalSystemSize)$($Script:Colors.Reset)"

    # Calculate total overhead
    $totalOverhead = $ntfsOverhead.TotalOverhead + $vssOverhead.TotalOverhead + $systemFiles.TotalSystemSize
    Write-Output ""
    Write-Output "$($Script:Colors.Bold)$($Script:Colors.Yellow)Total System Overhead: $(Format-FileSize -SizeInBytes $totalOverhead)$($Script:Colors.Reset)"
    return $totalOverhead
}

function Show-DriveInfo {
    <#
    .SYNOPSIS
    Displays detailed drive information for a specified drive letter.
    .DESCRIPTION
    Retrieves and displays comprehensive drive information including:
    - Drive letter and label
    - File system type
    - Total, used, and free space
    - Health status
    Uses PowerShell's Get-Volume cmdlet.
    .PARAMETER DriveLetter
    The drive letter to display information for (without colon).
    .EXAMPLE
    Show-DriveInfo -DriveLetter "C"
    Displays information for the C: drive
    #>
    param (
        [Parameter(Mandatory=$true)]
        [ValidatePattern('^[A-Za-z]$')]
        [string]$DriveLetter
    )
    try {
        $volume = Get-Volume -DriveLetter $DriveLetter -ErrorAction Stop
        Write-Output "`n$($Script:Colors.Bold)$($Script:Colors.Green)Drive Volume Details for ${DriveLetter}:$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.Green)------------------------$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Drive Letter: $($Script:Colors.Cyan)$($volume.DriveLetter):$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Drive Label: $($Script:Colors.Cyan)$($volume.FileSystemLabel)$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)File System: $($Script:Colors.Cyan)$($volume.FileSystem)$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Drive Type: $($Script:Colors.Cyan)$($volume.DriveType)$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Size: $($Script:Colors.Cyan)$([math]::Round($volume.Size/1GB, 2)) GB$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Free Space: $($Script:Colors.Cyan)$([math]::Round($volume.SizeRemaining/1GB, 2)) GB$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Used Space: $($Script:Colors.Cyan)$([math]::Round(($volume.Size - $volume.SizeRemaining)/1GB, 2)) GB$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Free Space %: $($Script:Colors.Cyan)$([math]::Round(($volume.SizeRemaining/$volume.Size) * 100, 2))%$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Health Status: $($Script:Colors.Cyan)$($volume.HealthStatus)$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.White)Operational Status: $($Script:Colors.Cyan)$($volume.OperationalStatus)$($Script:Colors.Reset)"
        Write-Output ""
    } catch {
        Write-Output "$($Script:Colors.Red)Error retrieving drive information for ${DriveLetter}: $($_.Exception.Message)$($Script:Colors.Reset)"
    }
}

}
# End begin block

# Main execution
process {
    try {
        # WhatIf support
        if ($PSCmdlet.ShouldProcess($StartPath, "Analyze largest folders")) {

            # Start transcript logging
            $transcriptPath = Start-AdvancedTranscript -LogPath $PSScriptRoot
            # Normalize the start path - handle drive roots specially
            if ($StartPath -match '^[A-Za-z]:$') {
                # If just drive letter (e.g., "C:"), add backslash to make it root
                $StartPath = $StartPath + '\'
            } else {
                # Otherwise, trim any trailing backslashes except for drive roots
                $StartPath = $StartPath.TrimEnd('\')
                # Re-add backslash if it's a drive root
                if ($StartPath -match '^[A-Za-z]:$') {
                    $StartPath = $StartPath + '\'
                }
            }

            Write-Output "$($Script:Colors.Bold)$($Script:Colors.Green)Find Largest Folders Analyzer v1.2.3$($Script:Colors.Reset)"
            Write-Output "$($Script:Colors.Green)===============================================$($Script:Colors.Reset)"
            Write-Output "$($Script:Colors.White)Start Path: $($Script:Colors.Cyan)$StartPath$($Script:Colors.Reset)"
            Write-Output "$($Script:Colors.White)Max Depth: $($Script:Colors.Cyan)$MaxDepth$($Script:Colors.Reset)"
            Write-Output "$($Script:Colors.White)Top Count: $($Script:Colors.Cyan)$TopCount$($Script:Colors.Reset)"
            Write-Output "$($Script:Colors.White)Min Size: $($Script:Colors.Cyan)$MinSizeGB GB$($Script:Colors.Reset)"
            Write-Output "$($Script:Colors.White)Started: $($Script:Colors.Cyan)$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')$($Script:Colors.Reset)"
            Write-Output ""
            # Validate start path
            if (-not (Test-Path -Path $StartPath -PathType Container)) {
                throw "Start path does not exist or is not accessible: $StartPath"
            }

            # Extract drive letter for system overhead analysis
            $driveLetter = if ($StartPath -match '^([A-Za-z]):') { $matches[1] } else { $null }

            # Show system overhead if analyzing a drive root
            if ($driveLetter -and $StartPath -eq "${driveLetter}:\") {
                $totalOverhead = Show-SystemOverhead -DriveLetter $driveLetter
                Write-Output ""
            }

            # Start the recursive analysis
            Find-LargestFoldersRecursive -Path $StartPath -CurrentDepth 0 -MaxDepth $MaxDepth -TopCount $TopCount -MinSizeGB $MinSizeGB

            # Final summary
            $endTime = Get-Date
            $duration = $endTime - $Script:StartTime
            Write-Output ""
            Write-Output "$($Script:Colors.Bold)$($Script:Colors.Green)Analysis Complete$($Script:Colors.Reset)"
            Write-Output "$($Script:Colors.Green)==================$($Script:Colors.Reset)"
            Write-Output "$($Script:Colors.White)Folders Processed: $($Script:Colors.Cyan)$Script:ProcessedFolders$($Script:Colors.Reset)"
            Write-Output "$($Script:Colors.White)Total Duration: $($Script:Colors.Cyan)$($duration.ToString('hh\:mm\:ss'))$($Script:Colors.Reset)"
            Write-Output "$($Script:Colors.White)Completed: $($Script:Colors.Cyan)$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')$($Script:Colors.Reset)"

            # Report the largest file found
            if ($Script:LargestFileFound) {
                Write-Output ""
                Write-Output "$($Script:Colors.Bold)$($Script:Colors.Yellow)Largest File Found During Scan:$($Script:Colors.Reset)"
                Write-Output "$($Script:Colors.Yellow)======================================$($Script:Colors.Reset)"
                Write-Output "$($Script:Colors.White)File: $($Script:Colors.Cyan)$($Script:LargestFileFound.Name)$($Script:Colors.Reset)"
                Write-Output "$($Script:Colors.White)Size: $($Script:Colors.Cyan)$($Script:LargestFileFound.SizeFormatted)$($Script:Colors.Reset)"
                Write-Output "$($Script:Colors.White)Path: $($Script:Colors.Cyan)$($Script:LargestFileFound.Path)$($Script:Colors.Reset)"
            }

            # Show drive information at the end if we have a drive letter
            if ($driveLetter) {
                Show-DriveInfo -DriveLetter $driveLetter
            }

            if ($transcriptPath) {
                Write-Output "$($Script:Colors.White)Log File: $($Script:Colors.Cyan)$transcriptPath$($Script:Colors.Reset)"
            }

            Write-Output ""
        }
    } catch {
        Write-Output "$($Script:Colors.Red)Script execution failed: $($_.Exception.Message)$($Script:Colors.Reset)"
        Write-Output "$($Script:Colors.Red)$($_.ScriptStackTrace)$($Script:Colors.Reset)"
    } finally {
        # Stop transcript
        Stop-AdvancedTranscript
    }
}
