# =============================================================================
# Script: Find-LargestFolders.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-31 18:11:49 UTC
# Updated By: jdyer-nuvodia
# Version: 1.8.0
# Additional Info: Fixed critical OneDrive online-only file detection using fsutil reparse point tags and comprehensive attribute checking to prevent massive size discrepancies
# =============================================================================

<#
.SYNOPSIS
Efficiently finds the largest folders by recursively drilling down into only the largest subdirectories, showing size and last write time while accurately handling OneDrive online-only files.

.DESCRIPTION
This script scans a directory and identifies the largest subdirectories and files without scanning the entire drive.
It now includes comprehensive system overhead detection and intelligent OneDrive online-only file handling including:
- NTFS file system overhead (MFT, reserved clusters)
- Volume Shadow Copy (VSS) storage overhead
- System files (pagefile.sys, hiberfil.sys, System Volume Information, Recycle Bin)
- Hidden and system files/directories using the -Force parameter
- Last write time for both files and folders to help identify unused space consumers
- Smart OneDrive online-only file detection that distinguishes between actual placeholders and legitimate reparse points (junctions, symlinks)

The script works by:
1. Scanning the top-level directories in the specified path (including hidden/system directories)
2. Identifying the largest folders and files (including hidden/system files)
3. Displaying size, type, and last write time for each item
4. Recursively drilling down into only the largest folder
5. Repeating this process until reaching the specified depth or finding no more subdirectories
6. Displaying system overhead information when analyzing drive roots

This approach is much faster than scanning entire drives and focuses on finding the largest space consumers while accounting for all system overhead and providing write time information to identify potentially unused files and folders.

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
Analyzes C:\ and drills down into the largest folders, displaying size and last write time.

.EXAMPLE
.\Find-LargestFolders.ps1 -StartPath "D:\Data" -MaxDepth 5 -TopCount 5
Analyzes D:\Data with max depth of 5 levels, showing top 5 items at each level with write times.

.EXAMPLE
.\Find-LargestFolders.ps1 -StartPath "C:\Users" -MinSizeGB 1.0 -Debug
Analyzes C:\Users, only shows items larger than 1 GB with write times, and enables debug output for detailed system overhead analysis.
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
    $Script:StartTime          = Get-Date
    $Script:ProcessedFolders   = 0
    $Script:CurrentDepth       = 0
    $Script:LargestFileFound   = $null
    # Color codes for different PowerShell versions
    if ($PSVersionTable.PSVersion.Major -ge 7) {
        $Script:Colors = @{
            Reset    = "`e[0m"
            White    = "`e[37m"
            Cyan     = "`e[36m"
            Green    = "`e[32m"
            Yellow   = "`e[33m"
            Red      = "`e[31m"
            Magenta  = "`e[35m"
            DarkGray = "`e[90m"
            Bold     = "`e[1m"
        }
        $Script:UseAnsiColors = $true
    } else {
        # PowerShell 5.1 - Use console color mapping
        $Script:Colors = @{
            Reset    = ""
            White    = "White"
            Cyan     = "Cyan"
            Green    = "Green"
            Yellow   = "Yellow"
            Red      = "Red"
            Magenta  = "Magenta"
            DarkGray = "DarkGray"
            Bold     = ""
        }
        $Script:UseAnsiColors = $false
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

    function Format-WriteTime {
        <#
        .SYNOPSIS
        Formats a DateTime object into a readable string for write time display.
        .DESCRIPTION
        Converts a DateTime object to a formatted string showing the last write time.
        Returns a shortened format that fits well in table displays.
        .PARAMETER DateTime
        The DateTime object to format.
        .EXAMPLE
        Format-writeTime -DateTime (Get-Date)
        Returns formatted write time string.
        #>
        param(
            [Parameter(Mandatory = $true)]
            [DateTime]$DateTime
        )

        try {
            # Format as MM/dd/yyyy HH:mm for compact display
            return $DateTime.ToString("MM/dd/yyyy HH:mm")
        } catch {
            return "Unknown"
        }
    }

    function Get-DirectorySize {
        param(
            [string]$Path
        )

        try {
            # Get all files in the directory (not subdirectories)
            $files = Get-ChildItem -Path $Path -File -Force -ErrorAction SilentlyContinue
            $totalSizeSum = $files | Measure-Object -Property Length -Sum
            $totalSize = if ($totalSizeSum.Sum) { $totalSizeSum.Sum } else { 0 }

            return [PSCustomObject]@{
                Path         = $Path
                SizeBytes    = $totalSize
                FileCount    = $files.Count
                IsAccessible = $true
                Error        = $null
            }
        } catch {
            return [PSCustomObject]@{
                Path         = $Path
                SizeBytes    = 0
                FileCount    = 0
                IsAccessible = $false
                Error        = $_.Exception.Message
            }
        }
    }

    function Test-OneDriveOnlineOnly {
        <#
        .SYNOPSIS
        Accurately determines if a file is a OneDrive online-only placeholder file.
        .DESCRIPTION
        Uses comprehensive detection methods to identify OneDrive online-only files while preserving
        legitimate reparse points (symbolic links, junctions, etc.). This function addresses the issue
        where OneDrive online-only files show their full cloud size in the Length property, causing
        massive size discrepancies (e.g., 1.94TB reported on a 235GB drive).

        Detection methods:
        1. Checks for ReparsePoint attribute (required for OneDrive placeholders)
        2. Uses fsutil to get reparse point tag (OneDrive-specific tags: 0x9000601a, 0x9000701a)
        3. Checks for OneDrive-specific file attributes (Offline, RecallOnOpen, RecallOnDataAccess)
        4. Validates OneDrive path context
        5. Uses .NET FileAttributes for additional validation

        This function specifically distinguishes OneDrive placeholders from:
        - Symbolic links (tag 0xa000000c)
        - Junction points (tag 0xa0000003)
        - Mount points (tag 0xa0000003)
        - Hard links (no reparse point)

        .PARAMETER FileInfo
        The FileInfo object to test.
        .EXAMPLE
        Test-OneDriveOnlineOnly -FileInfo $fileInfo
        Returns $true if the file is a OneDrive online-only file, $false otherwise.
        #>
        param(
            [Parameter(Mandatory = $true)]
            [System.IO.FileInfo]$FileInfo
        )

        # First check if it has ReparsePoint attribute - required for all OneDrive placeholders
        if (-not ($FileInfo.Attributes -band [System.IO.FileAttributes]::ReparsePoint)) {
            return $false
        }

        # Check for OneDrive-specific attributes that indicate online-only status
        $hasOfflineAttribute = $FileInfo.Attributes -band [System.IO.FileAttributes]::Offline
        $hasRecallOnOpenAttribute = $FileInfo.Attributes -band [System.IO.FileAttributes]::RecallOnOpen
        $hasRecallOnDataAccessAttribute = $FileInfo.Attributes -band [System.IO.FileAttributes]::RecallOnDataAccess

        # Strong indicator: Files with Offline + RecallOnOpen/RecallOnDataAccess are OneDrive online-only
        if ($hasOfflineAttribute -and ($hasRecallOnOpenAttribute -or $hasRecallOnDataAccessAttribute)) {
            Write-DebugInfo -Message "OneDrive online-only detected via attributes: $($FileInfo.FullName)" -Category "ONEDRIVE"
            return $true
        }

        # Secondary check: Use fsutil to get the specific reparse point tag
        try {
            $reparseInfo = & fsutil reparsepoint query "$($FileInfo.FullName)" 2>$null | Out-String
            if ($reparseInfo) {
                # OneDrive Files On-Demand uses specific reparse point tags:
                # 0x9000601a - OneDrive online-only file
                # 0x9000701a - OneDrive partially downloaded file
                if ($reparseInfo -match "Tag value: 0x9000[67]01a") {
                    Write-DebugInfo -Message "OneDrive online-only detected via fsutil tag: $($FileInfo.FullName)" -Category "ONEDRIVE"
                    return $true
                }
                # These are NOT OneDrive files (legitimate reparse points):
                # 0xa000000c - Symbolic link
                # 0xa0000003 - Junction point/Mount point
                elseif ($reparseInfo -match "Tag value: 0xa000000[3c]") {
                    Write-DebugInfo -Message "Legitimate reparse point detected (symlink/junction): $($FileInfo.FullName)" -Category "REPARSE"
                    return $false
                }
            }
        } catch {
            # If fsutil fails, continue with other checks
            Write-DebugInfo -Message "fsutil query failed for $($FileInfo.FullName): $($_.Exception.Message)" -Category "ONEDRIVE"
        }

        # Tertiary check: OneDrive path context with additional validation
        $isOneDrivePath = $FileInfo.FullName -match "OneDrive|SkyDrive"
        if ($isOneDrivePath) {
            # In OneDrive paths, files with ReparsePoint + Offline attributes are likely online-only
            if ($hasOfflineAttribute) {
                Write-DebugInfo -Message "OneDrive online-only detected via path+attributes: $($FileInfo.FullName)" -Category "ONEDRIVE"
                return $true
            }

            # Additional OneDrive detection: Check for cloud icon overlay or sparse files
            if ($FileInfo.Attributes -band [System.IO.FileAttributes]::SparseFile) {
                Write-DebugInfo -Message "OneDrive sparse file detected: $($FileInfo.FullName)" -Category "ONEDRIVE"
                return $true
            }
        }

        # Final check: Use .NET File.GetAttributes for additional attribute detection
        try {
            $netAttributes = [System.IO.File]::GetAttributes($FileInfo.FullName)
            $hasNotContentIndexed = $netAttributes -band [System.IO.FileAttributes]::NotContentIndexed

            # NotContentIndexed + ReparsePoint + OneDrive path often indicates online-only
            if ($isOneDrivePath -and $hasNotContentIndexed) {
                Write-DebugInfo -Message "OneDrive online-only detected via .NET attributes: $($FileInfo.FullName)" -Category "ONEDRIVE"
                return $true
            }
        } catch {
            # Continue if .NET attributes check fails - log the error for debugging
            Write-DebugInfo -Message ".NET File.GetAttributes failed for $($FileInfo.FullName): $($_.Exception.Message)" -Category "ONEDRIVE"
        }

        # If we reach here, it's likely a legitimate reparse point (symlink, junction, etc.)
        Write-DebugInfo -Message "Legitimate reparse point preserved: $($FileInfo.FullName)" -Category "REPARSE"
        return $false
    }

    function Get-SubdirectoryTotalSize {
        param(
            [string]$Path
        )

        try {
            # Get total size including all subdirectories, hidden files, and system files using Get-ChildItem -Recurse -Force
            $files = Get-ChildItem -Path $Path -File -Recurse -Force -ErrorAction SilentlyContinue

            # Filter out OneDrive online-only files more intelligently
            $localFiles = $files | Where-Object {
                -not (Test-OneDriveOnlineOnly -FileInfo $_)
            }

            $totalSizeSum = $localFiles | Measure-Object -Property Length -Sum
            $totalSize = if ($totalSizeSum.Sum) { $totalSizeSum.Sum } else { 0 }
            $onlineOnlyCount = $files.Count - $localFiles.Count
            $reparsePointFiles = $files | Where-Object { $_.Attributes -band [System.IO.FileAttributes]::ReparsePoint }
            $legitimateReparsePoints = $reparsePointFiles.Count - $onlineOnlyCount

            if ($onlineOnlyCount -gt 0) {
                Write-DebugInfo -Message "Directory '$Path' - Excluded $onlineOnlyCount OneDrive online-only files" -Category "ONEDRIVE"
            }
            if ($legitimateReparsePoints -gt 0) {
                Write-DebugInfo -Message "Directory '$Path' - Included $legitimateReparsePoints legitimate reparse points" -Category "REPARSE"
            }

            $formattedSize = Format-FileSize -SizeInBytes $totalSize
            Write-DebugInfo -Message "Directory '$Path' local size: $formattedSize" -Category "SIZE"

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
            Write-ColorOutput -Message "Scanning: $Path" -Color "Cyan"
            # Get subdirectories (including hidden and system directories)
            $subdirs = Get-ChildItem -Path $Path -Directory -Force -ErrorAction SilentlyContinue

            foreach ($subdir in $subdirs) {
                $Script:ProcessedFolders++
                Write-Progress -Activity "Analyzing Folders" -Status "Processing: $($subdir.Name)" -PercentComplete -1

                $totalSize = Get-SubdirectoryTotalSize -Path $subdir.FullName

                if ($totalSize -ge $minSizeBytes) {
                    $results += [PSCustomObject]@{
                        Type               = "Folder"
                        Name               = $subdir.Name
                        Path               = $subdir.FullName
                        SizeBytes          = $totalSize
                        SizeFormatted      = Format-FileSize -SizeInBytes $totalSize
                        lastwriteTime      = $subdir.lastwriteTime
                        lastwriteFormatted = Format-writeTime -DateTime $subdir.lastwriteTime
                    }
                }
            }

            # Get files in current directory (including hidden and system files)
            $files = Get-ChildItem -Path $Path -File -Force -ErrorAction SilentlyContinue

            # Filter out OneDrive online-only files more intelligently
            $localFiles = $files | Where-Object {
                -not (Test-OneDriveOnlineOnly -FileInfo $_)
            }

            foreach ($file in $localFiles) {
                if ($file.Length -ge $minSizeBytes) {
                    $fileItem = [PSCustomObject]@{
                        Type               = "File"
                        Name               = $file.Name
                        Path               = $file.FullName
                        SizeBytes          = $file.Length
                        SizeFormatted      = Format-FileSize -SizeInBytes $file.Length
                        lastwriteTime      = $file.lastwriteTime
                        lastwriteFormatted = Format-writeTime -DateTime $file.lastwriteTime
                    }
                    $results += $fileItem

                    # Track the largest file found globally
                    if ($null -eq $Script:LargestFileFound -or $file.Length -gt $Script:LargestFileFound.SizeBytes) {
                        $Script:LargestFileFound = $fileItem
                    }
                }
            }

            # Report different types of reparse points if any were found
            $reparsePointFiles = $files | Where-Object { $_.Attributes -band [System.IO.FileAttributes]::ReparsePoint }
            if ($reparsePointFiles.Count -gt 0) {
                $onlineOnlyFiles = $reparsePointFiles | Where-Object { Test-OneDriveOnlineOnly -FileInfo $_ }
                $legitimateReparsePoints = $reparsePointFiles | Where-Object { -not (Test-OneDriveOnlineOnly -FileInfo $_) }

                if ($onlineOnlyFiles.Count -gt 0) {
                    Write-DebugInfo -Message "Found $($onlineOnlyFiles.Count) OneDrive online-only files in '$Path'" -Category "ONEDRIVE"
                }
                if ($legitimateReparsePoints.Count -gt 0) {
                    Write-DebugInfo -Message "Found $($legitimateReparsePoints.Count) legitimate reparse points in '$Path'" -Category "REPARSE"
                }
            }

            # Sort by size descending and take top items
            $topItems = $results | Sort-Object SizeBytes -Descending | Select-Object -First $TopCount

            return $topItems
        } catch {
            Write-ColorOutput -Message "Error accessing path: $Path - $($_.Exception.Message)" -Color "Red"
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
            Write-ColorOutput -Message "No items found meeting the minimum size criteria." -Color "Yellow"
            return
        }
        $indent = "  " * $Depth
        Write-Output ""
        Write-ColorOutput -Message "$indent=== LEVEL $($Depth + 1): $Path ===" -Color "Green"
        Write-Output ""
        # Sort by actual size in bytes (not formatted string) and filter out invalid items
        $sortedItems = $Items | Where-Object {
            $_.Type -and $_.Name -and $_.SizeFormatted -and $_.SizeBytes -gt 0
        } | Sort-Object @{ Expression = { $_.Type }; Descending = $true }, @{ Expression = { $_.SizeBytes }; Descending = $true }

        # Create simple table header with write time
        Write-Output "Type   Name                 Size        Last Write"
        Write-Output "----   ----                 ----        -------------"
        # Display each item with proper formatting
        foreach ($item in $sortedItems) {
            # Null safety for all properties
            $type = if ($item.Type) { $item.Type } else { "Unknown" }
            $name = if ($item.Name) { $item.Name } else { "Unknown" }
            $sizeFormatted = if ($item.SizeFormatted) { $item.SizeFormatted } else { "0 B" }
            $writeFormatted = if ($item.lastwriteFormatted) { $item.lastwriteFormatted } else { "Unknown" }

            $typeFormatted = $type.PadRight(6)
            $nameFormatted = if ($name.Length -gt 20) {
                $name.Substring(0, 17) + "..."
            } else {
                $name.PadRight(20)
            }
            $sizeFormattedPadded = $sizeFormatted.PadRight(11)
            $writeFormattedPadded = $writeFormatted.PadRight(16)

            Write-Output "$typeFormatted $nameFormatted $sizeFormattedPadded $writeFormattedPadded"
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
            Write-ColorOutput -Message "Maximum depth ($MaxDepth) reached." -Color "Yellow"
            return
        }
        # Get largest items at current level
        $largestItems = Get-LargestItem -Path $Path -TopCount $TopCount -MinSizeGB $MinSizeGB
        if ($largestItems.Count -eq 0) {
            Write-ColorOutput -Message "No items found at current level meeting size criteria." -Color "Yellow"
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
            Write-ColorOutput -Message "No more folders to drill down into at this level." -Color "Green"
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
FIND LARGEST FOLDERS ANALYZER v1.7.0
===============================================
Log started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC')
Computer: $computerName
PowerShell Version: $($PSVersionTable.PSVersion)
Process ID: $PID
===============================================
"@
                Set-Content -Path $fullLogPath -Value $headerText -Encoding UTF8
                Start-Transcript -Path $fullLogPath -Append -Force | Out-Null
                Write-ColorOutput -Message "Transcript started: $fullLogPath" -Color "Green"
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
                Write-ColorOutput -Message "Transcript stopped successfully" -Color "Green"
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
            # Use Write-ColorOutput but redirect to error stream to avoid data flow interference
            Write-ColorOutput -Message $formattedMessage -Color "Magenta" *>&1 | Out-Host
        }
    }

    function ConvertFrom-SizeString {
        <#
        .SYNOPSIS
        Converts size strings with units (GB, MB, KB, B) and returns size in bytes.
        .DESCRIPTION
        Helper function to convert size strings like "5.23 GB", "1, 024 MB", etc. to bytes.
        .PARAMETER SizeText
        The size string to convert (e.g., "5.23 GB", "1, 024 MB").
        .EXAMPLE
        ConvertFrom-SizeString -SizeText "5.23 GB"
        Returns the size in bytes equivalent to 5.23 GB.
        #>
        param(
            [Parameter(Mandatory = $true)]
            [string]$SizeText
        )
        try {
            # Remove common formatting characters and normalize
            $cleanText = $SizeText -replace ', ', '' -replace '\s+', ' '
            # Match number followed by optional unit
            if ($cleanText -match '(\d+\.?\d*)\s*(GB|MB|KB|B|BYTES)?') {
                $value = [double]$matches[1]
                $unit = if ($matches[2]) { $matches[2].ToUpper() } else { "B" }
                switch ($unit) {
                    "GB" { return [int64]($value * 1GB) }
                    "MB" { return [int64]($value * 1MB) }
                    "KB" { return [int64]($value * 1KB) }
                    "B" { return [int64]$value }
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
            [Parameter(Mandatory = $true)]
            [string]$DriveLetter
        )
        try {
            $overhead = [PSCustomObject]@{
                MFTSize                 = 0
                TotalReservedClusters   = 0
                StorageReservedClusters = 0
                MFTZoneSize             = 0
                BytesPerCluster         = 0
                TotalOverhead           = 0
                EstimationMethod        = "Unknown"
                RawNTFSInfo             = @{
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
                                    "B" { $overhead.MFTSize = [int64]$mftValue }
                                }
                                Write-DebugInfo -Message "Parsed MFT size: $($overhead.MFTSize) bytes" -Category "NTFS"
                            }
                        }
                        # Parse Total Reserved Clusters
                        elseif ($line -match "Total Reserved Clusters\s*:\s*([0-9, ]+)\s*\(\s*(.+?)\s*\)") {
                            $reservedClustersText = $matches[1] -replace ', ', ''
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
                                    "B" { $overhead.RawNTFSInfo['TotalReservedSize'] = [int64]$reservedValue }
                                }
                            }
                        }
                        # Parse Bytes Per Cluster for calculations
                        elseif ($line -match "Bytes Per Cluster\s*:\s*([0-9, ]+)") {
                            $bytesPerClusterText = $matches[1] -replace ', ', ''
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
                MFTSize                 = 0
                TotalReservedClusters   = 0
                StorageReservedClusters = 0
                MFTZoneSize             = 0
                BytesPerCluster         = 0
                TotalOverhead           = 0
                EstimationMethod        = "Error: $($_.Exception.Message)"
                RawNTFSInfo             = @{
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
            [Parameter(Mandatory = $true)]
            [string]$DriveLetter
        )
        try {
            $vssInfo = [PSCustomObject]@{
                AllocatedSpace   = 0
                UsedSpace        = 0
                MaxSpace         = 0
                ShadowCopyCount  = 0
                TotalOverhead    = 0
                EstimationMethod = "Unknown"
                RawVSSInfo       = @{
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
                AllocatedSpace   = 0
                UsedSpace        = 0
                MaxSpace         = 0
                ShadowCopyCount  = 0
                TotalOverhead    = 0
                EstimationMethod = "Error: $($_.Exception.Message)"
                RawVSSInfo       = @{
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
            [Parameter(Mandatory = $true)]
            [string]$DriveLetter
        )
        try {
            $systemInfo = [PSCustomObject]@{
                PageFileSize         = 0
                HibernationFileSize  = 0
                SystemVolumeInfoSize = 0
                TempFilesSize        = 0
                RecycleBinSize       = 0
                TotalSystemSize      = 0
                EstimationMethod     = "Get-ChildItem with -Force"
                Details              = @{
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
                        $formattedSize = Format-FileSize -SizeInBytes $pageFile.Length
                        Write-DebugInfo -Message "Found pagefile.sys: $formattedSize" -Category "SYSTEM"
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
                        $formattedSize = Format-FileSize -SizeInBytes $hiberFile.Length
                        Write-DebugInfo -Message "Found hiberfil.sys: $formattedSize" -Category "SYSTEM"
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
                        $sviSizeSum = $sviFiles | Measure-Object -Property Length -Sum
                        $sviSize = if ($sviSizeSum.Sum) { $sviSizeSum.Sum } else { 0 }
                        if ($sviSize -gt 0) {
                            $systemInfo.SystemVolumeInfoSize = $sviSize
                            $systemInfo.Details['SystemVolumeInfo'] = $sviSize
                            $formattedSize = Format-FileSize -SizeInBytes $sviSize
                            Write-DebugInfo -Message "Found System Volume Information: $formattedSize" -Category "SYSTEM"
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
                        $recycleSizeSum = $recycleFiles | Measure-Object -Property Length -Sum
                        $recycleSize = if ($recycleSizeSum.Sum) { $recycleSizeSum.Sum } else { 0 }
                        if ($recycleSize -gt 0) {
                            $systemInfo.RecycleBinSize = $recycleSize
                            $systemInfo.Details['RecycleBin'] = $recycleSize
                            $formattedSize = Format-FileSize -SizeInBytes $recycleSize
                            Write-DebugInfo -Message "Found Recycle Bin contents: $formattedSize" -Category "SYSTEM"
                        }
                    }
                } catch {
                    Write-DebugInfo -Message "Cannot access Recycle Bin: $($_.Exception.Message)" -Category "SYSTEM"
                }
            }

            # Calculate total system files size
            $systemInfo.TotalSystemSize = $systemInfo.PageFileSize + $systemInfo.HibernationFileSize + `
                $systemInfo.SystemVolumeInfoSize + $systemInfo.RecycleBinSize

            $formattedSize = Format-FileSize -SizeInBytes $systemInfo.TotalSystemSize
            Write-DebugInfo -Message "Total system files size: $formattedSize" -Category "SYSTEM"

            return $systemInfo
        } catch {
            return [PSCustomObject]@{
                PageFileSize         = 0
                HibernationFileSize  = 0
                SystemVolumeInfoSize = 0
                TempFilesSize        = 0
                RecycleBinSize       = 0
                TotalSystemSize      = 0
                EstimationMethod     = "Error: $($_.Exception.Message)"
                Details              = @{
                }
            }
        }
    }

    function Show-SystemOverhead {
        param(
            [string]$DriveLetter
        )
        Write-Output ""
        Write-ColorOutput -Message "System Overhead Analysis for Drive ${DriveLetter}:" -Color "Green"
        Write-ColorOutput -Message "============================================" -Color "Green"

        # Get NTFS overhead
        $ntfsOverhead = Get-NTFSOverhead -DriveLetter $DriveLetter
        Write-ColorOutput -Message "NTFS Overhead:" -Color "White"
        Write-ColorOutput -Message "  MFT Size: $(Format-FileSize -SizeInBytes $ntfsOverhead.MFTSize)" -Color "White"
        if ($ntfsOverhead.RawNTFSInfo['TotalReservedSize'] -gt 0) {
            Write-ColorOutput -Message "  Reserved Clusters: $(Format-FileSize -SizeInBytes $ntfsOverhead.RawNTFSInfo['TotalReservedSize'])" -Color "White"
        }
        Write-ColorOutput -Message "  Total NTFS Overhead: $(Format-FileSize -SizeInBytes $ntfsOverhead.TotalOverhead)" -Color "White"

        # Get VSS overhead
        $vssOverhead = Get-VSSOverhead -DriveLetter $DriveLetter
        Write-Output ""
        Write-ColorOutput -Message "Volume Shadow Copy (VSS) Overhead:" -Color "White"
        if ($vssOverhead.TotalOverhead -gt 0) {
            Write-ColorOutput -Message "  Used VSS Space: $(Format-FileSize -SizeInBytes $vssOverhead.UsedSpace)" -Color "White"
            Write-ColorOutput -Message "  Allocated VSS Space: $(Format-FileSize -SizeInBytes $vssOverhead.AllocatedSpace)" -Color "White"
            Write-ColorOutput -Message "  Total VSS Overhead: $(Format-FileSize -SizeInBytes $vssOverhead.TotalOverhead)" -Color "White"
        } else {
            Write-ColorOutput -Message "  No VSS storage detected" -Color "White"
        }

        # Get system files
        $systemFiles = Get-SystemFilesSize -DriveLetter $DriveLetter
        Write-Output ""
        Write-ColorOutput -Message "System Files:" -Color "White"
        if ($systemFiles.PageFileSize -gt 0) {
            Write-ColorOutput -Message "  Page File: $(Format-FileSize -SizeInBytes $systemFiles.PageFileSize)" -Color "White"
        }
        if ($systemFiles.HibernationFileSize -gt 0) {
            Write-ColorOutput -Message "  Hibernation File: $(Format-FileSize -SizeInBytes $systemFiles.HibernationFileSize)" -Color "White"
        }
        if ($systemFiles.SystemVolumeInfoSize -gt 0) {
            Write-ColorOutput -Message "  System Volume Information: $(Format-FileSize -SizeInBytes $systemFiles.SystemVolumeInfoSize)" -Color "White"
        }
        if ($systemFiles.RecycleBinSize -gt 0) {
            Write-ColorOutput -Message "  Recycle Bin: $(Format-FileSize -SizeInBytes $systemFiles.RecycleBinSize)" -Color "White"
        }
        Write-ColorOutput -Message "  Total System Files: $(Format-FileSize -SizeInBytes $systemFiles.TotalSystemSize)" -Color "White"

        # Calculate total overhead
        $totalOverhead = $ntfsOverhead.TotalOverhead + $vssOverhead.TotalOverhead + $systemFiles.TotalSystemSize
        Write-Output ""
        Write-ColorOutput -Message "Total System Overhead: $(Format-FileSize -SizeInBytes $totalOverhead)" -Color "Yellow"
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
            [Parameter(Mandatory = $true)]
            [ValidatePattern('^[A-Za-z]$')]
            [string]$DriveLetter
        )
        try {
            $volume = Get-Volume -DriveLetter $DriveLetter -ErrorAction Stop
            Write-Output ""
            Write-ColorOutput -Message "Drive Volume Details for ${DriveLetter}:" -Color "Green"
            Write-ColorOutput -Message "------------------------" -Color "Green"
            Write-ColorOutput -Message "Drive Letter: $($volume.DriveLetter):" -Color "White"
            Write-ColorOutput -Message "Drive Label: $($volume.FileSystemLabel)" -Color "White"
            Write-ColorOutput -Message "File System: $($volume.FileSystem)" -Color "White"
            Write-ColorOutput -Message "Drive Type: $($volume.DriveType)" -Color "White"
            Write-ColorOutput -Message "Size: $([math]::Round($volume.Size/1GB, 2)) GB" -Color "White"
            Write-ColorOutput -Message "Free Space: $([math]::Round($volume.SizeRemaining/1GB, 2)) GB" -Color "White"
            Write-ColorOutput -Message "Used Space: $([math]::Round(($volume.Size - $volume.SizeRemaining)/1GB, 2)) GB" -Color "White"
            Write-ColorOutput -Message "Free Space %: $([math]::Round(($volume.SizeRemaining/$volume.Size) * 100, 2))%" -Color "White"
            Write-ColorOutput -Message "Health Status: $($volume.HealthStatus)" -Color "White"
            Write-ColorOutput -Message "Operational Status: $($volume.OperationalStatus)" -Color "White"
            Write-Output ""
        } catch {
            Write-ColorOutput -Message "Error retrieving drive information for ${DriveLetter}: $($_.Exception.Message)" -Color "Red"
        }
    }

    function Write-ColorOutput {
        <#
        .SYNOPSIS
        Writes colored output that works in both PowerShell 5.1 and 7+ without using Write-Host.
        .DESCRIPTION
        Uses ANSI escape codes for PowerShell 7+ and console color changes for PowerShell 5.1.
        This function complies with copilot-instructions.md by using Write-Output instead of Write-Host.
        .PARAMETER Message
        The message to write with color.
        .PARAMETER Color
        The color to use (White, Cyan, Green, Yellow, Red, Magenta, DarkGray).
        .EXAMPLE
        Write-ColorOutput -Message "Success!" -Color "Green"
        Writes "Success!" in green color.
        #>
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
            Write-ColorOutput -Message "Find Largest Folders Analyzer v1.8.0" -Color "Green"
            Write-ColorOutput -Message "===============================================" -Color "Green"
            Write-ColorOutput -Message "Start Path: $StartPath" -Color "White"
            Write-ColorOutput -Message "Max Depth: $MaxDepth" -Color "White"
            Write-ColorOutput -Message "Top Count: $TopCount" -Color "White"
            Write-ColorOutput -Message "Min Size: $MinSizeGB GB" -Color "White"
            Write-ColorOutput -Message "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Color "White"
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
            Write-ColorOutput -Message "Analysis Complete" -Color "Green"
            Write-ColorOutput -Message "==================" -Color "Green"
            Write-ColorOutput -Message "Folders Processed: $Script:ProcessedFolders" -Color "White"
            Write-ColorOutput -Message "Total Duration: $($duration.ToString('hh\:mm\:ss'))" -Color "White"
            Write-ColorOutput -Message "Completed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Color "White"
            # Report the largest file found
            if ($Script:LargestFileFound) {
                Write-Output ""
                Write-ColorOutput -Message "Largest File Found During Scan:" -Color "Yellow"
                Write-ColorOutput -Message "======================================" -Color "Yellow"
                Write-ColorOutput -Message "File: $($Script:LargestFileFound.Name)" -Color "White"
                Write-ColorOutput -Message "Size: $($Script:LargestFileFound.SizeFormatted)" -Color "White"
                Write-ColorOutput -Message "Path: $($Script:LargestFileFound.Path)" -Color "White"
            }

            # Show drive information at the end if we have a drive letter
            if ($driveLetter) {
                Show-DriveInfo -DriveLetter $driveLetter
            }
            if ($transcriptPath) {
                Write-ColorOutput -Message "Log File: $transcriptPath" -Color "White"
            }

            Write-Output ""
        }
    } catch {
        Write-ColorOutput -Message "Script execution failed: $($_.Exception.Message)" -Color "Red"
        Write-ColorOutput -Message "$($_.ScriptStackTrace)" -Color "Red"
    } finally {
        # Stop transcript
        Stop-AdvancedTranscript
    }
}
