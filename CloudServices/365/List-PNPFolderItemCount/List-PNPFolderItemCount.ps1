# =============================================================================
# Script: List-PNPFolderItemCount.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1.1
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
# =============================================================================

<#
.SYNOPSIS
Lists all files, folders, and subfolders in a SharePoint Online site with detailed item counts.

.DESCRIPTION
This script connects to a SharePoint Online site using PnP PowerShell and recursively
enumerates all folders, subfolders, and items within a specified parent folder.
It produces a table showing the path, type (file/folder), and total item counts.

.PARAMETER SiteUrl
The URL of the SharePoint site to connect to.

.PARAMETER ClientId
The client ID used for authentication.

.PARAMETER ParentFolder
The path to the parent folder in SharePoint to scan.

.EXAMPLE
.\List-PNPCrawl.ps1
Lists all items in the specified SharePoint folder and outputs results to the console and a CSV file.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$SiteUrl = "SITEURL",

    [Parameter()]
    [string]$ClientId = "CLIENTID",

    [Parameter()]
    [string]$ParentFolder = "Shared Documents/LIT/NUVO Backup"
)

# Script initialization
$ErrorActionPreference = 'Stop'
$currentTime = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$computerName = $env:COMPUTERNAME
$scriptPath = $PSScriptRoot
$logFile = Join-Path -Path $scriptPath -ChildPath "PNPCrawl_${ computerName}_${ currentTime}.log"
$csvOutputFile = Join-Path -Path $scriptPath -ChildPath "PNPCrawl_Results_${computerName}_${currentTime}.csv"
$results = [System.Collections.ArrayList]@()

# Start logging
Start-Transcript -Path $logFile

try {
    # Connect to SharePoint site with PnP
    Write-Information -MessageData "Connecting to SharePoint site: $SiteUrl" -InformationAction Continue
    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId
    # Function to recursively get folder contents and item counts
    Write-Information -MessageData "Successfully connected to SharePoint" -InformationAction Continue
    function Get-FolderContentsRecursive {
        param (
            [Parameter(Mandatory = $true)]
            [string]$FolderPath,

            [Parameter()]
            [int]$Level = 0
        )

        Write-Progress -Activity "Scanning SharePoint Structure" -Status "Processing $FolderPath" -PercentComplete -1
        Write-Information -MessageData "Scanning folder: $FolderPath" -InformationAction Continue

        # Get items in the current folder - using more specific PnP commands for better reliability
        try {
            # First, add the folder itself to the results
            # Don't add the parent folder
            if ($Level -gt 0) {
                $folderObj = Get-PnPFolder -Url $FolderPath -ErrorAction Stop
                $folderInfo = [PSCustomObject]@{
                    Path = $FolderPath
                    Name = Split-Path -Path $FolderPath -Leaf
                    Type = "Folder"
                    # Will be updated later
                    ItemCount = 0
                    Size = 0
                    LastModified = $folderObj.TimeLastModified
                    Level = $Level
                }
                [void]$results.Add($folderInfo)
            }

            # Get files in this folder
            $files = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderPath -ItemType File -ErrorAction Stop

            if ($null -ne $files) {
                Write-Information -MessageData "Found $($files.Count) files in $FolderPath" -InformationAction Continue

                foreach ($file in $files) {
                    $itemPath = "$FolderPath/$($file.Name)"

                    # Add the file to results
                    $fileInfo = [PSCustomObject]@{
                        Path = $itemPath
                        Name = $file.Name
                        Type = "File"
                        ItemCount = 1
                        Size = $file.Length
                        LastModified = $file.TimeLastModified
                        Level = $Level
                    }

                    [void]$results.Add($fileInfo)
                }
            }

            # Get subfolders and process them recursively
            $subfolders = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderPath -ItemType Folder -ErrorAction Stop

            if ($null -ne $subfolders) {
                Write-Information -MessageData "Found $($subfolders.Count) subfolders in $FolderPath" -InformationAction Continue

                foreach ($subfolder in $subfolders) {
                    $subfolderPath = "$FolderPath/$($subfolder.Name)"

                    # Process this subfolder recursively
                    Get-FolderContentsRecursive -FolderPath $subfolderPath -Level ($Level + 1)
                }
            }

            # Now update the folder item count
            # Don't update the parent folder here
            if ($Level -gt 0) {
                $folderItems = $results | Where-Object {
                    $_.Path -like "$FolderPath*" -and
                    $_.Path -ne $FolderPath -and
                    # Count only files, not subfolders
                    $_.Type -eq "File"
                }
                $itemCount = ($folderItems | Measure-Object).Count
                ($results | Where-Object { $_.Path -eq $FolderPath }).ItemCount = $itemCount
            }
        } catch {
            Write-Warning -Message "Error processing folder $FolderPath : $_"
        }
    }
    # Start the recursive scan from the parent folder
    Write-Information -MessageData "Starting scan of $ParentFolder" -InformationAction Continue

    # First, check if the parent folder exists
    try {
        $parentFolderObj = Get-PnPFolder -Url $ParentFolder -ErrorAction Stop
        if ($null -eq $parentFolderObj) {
            throw "Parent folder not found"
        }

        Write-Information -MessageData "Parent folder found, beginning recursive scan" -InformationAction Continue

        # Start the recursive processing
        Get-FolderContentsRecursive -FolderPath $ParentFolder
        Write-Information -MessageData "Scan completed. Found $($results.Count) items" -InformationAction Continue

        # Process and display results
        Write-Information -MessageData "Processing results..." -InformationAction Continue

        # Get direct children of the parent folder for the summary
        $topLevelItems = $results | Where-Object {
            ($_.Path -like "$ParentFolder/*" -and
            ($_.Path -notlike "$ParentFolder/*/*" -or
            ($_.Path -split "/" | Measure-Object).Count -eq ($ParentFolder -split "/" | Measure-Object).Count + 1))
        }

        Write-Information -MessageData "`nTop-level folders and their item counts:" -InformationAction Continue
        $topLevelItems | Sort-Object -Property ItemCount -Descending |
            Select-Object Name, Type, ItemCount, @{ Name = "SizeKB"; Expression = { [math]::Round($_.Size / 1KB, 2) } } |
            Format-Table -AutoSize

        # Detailed results - show as a tree structure with indentation based on level
        Write-Information -MessageData "`nDetailed structure:" -InformationAction Continue

        $sortedResults = $results | Sort-Object Path
        foreach ($item in $sortedResults) {
            $indent = "    " * $item.Level
            $displaySize = if ($item.Type -eq "File") { " ({ 0:N2} KB)" -f ($item.Size / 1KB) } else { "" }
            $itemCountDisplay = if ($item.Type -eq "Folder") { " [{ 0} items]" -f $item.ItemCount } else { "" }
            Write-Information -MessageData (" { 0}{ 1} - { 2}{ 3}{ 4}" -f $indent, $item.Name, $item.Type, $itemCountDisplay, $displaySize) -InformationAction Continue
        }

        # Folder summary for quick reference
        Write-Information -MessageData "`nFolder Summary (sorted by item count):" -InformationAction Continue
        $folderSummary = $results | Where-Object { $_.Type -eq "Folder" } |
            Sort-Object -Property ItemCount -Descending |
            Select-Object Path, ItemCount
        $folderSummary | Format-Table -AutoSize
    } catch {
        Write-Error -Message "Error accessing parent folder $ParentFolder : $_"
    }

    # Export results to CSV
    $results | Export-Csv -Path $csvOutputFile -NoTypeInformation
    Write-Information -MessageData "Results exported to $csvOutputFile" -InformationAction Continue
} catch {
    Write-Error -Message "An error occurred: $_"
    Write-Error -Message $_.Exception.StackTrace
} finally {
    # Disconnect from SharePoint
    try {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        Write-Information -MessageData "Disconnected from SharePoint" -InformationAction Continue
    } catch {
        Write-Warning -Message "Error disconnecting from SharePoint: $_"
    }

    # Stop logging
    Stop-Transcript
}
