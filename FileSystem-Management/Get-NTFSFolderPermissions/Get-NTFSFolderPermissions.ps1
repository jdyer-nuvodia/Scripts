# =============================================================================
# Script: Get-NTFSFolderPermissions.ps1
# Created: 2025-03-06 21:06:43 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-06 21:19:10 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2
# Additional Info: Script to extract and report NTFS permissions with consolidated folder display
# =============================================================================

<#
.SYNOPSIS
    Extracts and reports NTFS permissions for a specified folder and its subfolders.
.DESCRIPTION
    This script retrieves NTFS permissions for a specified folder path and all its subfolders.
    It provides a detailed report including identity references, file system rights, access control types,
    and inheritance settings. The results are displayed in the console as separate tables for each folder
    (omitting subfolders with identical permissions) and exported to a text file in the same directory
    as the script.
    
    - Traverses the specified folder structure recursively
    - Captures all NTFS permission entries for each folder
    - Groups folders with identical permissions to reduce output clutter
    - Exports results to a text file in the script's directory with the same format as console output
.PARAMETER FolderPath
    The path to the folder for which permissions will be extracted. This parameter is mandatory.
.EXAMPLE
    .\Get-NTFSFolderPermissions.ps1 -FolderPath "C:\Important\Data"
    Retrieves NTFS permissions for C:\Important\Data and all subfolders, and exports the results to a text file.
.EXAMPLE
    .\Get-NTFSFolderPermissions.ps1 -FolderPath "\\server\share\folder"
    Retrieves NTFS permissions for the specified network share and displays results.
.NOTES
    Security Level: Low
    Required Permissions: Read access to the folders being scanned
    Validation Requirements: Verify FolderPath exists and is accessible
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$FolderPath
)

# Get the script's directory to use for output files
$ScriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$OutputTxt = Join-Path -Path $ScriptDirectory -ChildPath "NTFSPermissions_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').txt"

# Initialize an array to store results
$Results = @()

# Helper function to compare two permission sets
function Compare-PermissionSets {
    param (
        [Array]$Set1,
        [Array]$Set2
    )
    
    if ($Set1.Count -ne $Set2.Count) { return $false }
    
    # Get sorted string representations of both sets for comparison
    $SortedSet1 = ($Set1 | Sort-Object IdentityReference, FileSystemRights, AccessControlType, IsInherited | 
                 ForEach-Object { "$($_.IdentityReference)|$($_.FileSystemRights)|$($_.AccessControlType)|$($_.IsInherited)" }) -join ";"
    
    $SortedSet2 = ($Set2 | Sort-Object IdentityReference, FileSystemRights, AccessControlType, IsInherited | 
                 ForEach-Object { "$($_.IdentityReference)|$($_.FileSystemRights)|$($_.AccessControlType)|$($_.IsInherited)" }) -join ";"
    
    return $SortedSet1 -eq $SortedSet2
}

# Helper function to check if a folder is a direct subfolder of another
function Is-DirectSubfolder {
    param (
        [string]$ParentPath,
        [string]$PotentialChildPath
    )
    
    # Convert paths to consistent format
    $NormalizedParent = $ParentPath.TrimEnd('\').ToLower()
    $NormalizedChild = $PotentialChildPath.TrimEnd('\').ToLower()
    
    # Skip self comparison
    if ($NormalizedParent -eq $NormalizedChild) { return $false }
    
    # Check if child path starts with parent path
    if ($NormalizedChild.StartsWith($NormalizedParent)) {
        # Get remaining path after parent
        $Remaining = $NormalizedChild.Substring($NormalizedParent.Length).TrimStart('\')
        
        # If there's only one folder level difference (no additional '\')
        return ($Remaining -notcontains '\')
    }
    
    return $false
}

# Start a string builder for text file output
$OutputText = [System.Text.StringBuilder]::new()
[void]$OutputText.AppendLine("NTFS Permissions Report - Generated $(Get-Date)")
[void]$OutputText.AppendLine("Folder Path: $FolderPath")
[void]$OutputText.AppendLine("=" * 80)
[void]$OutputText.AppendLine("")

# Display start message
Write-Host "Starting NTFS permissions analysis for: $FolderPath" -ForegroundColor Cyan
[void]$OutputText.AppendLine("Starting NTFS permissions analysis for: $FolderPath")

try {
    # Get all folders and subfolders recursively
    Write-Host "Retrieving folder structure..." -ForegroundColor Cyan
    [void]$OutputText.AppendLine("Retrieving folder structure...")
    
    $Folders = Get-ChildItem -Path $FolderPath -Recurse -Directory -ErrorAction Stop
    
    # Include the root folder in the list
    $Folders += Get-Item -Path $FolderPath -ErrorAction Stop
    
    $TotalFolders = ($Folders | Measure-Object).Count
    Write-Host "Found $TotalFolders folders to process" -ForegroundColor Cyan
    [void]$OutputText.AppendLine("Found $TotalFolders folders to process")
    
    $CurrentFolder = 0
    
    # Dictionary to store permissions by folder path
    $FolderPermissionsMap = @{}
    
    # Loop through each folder and retrieve NTFS permissions
    foreach ($Folder in $Folders) {
        $CurrentFolder++
        Write-Host "Processing folder ($CurrentFolder/$TotalFolders): $($Folder.FullName)" -ForegroundColor DarkGray
        
        try {
            $Acl = Get-Acl -Path $Folder.FullName -ErrorAction Stop
            $FolderPermissions = @()
            
            foreach ($Access in $Acl.Access) {
                # Create a custom object for each permission entry
                $Result = [PSCustomObject]@{
                    FolderPath       = $Folder.FullName
                    IdentityReference = $Access.IdentityReference
                    FileSystemRights  = $Access.FileSystemRights
                    AccessControlType = $Access.AccessControlType
                    IsInherited       = $Access.IsInherited
                    InheritanceFlags  = $Access.InheritanceFlags
                    PropagationFlags  = $Access.PropagationFlags
                }
                
                # Add the result to overall results array
                $Results += $Result
                
                # Add to folder-specific array
                $FolderPermissions += $Result
            }
            
            # Store the permissions for this folder path
            $FolderPermissionsMap[$Folder.FullName] = $FolderPermissions
        }
        catch {
            Write-Host "Error processing folder: $($Folder.FullName)" -ForegroundColor Yellow
            Write-Host "Error details: $_" -ForegroundColor Yellow
            [void]$OutputText.AppendLine("Error processing folder: $($Folder.FullName)")
            [void]$OutputText.AppendLine("Error details: $_")
        }
    }
    
    # Display completion message
    Write-Host "Analysis completed. Found $($Results.Count) permission entries across $TotalFolders folders." -ForegroundColor Green
    [void]$OutputText.AppendLine("Analysis completed. Found $($Results.Count) permission entries across $TotalFolders folders.")
    
    # Display results grouped by folder with separate tables
    Write-Host "`nDisplaying permissions by folder:" -ForegroundColor Cyan
    [void]$OutputText.AppendLine("")
    [void]$OutputText.AppendLine("Displaying permissions by folder:")
    
    # Get all folder paths and sort them by depth (for parent-child relationship checking)
    $SortedFolderPaths = $FolderPermissionsMap.Keys | Sort-Object { ($_ -split '\\').Count }
    
    # Keep track of folders already displayed
    $DisplayedFolders = @{}
    $SkippedFolders = @()
    
    foreach ($FolderPath in $SortedFolderPaths) {
        # Skip if already processed as part of a group
        if ($DisplayedFolders.ContainsKey($FolderPath)) {
            continue
        }
        
        $CurrentFolderPermissions = $FolderPermissionsMap[$FolderPath]
        
        # Create a visual separator
        $SeparatorLength = [Math]::Min(100, $FolderPath.Length + 10)
        $Separator = "-" * $SeparatorLength
        
        Write-Host "`n$Separator" -ForegroundColor White
        Write-Host "Folder: $FolderPath" -ForegroundColor White
        Write-Host "$Separator" -ForegroundColor White
        
        [void]$OutputText.AppendLine("")
        [void]$OutputText.AppendLine($Separator)
        [void]$OutputText.AppendLine("Folder: $FolderPath")
        [void]$OutputText.AppendLine($Separator)
        
        # Find all child folders with identical permissions
        $IdenticalSubfolders = @()
        foreach ($OtherPath in $SortedFolderPaths) {
            # Skip self or already displayed
            if (($OtherPath -eq $FolderPath) -or ($DisplayedFolders.ContainsKey($OtherPath))) {
                continue
            }
            
            # Check if the folder is somewhere in the subfolder tree
            if ($OtherPath.StartsWith($FolderPath + "\")) {
                $OtherPermissions = $FolderPermissionsMap[$OtherPath]
                
                # Compare permissions
                if (Compare-PermissionSets -Set1 $CurrentFolderPermissions -Set2 $OtherPermissions) {
                    $IdenticalSubfolders += $OtherPath
                    $DisplayedFolders[$OtherPath] = $true
                    $SkippedFolders += $OtherPath
                }
            }
        }
        
        # Display the permissions
        $SimplifiedPermissions = $CurrentFolderPermissions | Select-Object IdentityReference, FileSystemRights, AccessControlType, IsInherited
        $PermissionsTable = $SimplifiedPermissions | Format-Table -AutoSize | Out-String
        
        Write-Host $PermissionsTable
        [void]$OutputText.Append($PermissionsTable)
        
        # Mark this folder as displayed
        $DisplayedFolders[$FolderPath] = $true
        
        # If there are subfolders with identical permissions, list them
        if ($IdenticalSubfolders.Count -gt 0) {
            Write-Host "The following subfolders have identical permissions:" -ForegroundColor Cyan
            [void]$OutputText.AppendLine("The following subfolders have identical permissions:")
            
            foreach ($Subfolder in $IdenticalSubfolders) {
                Write-Host "  - $Subfolder" -ForegroundColor DarkGray
                [void]$OutputText.AppendLine("  - $Subfolder")
            }
        }
    }
    
    # Report skipped folders
    $SkippedCount = $SkippedFolders.Count
    Write-Host "`nSkipped displaying $SkippedCount folders with permissions identical to their parent folders." -ForegroundColor Cyan
    [void]$OutputText.AppendLine("")
    [void]$OutputText.AppendLine("Skipped displaying $SkippedCount folders with permissions identical to their parent folders.")
    
    # Save the output to text file
    $OutputText.ToString() | Out-File -FilePath $OutputTxt -Encoding UTF8
    Write-Host "`nPermissions report exported to: $OutputTxt" -ForegroundColor Green
}
catch {
    Write-Host "An error occurred during the NTFS permissions analysis:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    [void]$OutputText.AppendLine("An error occurred during the NTFS permissions analysis:")
    [void]$OutputText.AppendLine($_.Exception.Message)
    
    # Try to save what we have so far
    if ($OutputText.Length -gt 0) {
        $OutputText.ToString() | Out-File -FilePath $OutputTxt -Encoding UTF8
    }
}
