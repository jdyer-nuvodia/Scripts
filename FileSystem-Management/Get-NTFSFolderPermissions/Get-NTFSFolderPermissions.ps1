# =============================================================================
# Script: Get-NTFSFolderPermissions.ps1
# Created: 2025-03-06 21:06:43 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-06 22:13:45 UTC
# Updated By: jdyer-nuvodia
# Version: 1.3.2
# Additional Info: Fixed parenthesis syntax error in progress update condition
# =============================================================================

<#
.SYNOPSIS
    Extracts and reports NTFS permissions for a specified folder and its subfolders with optimized performance.
.DESCRIPTION
    This script retrieves NTFS permissions for a specified folder path and all its subfolders.
    It provides a detailed report including identity references, file system rights, access control types,
    and inheritance settings. The results are displayed in the console as separate tables for each folder
    (omitting subfolders with identical permissions) and exported to a text file in the same directory
    as the script.
    
    - Uses optimized directory traversal methods for better performance
    - Processes folders in parallel with configurable thread limit
    - Captures all NTFS permission entries for each folder
    - Groups folders with identical permissions to reduce output clutter
    - Exports results to a text file in the script's directory with the same format as console output
    
    Performance improvements:
    - Uses .NET methods for faster directory traversal
    - Implements parallel processing with runspaces
    - Optimized permission comparison logic
    - Includes memory management improvements
.PARAMETER FolderPath
    The path to the folder for which permissions will be extracted. This parameter is mandatory.
.PARAMETER MaxThreads
    Maximum number of parallel threads to use for processing. Default is 10.
.PARAMETER MaxDepth
    Maximum folder depth to traverse. Set to 0 for unlimited depth. Default is 0.
.PARAMETER SkipUniquenessCounting
    Skip counting unique permissions for large directories to improve performance. Default is false.
.EXAMPLE
    .\Get-NTFSFolderPermissions.ps1 -FolderPath "C:\Important\Data"
    Retrieves NTFS permissions for C:\Important\Data and all subfolders, and exports the results to a text file.
.EXAMPLE
    .\Get-NTFSFolderPermissions.ps1 -FolderPath "\\server\share\folder" -MaxThreads 20
    Uses 20 parallel threads to process folders, improving performance on network shares.
.EXAMPLE
    .\Get-NTFSFolderPermissions.ps1 -FolderPath "C:\VeryLargeFolder" -MaxDepth 3
    Only processes folders up to a maximum depth of 3 levels from the root folder.
.NOTES
    Security Level: Low
    Required Permissions: Read access to the folders being scanned
    Validation Requirements: Verify FolderPath exists and is accessible
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$FolderPath,
    
    [Parameter(Mandatory = $false)]
    [int]$MaxThreads = 10,
    
    [Parameter(Mandatory = $false)]
    [int]$MaxDepth = 0,
    
    [Parameter(Mandatory = $false)]
    [switch]$SkipUniquenessCounting
)

# Get the script's directory to use for output files
$ScriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$OutputTxt = Join-Path -Path $ScriptDirectory -ChildPath "NTFSPermissions_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').txt"

# Start a string builder for text file output
$OutputText = [System.Text.StringBuilder]::new()
[void]$OutputText.AppendLine("NTFS Permissions Report - Generated $(Get-Date)")
[void]$OutputText.AppendLine("Folder Path: $FolderPath")
[void]$OutputText.AppendLine("=" * 80)
[void]$OutputText.AppendLine("")

# Display start message with optimization info
Write-Host "Starting optimized NTFS permissions analysis for: $FolderPath" -ForegroundColor Cyan
Write-Host "Using up to $MaxThreads parallel threads" -ForegroundColor DarkGray
if ($MaxDepth -gt 0) {
    Write-Host "Limited to maximum depth of $MaxDepth levels" -ForegroundColor DarkGray
}
[void]$OutputText.AppendLine("Starting NTFS permissions analysis for: $FolderPath")

# Helper function to compare two permission sets efficiently
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

# Helper function to generate a permissions hash for faster comparison
function Get-PermissionsHash {
    param (
        [Array]$Permissions
    )
    
    $SortedPermissions = ($Permissions | Sort-Object IdentityReference, FileSystemRights, AccessControlType, IsInherited | 
                        ForEach-Object { "$($_.IdentityReference)|$($_.FileSystemRights)|$($_.AccessControlType)|$($_.IsInherited)" }) -join ";"
    
    return $SortedPermissions.GetHashCode()
}

# Function to get all subdirectories with fast .NET methods
function Get-SubdirectoriesFast {
    param (
        [string]$Path,
        [int]$CurrentDepth = 0,
        [int]$MaxDepth = 0
    )

    try {
        # Add the current directory
        $CurrentDir = [System.IO.DirectoryInfo]::new($Path)
        
        # Return if we've reached max depth
        if ($MaxDepth -gt 0 -and $CurrentDepth -ge $MaxDepth) {
            return @($CurrentDir)
        }
        
        # Get immediate subdirectories
        $SubDirs = [System.IO.Directory]::GetDirectories($Path)
        
        # Process each subdirectory recursively
        $AllDirs = @($CurrentDir)
        foreach ($Dir in $SubDirs) {
            try {
                $AllDirs += Get-SubdirectoriesFast -Path $Dir -CurrentDepth ($CurrentDepth + 1) -MaxDepth $MaxDepth
            }
            catch {
                # Silently continue if we can't access a directory
            }
        }
        
        return $AllDirs
    }
    catch {
        # Return just the current directory if we can't process subdirectories
        if (Test-Path $Path -PathType Container) {
            return @([System.IO.DirectoryInfo]::new($Path))
        }
        return @()
    }
}

# Helper function to get permissions for a folder (used in runspaces)
function Get-FolderPermissionsWorker {
    param (
        [string]$FolderPath
    )
    
    try {
        $Acl = [System.Security.AccessControl.DirectorySecurity]::new()
        $Acl.SetAccessRuleProtection($false, $false)
        
        $Acl = [System.IO.Directory]::GetAccessControl($FolderPath)
        $Permissions = @()
        
        foreach ($Access in $Acl.Access) {
            # Create a custom object for each permission entry
            $Permission = [PSCustomObject]@{
                FolderPath       = $FolderPath
                IdentityReference = $Access.IdentityReference
                FileSystemRights  = $Access.FileSystemRights
                AccessControlType = $Access.AccessControlType
                IsInherited       = $Access.IsInherited
                InheritanceFlags  = $Access.InheritanceFlags
                PropagationFlags  = $Access.PropagationFlags
            }
            
            $Permissions += $Permission
        }
        
        return @{
            Success = $true
            FolderPath = $FolderPath
            Permissions = $Permissions
            PermissionsHash = (Get-PermissionsHash -Permissions $Permissions)
            Error = $null
        }
    }
    catch {
        return @{
            Success = $false
            FolderPath = $FolderPath
            Permissions = @()
            PermissionsHash = 0
            Error = $_.Exception.Message
        }
    }
}

try {
    # Get all folders and subfolders recursively using optimized method
    Write-Host "Retrieving folder structure (optimized method)..." -ForegroundColor Cyan
    [void]$OutputText.AppendLine("Retrieving folder structure...")
    
    $StartTime = Get-Date
    $Folders = @()
    
    # Process root folder first
    $RootFolder = Get-Item -Path $FolderPath -ErrorAction Stop
    $Folders += $RootFolder
    
    # Add subfolders with optimized method
    Write-Host "Finding subfolders..." -ForegroundColor DarkGray
    $SubFolders = Get-SubdirectoriesFast -Path $FolderPath -MaxDepth $MaxDepth
    
    # Remove the first item as it's the root folder we already added
    if ($SubFolders.Count -gt 1) {
        $SubFolders = $SubFolders[1..($SubFolders.Count - 1)]
        $Folders += $SubFolders
    }
    
    $TotalFolders = $Folders.Count
    $TimeElapsed = (Get-Date) - $StartTime
    
    Write-Host "Found $TotalFolders folders to process in $($TimeElapsed.TotalSeconds.ToString("0.00")) seconds" -ForegroundColor Cyan
    [void]$OutputText.AppendLine("Found $TotalFolders folders to process")
    
    # Set up runspace pool for parallel processing
    Write-Host "Initializing parallel processing engine..." -ForegroundColor Cyan
    $RunspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads)
    $RunspacePool.Open()
    
    # Create runspaces for each folder
    $Runspaces = @()
    $Counter = 0
    
    # Dictionary to store permissions by folder path
    $FolderPermissionsMap = @{}
    
    # Process folders in batches to prevent memory issues with very large directories
    $BatchSize = [Math]::Min(500, [Math]::Max(50, $TotalFolders / 10))
    $FolderBatches = [System.Collections.Generic.List[Object]]::new()
    
    for ($i = 0; $i -lt $TotalFolders; $i += $BatchSize) {
        $End = [Math]::Min($i + $BatchSize - 1, $TotalFolders - 1)
        $FolderBatches.Add($Folders[$i..$End])
    }
    
    $BatchCounter = 0
    $TotalBatches = $FolderBatches.Count
    
    foreach ($Batch in $FolderBatches) {
        $BatchCounter++
        $BatchFolderCount = $Batch.Count
        
        Write-Host "Processing batch $BatchCounter of $TotalBatches ($BatchFolderCount folders)..." -ForegroundColor Cyan
        
        # Create and invoke runspaces for this batch
        $BatchRunspaces = @()
        
        foreach ($Folder in $Batch) {
            $Counter++
            
            # Create PowerShell instance and add script
            $PowerShell = [powershell]::Create().AddScript({
                param($FolderPath, $GetPermissionsHash)
                
                # Define the worker function inline in each runspace
                function Get-FolderPermissionsWorker {
                    param (
                        [string]$Path
                    )
                    
                    try {
                        $Acl = [System.Security.AccessControl.DirectorySecurity]::new()
                        $Acl.SetAccessRuleProtection($false, $false)
                        
                        $Acl = [System.IO.Directory]::GetAccessControl($Path)
                        $Permissions = @()
                        
                        foreach ($Access in $Acl.Access) {
                            # Create a custom object for each permission entry
                            $Permission = [PSCustomObject]@{
                                FolderPath       = $Path
                                IdentityReference = $Access.IdentityReference
                                FileSystemRights  = $Access.FileSystemRights
                                AccessControlType = $Access.AccessControlType
                                IsInherited       = $Access.IsInherited
                                InheritanceFlags  = $Access.InheritanceFlags
                                PropagationFlags  = $Access.PropagationFlags
                            }
                            
                            $Permissions += $Permission
                        }
                        
                        return @{
                            Success = $true
                            FolderPath = $Path
                            Permissions = $Permissions
                            PermissionsHash = (& $GetPermissionsHash -Permissions $Permissions)
                            Error = $null
                        }
                    }
                    catch {
                        return @{
                            Success = $false
                            FolderPath = $Path
                            Permissions = @()
                            PermissionsHash = 0
                            Error = $_.Exception.Message
                        }
                    }
                }
                
                # Call the worker function with the provided parameters
                return Get-FolderPermissionsWorker -Path $FolderPath
            }).AddArgument($Folder.FullName).AddArgument(${function:Get-PermissionsHash})
            
            $PowerShell.RunspacePool = $RunspacePool
            
            # Save runspace info
            $BatchRunspaces += [PSCustomObject]@{
                Instance = $PowerShell
                Handle = $PowerShell.BeginInvoke()
                FolderPath = $Folder.FullName
                Counter = $Counter
                Total = $TotalFolders
                Completed = $false
            }
        }
        
        # Wait for all runspaces in this batch to complete
        $CompletedCount = 0
        $LastProgressUpdate = Get-Date
        
        while ($BatchRunspaces.Where({-not $_.Completed}, 'First').Count -gt 0) {
            foreach ($Runspace in $BatchRunspaces.Where({-not $_.Completed})) {
                if ($Runspace.Handle.IsCompleted) {
                    $CompletedCount++
                    $Runspace.Completed = $true
                    
                    # Get results from this runspace
                    $Result = $Runspace.Instance.EndInvoke($Runspace.Handle)
                    
                    if ($Result.Success) {
                        # Store the permissions for this folder path
                        $FolderPermissionsMap[$Result.FolderPath] = @{
                            Permissions = $Result.Permissions
                            Hash = $Result.PermissionsHash
                        }
                    }
                    else {
                        Write-Host "Error processing folder: $($Result.FolderPath)" -ForegroundColor Yellow
                        Write-Host "Error details: $($Result.Error)" -ForegroundColor Yellow
                        [void]$OutputText.AppendLine("Error processing folder: $($Result.FolderPath)")
                        [void]$OutputText.AppendLine("Error details: $($Result.Error)")
                    }
                    
                    # Cleanup
                    $Runspace.Instance.Dispose()
                }
            }
            
            # Update progress less frequently to reduce overhead - Fixed parenthesis issue
            if (((Get-Date) - $LastProgressUpdate).TotalMilliseconds -gt 500) {
                $CurrentProgress = "$CompletedCount of $BatchFolderCount folders in current batch"
                $OverallProgress = "Overall: $Counter of $TotalFolders folders ($([Math]::Round($Counter / $TotalFolders * 100))%)"
                Write-Host "`rProcessing: $CurrentProgress | $OverallProgress" -ForegroundColor DarkGray -NoNewline
                $LastProgressUpdate = Get-Date
            }
            
            # Small sleep to prevent CPU hogging
            Start-Sleep -Milliseconds 50
        }
        
        Write-Host "`nBatch $BatchCounter completed." -ForegroundColor Green
    }
    
    # Clean up the runspace pool
    $RunspacePool.Close()
    $RunspacePool.Dispose()
    
    # Compute total results count
    $TotalPermissions = ($FolderPermissionsMap.Values | ForEach-Object { $_.Permissions.Count } | Measure-Object -Sum).Sum
    
    # Display completion message
    $TimeElapsed = (Get-Date) - $StartTime
    Write-Host "Analysis completed in $($TimeElapsed.TotalSeconds.ToString("0.00")) seconds." -ForegroundColor Green
    Write-Host "Found $TotalPermissions permission entries across $TotalFolders folders." -ForegroundColor Green
    [void]$OutputText.AppendLine("Analysis completed. Found $TotalPermissions permission entries across $TotalFolders folders.")
    
    # Display results grouped by folder with separate tables
    Write-Host "`nDisplaying permissions by folder:" -ForegroundColor Cyan
    [void]$OutputText.AppendLine("")
    [void]$OutputText.AppendLine("Displaying permissions by folder:")
    
    # Get all folder paths and sort them by depth (for parent-child relationship checking)
    $SortedFolderPaths = $FolderPermissionsMap.Keys | Sort-Object { ($_ -split '\\').Count }
    
    # Keep track of folders already displayed and build hash table for faster lookups
    $DisplayedFolders = @{}
    $SkippedFolders = @()
    
    # Build hash to permission map for faster comparison
    $HashToFoldersMap = @{}
    foreach ($FolderPath in $SortedFolderPaths) {
        $Hash = $FolderPermissionsMap[$FolderPath].Hash
        if (-not $HashToFoldersMap.ContainsKey($Hash)) {
            $HashToFoldersMap[$Hash] = @()
        }
        $HashToFoldersMap[$Hash] += $FolderPath
    }
    
    Write-Host "Processing folder groups for display..." -ForegroundColor DarkGray
    
    foreach ($FolderPath in $SortedFolderPaths) {
        # Skip if already processed as part of a group
        if ($DisplayedFolders.ContainsKey($FolderPath)) {
            continue
        }
        
        $CurrentHash = $FolderPermissionsMap[$FolderPath].Hash
        $CurrentFolderPermissions = $FolderPermissionsMap[$FolderPath].Permissions
        
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
        
        # Find all child folders with identical permissions - optimized using hash lookup
        $IdenticalSubfolders = @()
        
        # Use hash-based matching first (faster)
        foreach ($OtherPath in $HashToFoldersMap[$CurrentHash]) {
            # Skip self or already displayed
            if (($OtherPath -eq $FolderPath) -or ($DisplayedFolders.ContainsKey($OtherPath))) {
                continue
            }
            
            # Check if it's a subfolder
            if ($OtherPath.StartsWith($FolderPath + "\")) {
                # Verify with full comparison if needed for absolute certainty
                if ($SkipUniquenessCounting -or 
                    (Compare-PermissionSets -Set1 $CurrentFolderPermissions -Set2 $FolderPermissionsMap[$OtherPath].Permissions)) {
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
            
            # For very large lists, summarize instead of showing all
            if ($IdenticalSubfolders.Count -gt 20) {
                Write-Host "  - $($IdenticalSubfolders.Count) identical subfolders" -ForegroundColor DarkGray
                [void]$OutputText.AppendLine("  - $($IdenticalSubfolders.Count) identical subfolders")
                
                # Show first 10 as examples
                foreach ($Subfolder in $IdenticalSubfolders[0..9]) {
                    Write-Host "  - $Subfolder" -ForegroundColor DarkGray
                    [void]$OutputText.AppendLine("  - $Subfolder")
                }
                Write-Host "  - ... (and $($IdenticalSubfolders.Count - 10) more)" -ForegroundColor DarkGray
                [void]$OutputText.AppendLine("  - ... (and $($IdenticalSubfolders.Count - 10) more)")
            }
            else {
                foreach ($Subfolder in $IdenticalSubfolders) {
                    Write-Host "  - $Subfolder" -ForegroundColor DarkGray
                    [void]$OutputText.AppendLine("  - $Subfolder")
                }
            }
        }
        
        # Force garbage collection periodically to reduce memory pressure
        if ($DisplayedFolders.Count % 100 -eq 0) {
            [System.GC]::Collect()
        }
    }
    
    # Report skipped folders
    $SkippedCount = $SkippedFolders.Count
    Write-Host "`nSkipped displaying $SkippedCount folders with permissions identical to their parent folders." -ForegroundColor Cyan
    [void]$OutputText.AppendLine("")
    [void]$OutputText.AppendLine("Skipped displaying $SkippedCount folders with permissions identical to their parent folders.")
    
    # Save the output to text file
    Write-Host "Writing report to file..." -ForegroundColor DarkGray
    $OutputText.ToString() | Out-File -FilePath $OutputTxt -Encoding UTF8
    Write-Host "`nPermissions report exported to: $OutputTxt" -ForegroundColor Green
    
    # Final performance summary
    $TotalTime = (Get-Date) - $StartTime
    Write-Host "`nTotal execution time: $($TotalTime.TotalSeconds.ToString("0.00")) seconds" -ForegroundColor Green
    Write-Host "Processed $TotalFolders folders ($($TotalTime.TotalSeconds / $TotalFolders * 1000 -as [int]) ms per folder)" -ForegroundColor Green
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
