# =============================================================================
# Script: Get-NTFSFolderPermissions.ps1
# Created: 5-03-06 21:06:43 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-06 22:43:15 UTC
# Updated By: jdyer-nuvodia
# Version: 1.4.0
# Additional Info: Added ultra-restricted environment compatibility mode
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

# EXTREME RESTRICTED MODE COMPATIBILITY
# Use only .NET core classes that are available in any PowerShell environment
# This mode uses only basic .NET operations and avoids PowerShell-specific features

# Create constants using only .NET methods (avoid Get-Date cmdlet)
$dateTimeNow = [DateTime]::Now
$formattedDateTime = $dateTimeNow.ToString("yyyy-MM-dd_HHmmss")
$consoleErrorColor = [ConsoleColor]::Red
$consoleNormalColor = [ConsoleColor]::White
$consoleSuccessColor = [ConsoleColor]::Green
$consoleWarnColor = [ConsoleColor]::Yellow
$consoleInfoColor = [ConsoleColor]::Cyan
$consoleTechColor = [ConsoleColor]::DarkGray

# Get script directory using .NET methods rather than PowerShell cmdlets
$scriptDirectory = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Path)
$outputLogPath = [System.IO.Path]::Combine($scriptDirectory, "NTFSPermissions_$formattedDateTime.log")

# Create an output log file using pure .NET methods
try {
    $fileStream = [System.IO.FileStream]::new($outputLogPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
    $streamWriter = [System.IO.StreamWriter]::new($fileStream, [System.Text.Encoding]::UTF8)
    
    # Write basic header info 
    $streamWriter.WriteLine("NTFS Permissions Report - Generated $dateTimeNow")
    $streamWriter.WriteLine("Folder Path: $FolderPath")
    $streamWriter.WriteLine([string]::new("=", 80))
    $streamWriter.WriteLine("")
    $streamWriter.WriteLine("Starting NTFS permissions analysis for: $FolderPath")
    
    # Keep the stream writer open for the rest of the script
} 
catch {
    [Console]::ForegroundColor = $consoleErrorColor
    [Console]::WriteLine("ERROR: Failed to create log file: $_")
    [Console]::ResetColor()
}

# Direct console output functions - inline to avoid scope issues
try {
    # Hello message
    [Console]::ForegroundColor = $consoleInfoColor
    [Console]::WriteLine("Starting optimized NTFS permissions analysis for: $FolderPath")
    [Console]::ForegroundColor = $consoleTechColor
    [Console]::WriteLine("Using up to $MaxThreads parallel threads")
    if ($MaxDepth -gt 0) {
        [Console]::WriteLine("Limited to maximum depth of $MaxDepth levels")
    }
    [Console]::ResetColor()

    # Use System.IO methods to get directory info
    $rootDir = [System.IO.DirectoryInfo]::new($FolderPath)
    
    # Record start time using .NET DateTime
    $startTime = [DateTime]::Now

    # Function to get all subdirectories using pure .NET (no cmdlets)
    function Get-AllDirectories {
        param(
            [string]$rootPath,
            [int]$maxDepth = 0,
            [int]$currentDepth = 0
        )
        
        try {
            # Start with root folder

# Display start message with optimization info
Write-SafeOutput "Starting optimized NTFS permissions analysis for: $FolderPath" -ForegroundColor Cyan
Write-SafeOutput "Using up to $MaxThreads parallel threads" -ForegroundColor DarkGray
if ($MaxDepth -gt 0) {
    Write-SafeOutput "Limited to maximum depth of $MaxDepth levels" -ForegroundColor DarkGray
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

# Safe direct output function that works with no PowerShell cmdlets
function Write-Output-Safe {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [string]$Color = "White",
        [switch]$NoNewline
    )
    
    try {
        # Try direct .NET console output (works everywhere)
        if ($NoNewline) {
            [Console]::Write($Message)
        } else {
            [Console]::WriteLine($Message)
        }
    }
    catch {
        # If even Console fails, we can't output anything
    }
}

# Direct file writing function that doesn't rely on Out-File
function Write-File-Safe {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Content,
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        [string]$Encoding = "UTF8"
    )
    
    try {
        # Use System.IO.File directly
        [System.IO.File]::WriteAllText($FilePath, $Content, [System.Text.Encoding]::UTF8)
        return $true
    }
    catch {
        try {
            # Try alternate method
            $stream = New-Object System.IO.StreamWriter($FilePath, $false, [System.Text.Encoding]::UTF8)
            $stream.Write($Content)
            $stream.Close()
            return $true
        }
        catch {
            # If even direct .NET methods fail, we can't write to file
            return $false
        }
    }
}

try {
    # Get all folders and subfolders recursively using optimized method
    Write-SafeOutput "Retrieving folder structure (optimized method)..." -ForegroundColor Cyan
    [void]$OutputText.AppendLine("Retrieving folder structure...")
    
    $StartTime = Get-Date
    $Folders = @()
    
    # Process root folder first
    $RootFolder = Get-Item -Path $FolderPath -ErrorAction Stop
    $Folders += $RootFolder
    
    # Add subfolders with optimized method
    Write-SafeOutput "Finding subfolders..." -ForegroundColor DarkGray
    $SubFolders = Get-SubdirectoriesFast -Path $FolderPath -MaxDepth $MaxDepth
    
    # Remove the first item as it's the root folder we already added
    if ($SubFolders.Count -gt 1) {
        $SubFolders = $SubFolders[1..($SubFolders.Count - 1)]
        $Folders += $SubFolders
    }
    
    $TotalFolders = $Folders.Count
    $TimeElapsed = (Get-Date) - $StartTime
    
    Write-SafeOutput "Found $TotalFolders folders to process in $($TimeElapsed.TotalSeconds.ToString("0.00")) seconds" -ForegroundColor Cyan
    [void]$OutputText.AppendLine("Found $TotalFolders folders to process")
    
    # Set up runspace pool for parallel processing
    Write-SafeOutput "Initializing parallel processing engine..." -ForegroundColor Cyan
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
        
        Write-SafeOutput "Processing batch $BatchCounter of $TotalBatches ($BatchFolderCount folders)..." -ForegroundColor Cyan
        
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
                        # More robust method to get ACL - using Get-Acl instead of DirectoryInfo.GetAccessControl
                        # This is more reliable across different environments and directory types
                        $Acl = Get-Acl -Path $Path -ErrorAction Stop
                        
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
                        Write-SafeOutput "Error processing folder: $($Result.FolderPath)" -ForegroundColor Yellow
                        Write-SafeOutput "Error details: $($Result.Error)" -ForegroundColor Yellow
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
                try {
                    Write-Host "`rProcessing: $CurrentProgress | $OverallProgress" -ForegroundColor DarkGray -NoNewline
                } catch {
                    # Fallback for progress updates if Write-Host fails
                    if ($CompletedCount % 10 -eq 0) {  # Only show periodically to avoid console spam
                        [Console]::WriteLine("Processing: $CurrentProgress | $OverallProgress")
                    }
                }
                $LastProgressUpdate = Get-Date
            }
            
            # Small sleep to prevent CPU hogging
            Start-Sleep -Milliseconds 50
        }
        
        Write-SafeOutput "`nBatch $BatchCounter completed." -ForegroundColor Green
    }
    
    # Clean up the runspace pool
    $RunspacePool.Close()
    $RunspacePool.Dispose()
    
    # Compute total results count
    $TotalPermissions = ($FolderPermissionsMap.Values | ForEach-Object { $_.Permissions.Count } | Measure-Object -Sum).Sum
    
    # Display completion message
    $TimeElapsed = (Get-Date) - $StartTime
    Write-SafeOutput "Analysis completed in $($TimeElapsed.TotalSeconds.ToString("0.00")) seconds." -ForegroundColor Green
    Write-SafeOutput "Found $TotalPermissions permission entries across $TotalFolders folders." -ForegroundColor Green
    [void]$OutputText.AppendLine("Analysis completed. Found $TotalPermissions permission entries across $TotalFolders folders.")
    
    # Display results grouped by folder with separate tables
    Write-SafeOutput "`nDisplaying permissions by folder:" -ForegroundColor Cyan
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
    
    Write-SafeOutput "Processing folder groups for display..." -ForegroundColor DarkGray
    
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
        
        Write-SafeOutput "`n$Separator" -ForegroundColor White
        Write-SafeOutput "Folder: $FolderPath" -ForegroundColor White
        Write-SafeOutput "$Separator" -ForegroundColor White
        
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
        
        Write-SafeOutput $PermissionsTable
        [void]$OutputText.Append($PermissionsTable)
        
        # Mark this folder as displayed
        $DisplayedFolders[$FolderPath] = $true
        
        # If there are subfolders with identical permissions, list them
        if ($IdenticalSubfolders.Count -gt 0) {
            Write-SafeOutput "The following subfolders have identical permissions:" -ForegroundColor Cyan
            [void]$OutputText.AppendLine("The following subfolders have identical permissions:")
            
            # For very large lists, summarize instead of showing all
            if ($IdenticalSubfolders.Count -gt 20) {
                Write-SafeOutput "  - $($IdenticalSubfolders.Count) identical subfolders" -ForegroundColor DarkGray
                [void]$OutputText.AppendLine("  - $($IdenticalSubfolders.Count) identical subfolders")
                
                # Show first 10 as examples
                foreach ($Subfolder in $IdenticalSubfolders[0..9]) {
                    Write-SafeOutput "  - $Subfolder" -ForegroundColor DarkGray
                    [void]$OutputText.AppendLine("  - $Subfolder")
                }
                Write-SafeOutput "  - ... (and $($IdenticalSubfolders.Count - 10) more)" -ForegroundColor DarkGray
                [void]$OutputText.AppendLine("  - ... (and $($IdenticalSubfolders.Count - 10) more)")
            }
            else {
                foreach ($Subfolder in $IdenticalSubfolders) {
                    Write-SafeOutput "  - $Subfolder" -ForegroundColor DarkGray
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
    Write-SafeOutput "`nSkipped displaying $SkippedCount folders with permissions identical to their parent folders." -ForegroundColor Cyan
    [void]$OutputText.AppendLine("")
    [void]$OutputText.AppendLine("Skipped displaying $SkippedCount folders with permissions identical to their parent folders.")
    
    # Save the output to text file
    Write-SafeOutput "Writing report to file..." -ForegroundColor DarkGray
    $OutputText.ToString() | Out-File -FilePath $OutputLog -Encoding UTF8
    Write-SafeOutput "`nPermissions report exported to: $OutputLog" -ForegroundColor Green
    
    # Final performance summary
    $TotalTime = (Get-Date) - $StartTime
    Write-SafeOutput "`nTotal execution time: $($TotalTime.TotalSeconds.ToString("0.00")) seconds" -ForegroundColor Green
    Write-SafeOutput "Processed $TotalFolders folders ($($TotalTime.TotalSeconds / $TotalFolders * 1000 -as [int]) ms per folder)" -ForegroundColor Green
}
catch {
    # Super failsafe error handling with no dependencies on PowerShell cmdlets
    try {
        [Console]::Error.WriteLine("An error occurred during the NTFS permissions analysis:")
        [Console]::Error.WriteLine($_.Exception.Message)
    }
    catch {
        # We tried our best, but even Console output failed
    }
    
    [void]$OutputText.AppendLine("An error occurred during the NTFS permissions analysis:")
    [void]$OutputText.AppendLine($_.Exception.Message)
    
    # Try to save what we have so far using our safe file writing function
    if ($OutputText.Length -gt 0) {
        Write-File-Safe -Content $OutputText.ToString() -FilePath $OutputLog
    }
}
