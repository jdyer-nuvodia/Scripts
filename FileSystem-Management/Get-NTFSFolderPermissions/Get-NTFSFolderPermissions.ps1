# =============================================================================
# Script: Get-NTFSFolderPermissions.ps1
# Created: 2025-03-06 21:06:43 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-07 15:32:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.4.14
# Additional Info: Fixed GetAccessControl method by using proper .NET API
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

$StartTime = [DateTime]::Now

# EXTREME RESTRICTED MODE COMPATIBILITY
# Use only .NET core classes that are available in any PowerShell environment
# This mode uses only basic .NET operations and avoids PowerShell-specific features

# Initialize StringBuilder for collecting output text
$OutputText = [System.Text.StringBuilder]::new()

# Create constants using only .NET methods (avoid Get-Date cmdlet)
$dateTimeNow = [DateTime]::Now
$formattedDateTime = $dateTimeNow.ToString("yyyy-MM-dd_HHmmss")
$consoleErrorColor = [ConsoleColor]::Red
# Remove unused color variables and simplify
$consoleTechColor = [ConsoleColor]::DarkGray
# Added missing definition for console info color
$consoleInfoColor = [ConsoleColor]::Cyan

# Get script directory using .NET methods rather than PowerShell cmdlets
$scriptDirectory = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Path)
$outputLogPath = [System.IO.Path]::Combine($scriptDirectory, "NTFSPermissions_$formattedDateTime.log")
$OutputLog = $outputLogPath # Ensure OutputLog is set to the proper path

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

# Safe output function with PowerShell Write-Host fallback (MUST BE DEFINED BEFORE FIRST USE)
function global:Write-SafeOutput {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [string]$Color = "White",
        [switch]$NoNewline
    )
    
    try {
        # First try Write-Host for normal PowerShell environments
        if ($NoNewline) {
            Write-Host $Message -ForegroundColor $Color -NoNewline
        } else {
            Write-Host $Message -ForegroundColor $Color
        }
        
        # Also capture to our OutputText for logging
        if ($null -ne $OutputText) {
            if (!$NoNewline) {
                [void]$OutputText.AppendLine($Message)
            } else {
                [void]$OutputText.Append($Message)
            }
        }
    }
    catch {
        # Try direct .NET console output (works everywhere)
        try {
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
}

# Create an output log file using pure .NET methods
try {
    $fileStream = [System.IO.FileStream]::new($OutputLog, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
    $streamWriter = [System.IO.StreamWriter]::new($fileStream, [System.Text.Encoding]::UTF8)
    
    # Write basic header info 
    $streamWriter.WriteLine("NTFS Permissions Report - Generated $dateTimeNow")
    $streamWriter.WriteLine("Folder Path: $FolderPath")
    $streamWriter.WriteLine([string]::new("=", 80))
    $streamWriter.WriteLine("")
    $streamWriter.WriteLine("Starting NTFS permissions analysis for: $FolderPath")
    
    # Close the stream so the file is not locked later
    $streamWriter.Flush()
    $streamWriter.Close()
    $fileStream.Close()
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
}
catch {
    # If console output fails, continue silently
    # In restricted environments, console methods might not be available
}

# New helper functions for modularization

function Get-AllDirectoriesModule {
    param(
        [string]$Path,
        [int]$MaxDepth = 0
    )
    try {
        $dirInfo = [System.IO.DirectoryInfo]::new($Path)
        if ($MaxDepth -gt 0) {
            return Get-AllDirectoriesModuleRecursive -dir $dirInfo -maxDepth $MaxDepth -current 0
        }
        else {
            return $dirInfo.GetDirectories("*", 'AllDirectories')
        }
    }
    catch {
        return @()
    }
}

function Get-AllDirectoriesModuleRecursive {
    param(
        [System.IO.DirectoryInfo]$dir,
        [int]$maxDepth,
        [int]$current
    )
    $results = @($dir)
    if ($current -lt $maxDepth) {
        try {
            foreach ($sub in $dir.GetDirectories()) {
                $results += Get-AllDirectoriesModuleRecursive -dir $sub -maxDepth $maxDepth -current ($current+1)
            }
        }
        catch { }
    }
    return $results
}

# Define Write-Log function first so it's available to other functions
function Write-Log {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    try {
        Write-Host "[DEBUG] $Message" -ForegroundColor $Color
    }
    catch {
        [Console]::WriteLine("[DEBUG] " + $Message)
    }
    if ($null -ne $OutputText) {
        [void]$OutputText.AppendLine($Message)
    }
}

# Updated Get-FolderPermissionsModule with fixed string handling
function Get-FolderPermissionsModule {
    param([string]$FolderPath)
    
    try {
        # Use Get-Acl with provider path instead of DirectoryInfo
        $acl = Get-Acl -LiteralPath $FolderPath
        if ($null -eq $acl) {
            throw "Unable to get ACL"
        }
        
        $permissions = foreach ($access in $acl.Access) {
            [PSCustomObject]@{
                IdentityReference = $access.IdentityReference
                FileSystemRights = $access.FileSystemRights
                AccessControlType = $access.AccessControlType
                IsInherited = $access.IsInherited
                InheritanceFlags = $access.InheritanceFlags
                PropagationFlags = $access.PropagationFlags
            }
        }
        
        if ($permissions.Count -eq 0) {
            $permissions = @([PSCustomObject]@{
                FolderPath = $FolderPath
                IdentityReference = "LIMITED ACCESS"
                FileSystemRights = "Unknown"
                AccessControlType = "Unknown"
                IsInherited = $true
                InheritanceFlags = "None"
                PropagationFlags = "None"
            })
        }
        
        $sorted = ($permissions | Sort-Object IdentityReference, FileSystemRights, AccessControlType, IsInherited |
                   ForEach-Object { "$($_.IdentityReference)|$($_.FileSystemRights)|$($_.AccessControlType)|$($_.IsInherited)" }) -join ";"
        
        return @{ Success = $true; FolderPath = $FolderPath; Permissions = $permissions; Hash = $sorted.GetHashCode() }
    }
    catch {
        Write-Log "Error accessing permissions for $FolderPath : $($_.Exception.Message)" "Yellow"
        return @{
            Success = $false
            FolderPath = $FolderPath
            Error = $_.Exception.Message
            Permissions = @([PSCustomObject]@{
                IdentityReference = "ERROR"
                FileSystemRights = "Access Denied: $($_.Exception.Message)"
                AccessControlType = "N/A"
                IsInherited = $true
                InheritanceFlags = "N/A"
                PropagationFlags = "N/A"
            })
            Hash = -1
        }
    }
}

function Start-FolderProcessing {
    param(
        [array]$Folders,
        [int]$MaxThreads,
        [switch]$SkipUniquenessCounting
    )
    
    $RunspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads)
    $RunspacePool.Open()
    $FolderPermissionsMap = @{}
    $Runspaces = @()
    
    # Get function definitions as strings
    $GetFolderPermissionsModuleDefinition = ${function:Get-FolderPermissionsModule}.ToString()
    $WriteLogDefinition = ${function:Write-Log}.ToString()

    foreach ($folder in $Folders) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $RunspacePool

        [void]$ps.AddScript({
            param($FolderPath, $GetFolderPermissionsModuleDef, $WriteLogDef, $OutputText)
            
            # Define required functions in runspace
            ${function:Write-Log} = $WriteLogDef
            ${function:Get-FolderPermissionsModule} = $GetFolderPermissionsModuleDef
            
            try {
                return Get-FolderPermissionsModule -FolderPath $FolderPath
            }
            catch {
                return @{
                    Success = $false
                    FolderPath = $FolderPath
                    Error = $_.Exception.Message
                    FullError = $_ | Out-String
                }
            }
        }).AddArgument($folder.FullName).AddArgument($GetFolderPermissionsModuleDefinition).AddArgument($WriteLogDefinition).AddArgument($OutputText)
        
        $Runspaces += [PSCustomObject]@{
            Instance = $ps
            Handle = $ps.BeginInvoke()
            Folder = $folder.FullName
        }
    }
    
    foreach ($r in $Runspaces) {
        try {
            $result = $r.Instance.EndInvoke($r.Handle)
            if ($result.Success) {
                $FolderPermissionsMap[$result.FolderPath] = @{ 
                    Permissions = $result.Permissions
                    Hash = $result.Hash 
                }
                Write-Log "Successfully processed folder: $($result.FolderPath)" "Cyan"
            }
            else {
                Write-Log "Error processing folder: $($r.Folder) - Error details: $($result.Error)" "Yellow"
                if ($result.FullError) {
                    Write-Log "Full error: $($result.FullError)" "Yellow"
                }
            }
        }
        catch {
            Write-Log "Critical error in runspace for folder $($r.Folder): $($_.Exception.Message)" "Red"
        }
        finally {
            $r.Instance.Dispose()
        }
    }
    
    $RunspacePool.Close()
    $RunspacePool.Dispose()
    return $FolderPermissionsMap
}

# --- Main processing block rewritten using new modular functions ---
try {
    # Enumerate folders using the new module function
    $Folders = Get-AllDirectoriesModule -Path $FolderPath -MaxDepth $MaxDepth
    $Folders = $Folders | Sort-Object FullName -Unique
    Write-Log "Found $($Folders.Count) folders to process" "Cyan"
    $TotalFolders = $Folders.Count

    # Process folders asynchronously and collect permissions
    $FolderPermissionsMap = Start-FolderProcessing -Folders $Folders -MaxThreads $MaxThreads -SkipUniquenessCounting:$SkipUniquenessCounting

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

    # Display results grouped by folder with separate tables
    Write-SafeOutput "`nDisplaying permissions by folder:" -Color Cyan
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

    Write-SafeOutput "Processing folder groups for display..." -Color DarkGray

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
        
        Write-SafeOutput "`n$Separator" -Color White
        Write-SafeOutput "Folder: $FolderPath" -Color White
        Write-SafeOutput "$Separator" -Color White
        
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
            Write-SafeOutput "The following subfolders have identical permissions:" -Color Cyan
            [void]$OutputText.AppendLine("The following subfolders have identical permissions:")
            
            # For very large lists, summarize instead of showing all
            if ($IdenticalSubfolders.Count -gt 20) {
                Write-SafeOutput "  - $($IdenticalSubfolders.Count) identical subfolders" -Color DarkGray
                [void]$OutputText.AppendLine("  - $($IdenticalSubfolders.Count) identical subfolders")
                
                # Show first 10 as examples
                foreach ($Subfolder in $IdenticalSubfolders[0..9]) {
                    Write-SafeOutput "  - $Subfolder" -Color DarkGray
                    [void]$OutputText.AppendLine("  - $Subfolder")
                }
                Write-SafeOutput "  - ... (and $($IdenticalSubfolders.Count - 10) more)" -Color DarkGray
                [void]$OutputText.AppendLine("  - ... (and $($IdenticalSubfolders.Count - 10) more)")
            }
            else {
                foreach ($Subfolder in $IdenticalSubfolders) {
                    Write-SafeOutput "  - $Subfolder" -Color DarkGray
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
    Write-SafeOutput "`nSkipped displaying $SkippedCount folders with permissions identical to their parent folders." -Color Cyan
    [void]$OutputText.AppendLine("")
    [void]$OutputText.AppendLine("Skipped displaying $SkippedCount folders with permissions identical to their parent folders.")

    # Save the output to text file
    try {
        Write-SafeOutput "Writing report to file..." -Color DarkGray
        Write-File-Safe -Content $OutputText.ToString() -FilePath $OutputLog -Encoding "UTF8"
        Write-SafeOutput "`nPermissions report exported to: $OutputLog" -Color Green

        $TotalTime = ([DateTime]::Now) - $StartTime
        Write-SafeOutput "`nTotal execution time: $($TotalTime.TotalSeconds.ToString('0.00')) seconds" -Color Green
        Write-SafeOutput "Processed $TotalFolders folders ($([int]($TotalTime.TotalSeconds / $TotalFolders * 1000)) ms per folder)" -Color Green
    }
    catch {
        Write-SafeOutput "An error occurred during log writing or performance summary: $($_.Exception.Message)" -Color Red
    }
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
    if ($null -ne $OutputText -and $OutputText.Length -gt 0 -and $null -ne $OutputLog -and $OutputLog -ne "") {
        try {
            [System.IO.File]::WriteAllText($OutputLog, $OutputText.ToString(), [System.Text.Encoding]::UTF8)
        }
        catch {
            try {
                # Last-ditch effort to write error to desktop
                $desktopPath = [Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)
                $fallbackLog = [System.IO.Path]::Combine($desktopPath, "NTFSPermissions_ERROR_$formattedDateTime.log")
                [System.IO.File]::WriteAllText($fallbackLog, "Critical error in NTFS permissions script: $($_.Exception.Message)")
            } catch {
                # We've tried everything possible
            }
        }
    } else {
        try {
            # Last-ditch effort to write error to desktop
            $desktopPath = [Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)
            $fallbackLog = [System.IO.Path]::Combine($desktopPath, "NTFSPermissions_ERROR_$formattedDateTime.log")
            [System.IO.File]::WriteAllText($fallbackLog, "Critical error in NTFS permissions script: $($_.Exception.Message)")
        } catch {
            # We've tried everything possible
        }
    }
}
