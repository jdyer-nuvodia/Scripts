# =============================================================================
# Script: Get-NTFSFolderPermissions.ps1
# Created: 2025-03-06 21:06:43 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-11 17:11:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.11.13
# Additional Info: Enhanced AD module import handling for non-domain environments
# =============================================================================

<#
.SYNOPSIS
    Extracts and reports NTFS permissions for specified folders with optimized performance.
.DESCRIPTION
    This script retrieves NTFS permissions for a specified folder path and all its subfolders.
    Key features:
    - Uses optimized directory traversal methods for improved performance
    - Processes folders in parallel with configurable thread limits
    - Forces Active Directory module loading for SID resolution
    - Supports SID resolution on non-domain controller systems
    - Groups folders with identical permissions to reduce output clutter
    - Exports results to a formatted log file
    
    Dependencies:
    - Windows PowerShell 5.1 or later
    - RSAT AD PowerShell module (auto-installed if missing)
    - Read access to target folders
.PARAMETER FolderPath
    The path to the folder for which permissions will be extracted.
    Example: "C:\Important\Data" or "\\server\share\folder"
.PARAMETER MaxThreads
    Maximum number of parallel threads to use for processing.
    Default: 10
.PARAMETER MaxDepth
    Maximum folder depth to traverse. Set to 0 for unlimited depth.
    Default: 0
.PARAMETER SkipUniquenessCounting
    Skip counting unique permissions for large directories to improve performance.
    Default: False
.PARAMETER SkipADResolution
    Skip Active Directory SID resolution to avoid AD module dependency.
    Default: False
.EXAMPLE
    .\Get-NTFSFolderPermissions.ps1 -FolderPath "C:\Important\Data"
    Retrieves NTFS permissions for C:\Important\Data and all subfolders
.EXAMPLE
    .\Get-NTFSFolderPermissions.ps1 -FolderPath "\\server\share\folder" -MaxThreads 20
    Uses 20 parallel threads to process folders on a network share
.EXAMPLE
    .\Get-NTFSFolderPermissions.ps1 -FolderPath "C:\VeryLargeFolder" -MaxDepth 3
    Processes only folders up to 3 levels deep from the root
.EXAMPLE
    .\Get-NTFSFolderPermissions.ps1 -FolderPath "C:\Data" -SkipADResolution
    Processes permissions without attempting to resolve SIDs through Active Directory
#>

using namespace System.Security.AccessControl
using namespace System.IO
using namespace System.Security.Principal

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
    [switch]$SkipUniquenessCounting,

    [Parameter(Mandatory = $false)]
    [switch]$SkipADResolution
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
$systemName = [System.Environment]::MachineName
$outputLogPath = [System.IO.Path]::Combine($scriptDirectory, "NTFSPermissions_${systemName}_$formattedDateTime.log")
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
    
    # Check if we're in verbose mode
    $isVerboseMode = $VerbosePreference -eq 'Continue' -or $DebugPreference -eq 'Continue'
    
    # Skip progress messages in non-verbose mode
    if (!$isVerboseMode -and $Message -match 'Successfully processed') {
        return
    }
    
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
        [string]$Color = "White",
        [switch]$NoNewline
    )
    
    # Always append to log file
    [void]$OutputText.AppendLine($Message)
    
    # Check if we're in verbose mode
    $isVerboseMode = $VerbosePreference -eq 'Continue' -or $DebugPreference -eq 'Continue'
    
    # Only write debug/progress messages to console in verbose mode
    if ($Message -match '^\[DEBUG\]' -or $Message -match 'Successfully processed') {
        if ($isVerboseMode) {
            Write-Host $Message -ForegroundColor $Color -NoNewline:$NoNewline
        }
        return
    }
    
    # Write non-debug messages normally
    Write-Host $Message -ForegroundColor $Color -NoNewline:$NoNewline
}

function Write-ProgressBar {
    param(
        [int]$Current,
        [int]$Total,
        [string]$Activity = "Processing Folders",
        [string]$Status = "Current Progress"
    )
    
    if ($VerbosePreference -ne 'Continue') {
        $percentComplete = [math]::Round(($Current / $Total) * 100, 1)
        $progressParams = @{
            Activity = $Activity
            Status = $Status
            PercentComplete = $percentComplete
            CurrentOperation = "Folder $Current of $Total"
        }
        Write-Progress @progressParams
    }
}

function Invoke-FolderProcessing {
    param(
        [string]$Path,
        [int]$CurrentCount,
        [int]$TotalCount
    )
    
    try {
        # Get permissions for the folder
        $permissions = Get-FolderPermissions -FolderPath $Path
        
        # Update progress
        Write-ProgressBar -Current $CurrentCount -Total $TotalCount -Status "Processing: $Path"
        
        # Log success in debug mode
        Write-Log "[DEBUG] Successfully processed folder: $Path" -Color "Magenta"
        
        return $permissions
    }
    catch {
        Write-Log "Error processing folder: $Path - $($_.Exception.Message)" -Color "Red"
        return $null
    }
}

# Updated Get-FolderPermissionsModule to use icacls
function Get-FolderPermissionsModule {
    param([string]$FolderPath)
    
    try {
        # Use icacls to get permissions
        $icaclsOutput = & icacls $FolderPath
        
        if ($LASTEXITCODE -ne 0) {
            throw "icacls command failed with exit code $LASTEXITCODE"
        }
        
        # Parse icacls output into permission objects
        $permissions = @()
        foreach ($line in $icaclsOutput | Select-Object -Skip 1) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            
            # Parse each permission line
            if ($line -match '^\s*(.+?):\((.*?)\)') {
                $identity = $matches[1].Trim()
                $rights = $matches[2]
                
                # Convert ICACLS flags to full descriptions
                $expandedRights = @()
                $isInherited = $false
                
                # Split multiple permission flags and process each one
                $rightParts = $rights -split '\)\('
                foreach($part in $rightParts) {
                    $part = $part.Trim('()')
                    
                    # Check if permission is explicitly marked as inherited
                    # Only mark as inherited if the 'I' flag is present without CI/OI flags
                    if ($part -eq 'I') {
                        $isInherited = $true
                    }
                    
                    # Map basic permissions
                    $basicPermission = switch -Regex ($part) {
                        '^F$' { 'Full control' }
                        '^M$' { 'Modify' }
                        '^RX$' { 'Read & Execute' }
                        '^R$' { 'Read' }
                        '^W$' { 'Write' }
                        '^D$' { 'Delete' }
                        default { $null }
                    }
                    
                    if ($basicPermission) {
                        $expandedRights += $basicPermission
                        continue
                    }
                    
                    # Handle inheritance and propagation flags
                    switch -Regex ($part) {
                        'OI' { $expandedRights += 'Object inherit' }
                        'CI' { $expandedRights += 'Container inherit' }
                        'IO' { $expandedRights += 'Inherit only' }
                        'NP' { $expandedRights += 'Do not propagate' }
                        'I[^O]' { 
                            # Only add inherited text if it's a pure inheritance flag
                            if ($part -eq 'I') {
                                $expandedRights += 'Inherited'
                                $isInherited = $true
                            }
                        }
                    }
                }
                
                # If there are inheritance flags (CI/OI) but no 'I' flag, it's not inherited
                if ($expandedRights -match '(Container inherit|Object inherit)' -and -not ($expandedRights -contains 'Inherited')) {
                    $isInherited = $false
                }
                
                # Join all expanded rights with commas for display
                $rightsDescription = ($expandedRights | Where-Object { $_ }) -join ', '
                
                $permissions += [PSCustomObject]@{
                    IdentityReference = $identity
                    FileSystemRights = $rightsDescription
                    AccessControlType = if ($rights -match '\bDENY\b') { 'Deny' } else { 'Allow' }
                    IsInherited = $isInherited
                }
            }
        }
        
        if ($permissions.Count -eq 0) {
            Write-Log "No permissions found for $FolderPath" "Yellow"
            $permissions = @([PSCustomObject]@{
                IdentityReference = "NO ACCESS"
                FileSystemRights = "N/A"
                AccessControlType = "N/A"
                IsInherited = $true
                InheritanceFlags = "None"
                PropagationFlags = "None"
            })
        }
        
        # Generate hash of permissions for comparison
        $sorted = ($permissions | Sort-Object IdentityReference, FileSystemRights, AccessControlType, IsInherited |
                  ForEach-Object { "$($_.IdentityReference)|$($_.FileSystemRights)|$($_.AccessControlType)|$($_.IsInherited)" }) -join ";"
        
        return @{
            Success = $true
            FolderPath = $FolderPath
            Permissions = $permissions
            Hash = $sorted.GetHashCode()
        }
    }
    catch {
        Write-Log "Error accessing permissions for $FolderPath : $($_.Exception.Message)" "Yellow"
        return @{
            Success = $false
            FolderPath = $FolderPath
            Error = $_.Exception.Message
            Permissions = @([PSCustomObject]@{
                IdentityReference = "ACCESS ERROR"
                FileSystemRights = "Access Error: $($_.Exception.Message)"
                AccessControlType = "N/A"
                IsInherited = $true
                InheritanceFlags = "N/A"
                PropagationFlags = "N/A"
            })
            Hash = -1
        }
    }
}

# Function must be defined before first usage
function Initialize-ADModule {
    if ($SkipADResolution) {
        return $false
    }

    try {
        # Try to load module directly first
        if (!(Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue)) {
            if (Get-Module -ListAvailable -Name ActiveDirectory -ErrorAction SilentlyContinue) {
                Import-Module ActiveDirectory -ErrorAction Stop
                Write-Log "Successfully loaded ActiveDirectory module" "Green"
            } else {
                # Module not available, try to install it
                Write-Log "ActiveDirectory module not found. Attempting installation..." "Yellow"
                
                # Try installing via Windows Optional Feature
                try {
                    $feature = Get-WindowsOptionalFeature -Online -FeatureName "Rsat.ActiveDirectory.DS-LDS.Tools" -ErrorAction Stop
                    if ($feature) {
                        if ($feature.State -ne "Enabled") {
                            Write-Log "Installing RSAT AD tools..." "Cyan"
                            Enable-WindowsOptionalFeature -Online -FeatureName "Rsat.ActiveDirectory.DS-LDS.Tools" -NoRestart -ErrorAction Stop
                            Write-Log "RSAT AD tools installation completed. Importing module..." "Green"
                            Import-Module ActiveDirectory -ErrorAction Stop
                        }
                    } else {
                        # Try alternative installation method for Windows 10/11
                        Write-Log "Attempting alternate RSAT installation method..." "Yellow"
                        Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -ErrorAction Stop
                        Import-Module ActiveDirectory -ErrorAction Stop
                    }
                } catch {
                    Write-Log "Could not install RSAT tools: $($_.Exception.Message)" "Yellow"
                    $Global:UseFallbackSIDResolution = $true
                    return $false
                }
            }
        }

        # Verify module loaded successfully
        if (Get-Module -Name ActiveDirectory) {
            # Test AD functionality
            try {
                $null = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
                Write-Log "Successfully verified AD domain connectivity" "Green"
                return $true
            } catch {
                Write-Log "Not domain-joined - using alternate SID resolution" "Yellow"
                $Global:UseFallbackSIDResolution = $true
                return $false
            }
        } else {
            Write-Log "Failed to load ActiveDirectory module" "Yellow"
            $Global:UseFallbackSIDResolution = $true
            return $false
        }
    } catch {
        Write-Log "AD module initialization error: $($_.Exception.Message)" "Yellow"
        $Global:UseFallbackSIDResolution = $true
        return $false
    }
}

# Initialize AD module at startup with retry logic
$maxRetries = 3
$retryCount = 0
$Global:ADModuleAvailable = $false

while (-not $Global:ADModuleAvailable -and $retryCount -lt $maxRetries) {
    $Global:ADModuleAvailable = Initialize-ADModule
    if (-not $Global:ADModuleAvailable) {
        $retryCount++
        if ($retryCount -lt $maxRetries) {
            Write-Log ("Retry $retryCount of $maxRetries : Attempting to load AD module again...") "Yellow"
            Start-Sleep -Seconds 2
        }
    }
}

if (-not $Global:ADModuleAvailable -and -not $SkipADResolution) {
    Write-Log "Warning: Active Directory module could not be loaded after $maxRetries attempts. SID resolution may be limited." -Color "Yellow"
    Write-Log "Use -SkipADResolution to suppress this warning." -Color "Yellow"
}

# Replace the existing Resolve-ADAccountFromSID function with this updated version
function Resolve-ADAccountFromSID {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SID,
        [string]$DomainController
    )
    
    try {
        # Skip if SID resolution is disabled
        if ($SkipADResolution) {
            return $SID
        }

        # Skip if not a valid SID format
        if (-not ($SID -match '^S-\d-\d+(-\d+)+$')) {
            return $SID
        }

        # Check for static/well-known SIDs first
        $wellKnownSIDs = @{
            'S-1-0'='Null Authority'
            'S-1-0-0'='Nobody'
            'S-1-1'='World Authority'
            'S-1-1-0'='Everyone'
            'S-1-2'='Local Authority'
            'S-1-2-0'='Local'
            'S-1-2-1'='Console Logon'
            'S-1-3'='Creator Authority'
            'S-1-3-0'='Creator Owner'
            'S-1-3-1'='Creator Group'
            'S-1-5-18'='Local System'
            'S-1-5-19'='NT Authority\Local Service'
            'S-1-5-20'='NT Authority\Network Service'
            'S-1-5-32-544'='BUILTIN\Administrators'
            'S-1-5-32-545'='BUILTIN\Users'
            'S-1-5-32-546'='BUILTIN\Guests'
            'S-1-5-32-547'='BUILTIN\Power Users'
        }

        if ($wellKnownSIDs.ContainsKey($SID)) {
            return $wellKnownSIDs[$SID]
        }

        # If in fallback mode, only use .NET translation
        if ($Global:UseFallbackSIDResolution) {
            try {
                $ntAccount = [System.Security.Principal.SecurityIdentifier]::new($SID).Translate([System.Security.Principal.NTAccount])
                return $ntAccount.Value
            }
            catch {
                return $SID
            }
        }

        # Not in fallback mode, try AD module first
        if (Get-Module ActiveDirectory) {
            try {
                $params = @{
                    Filter = "ObjectSID -eq '$SID'"
                    Properties = 'Name', 'SamAccountName'
                    ErrorAction = 'Stop'
                }

                if ($DomainController) {
                    $params['Server'] = $DomainController
                }

                $result = Get-ADObject @params
                if ($result.SamAccountName) {
                    return $result.SamAccountName
                }
                return $result.Name
            }
            catch {
                # AD module failed, try .NET translation as fallback
                try {
                    $ntAccount = [System.Security.Principal.SecurityIdentifier]::new($SID).Translate([System.Security.Principal.NTAccount])
                    return $ntAccount.Value
                }
                catch {
                    return $SID
                }
            }
        }
        else {
            # AD module not available, use .NET translation
            try {
                $ntAccount = [System.Security.Principal.SecurityIdentifier]::new($SID).Translate([System.Security.Principal.NTAccount])
                return $ntAccount.Value
            }
            catch {
                return $SID
            }
        }
    }
    catch {
        # Return original SID if resolution fails
        return $SID
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
    
    # Initialize progress counter
    $processedCount = 0
    $totalCount = $Folders.Count
    
    foreach ($r in $Runspaces) {
        try {
            $result = $r.Instance.EndInvoke($r.Handle)
            $processedCount++
            
            # Update progress bar in non-verbose mode
            if ($VerbosePreference -ne 'Continue' -and $DebugPreference -ne 'Continue') {
                Write-Progress -Activity "Processing Folders" -Status "Processed $processedCount of $totalCount folders" `
                             -PercentComplete (($processedCount / $totalCount) * 100)
            }
            
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
    
    # Complete the progress bar
    if ($VerbosePreference -ne 'Continue' -and $DebugPreference -ne 'Continue') {
        Write-Progress -Activity "Processing Folders" -Completed
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
        param(
            [Parameter(Mandatory=$true)]
            [Array]$Set1,
            [Parameter(Mandatory=$true)]
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
        # Replace the SimplifiedPermissions select statement with this updated version
        $SimplifiedPermissions = $CurrentFolderPermissions | Select-Object @{
            Name = 'Account'
            Expression = { 
                $identity = $_.IdentityReference
                if ($identity -match '^S-\d-\d+(-\d+)+$') {
                    $resolved = Resolve-ADAccountFromSID -SID $identity
                    if ($resolved -ne $identity) {
                        "$resolved ($identity)"
                    } else {
                        $identity
                    }
                } else {
                    $identity
                }
            }
        }, @{
            Name = 'Permissions'
            Expression = { 
                $perms = $_.FileSystemRights
                if ($perms -match '^\(.*\)$') {
                    $perms -replace '^\((.*)\)$', '$1'
                } else {
                    $perms
                }
            }
        }, @{
            Name = 'Type'
            Expression = { $_.AccessControlType }
        }, @{
            Name = 'Inherited'
            Expression = { if ($_.IsInherited) { 'Yes' } else { 'No' } }
        }
        
        # Use Format-Table with custom properties for better readability
        $PermissionsTable = ($SimplifiedPermissions | Format-Table -Wrap -AutoSize | Out-String).Trim()
        
        Write-SafeOutput $PermissionsTable
        [void]$OutputText.Append($PermissionsTable + "`n")
        
        # Mark this folder as displayed
        $DisplayedFolders[$FolderPath] = $true
        
        # If there are subfolders with identical permissions, split them by inheritance
        if ($IdenticalSubfolders.Count -gt 0) {
            $inheritedSubfolders = @()
            $nonInheritedSubfolders = @()
            
            foreach($subfolder in $IdenticalSubfolders) {
                $subfolderPerms = $FolderPermissionsMap[$subfolder].Permissions
                if ($subfolderPerms[0].IsInherited) {
                    $inheritedSubfolders += $subfolder
                } else {
                    $nonInheritedSubfolders += $subfolder
                }
            }
            
            # Display inherited permissions subfolders
            if ($inheritedSubfolders.Count -gt 0) {
                Write-SafeOutput "The following subfolders have identical inherited permissions:" -Color Cyan
                [void]$OutputText.AppendLine("The following subfolders have identical inherited permissions:")
                
                if ($inheritedSubfolders.Count -gt 20) {
                    Write-SafeOutput "  - $($inheritedSubfolders.Count) identical subfolders with inherited permissions" -Color DarkGray
                    [void]$OutputText.AppendLine("  - $($inheritedSubfolders.Count) identical subfolders with inherited permissions")
                    
                    foreach ($Subfolder in $inheritedSubfolders[0..9]) {
                        Write-SafeOutput "  - $Subfolder" -Color DarkGray
                        [void]$OutputText.AppendLine("  - $Subfolder")
                    }
                    Write-SafeOutput "  - ... (and $($inheritedSubfolders.Count - 10) more)" -Color DarkGray
                    [void]$OutputText.AppendLine("  - ... (and $($inheritedSubfolders.Count - 10) more)")
                } else {
                    foreach ($Subfolder in $inheritedSubfolders) {
                        Write-SafeOutput "  - $Subfolder" -Color DarkGray
                        [void]$OutputText.AppendLine("  - $Subfolder")
                    }
                }
            }
            
            # Display explicitly set permissions subfolders
            if ($nonInheritedSubfolders.Count -gt 0) {
                Write-SafeOutput "`nThe following subfolders have identical explicit permissions:" -Color Cyan
                [void]$OutputText.AppendLine("`nThe following subfolders have identical explicit permissions:")
                
                if ($nonInheritedSubfolders.Count -gt 20) {
                    Write-SafeOutput "  - $($nonInheritedSubfolders.Count) identical subfolders with explicit permissions" -Color DarkGray
                    [void]$OutputText.AppendLine("  - $($nonInheritedSubfolders.Count) identical subfolders with explicit permissions")
                    
                    foreach ($Subfolder in $nonInheritedSubfolders[0..9]) {
                        Write-SafeOutput "  - $Subfolder" -Color DarkGray
                        [void]$OutputText.AppendLine("  - $Subfolder")
                    }
                    Write-SafeOutput "  - ... (and $($nonInheritedSubfolders.Count - 10) more)" -Color DarkGray
                    [void]$OutputText.AppendLine("  - ... (and $($nonInheritedSubfolders.Count - 10) more)")
                } else {
                    foreach ($Subfolder in $nonInheritedSubfolders) {
                        Write-SafeOutput "  - $Subfolder" -Color DarkGray
                        [void]$OutputText.AppendLine("  - $Subfolder")
                    }
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

# Process folders and display results
try {
    $CurrentBatch = 0
    $BatchSize = 100 # Process folders in batches of 100
    
    while ($CurrentBatch * $BatchSize -lt $TotalFolders) {
        $StartIndex = $CurrentBatch * $BatchSize
        $EndIndex = [Math]::Min(($CurrentBatch + 1) * $BatchSize, $TotalFolders)
        $CurrentFolders = $Folders[$StartIndex..($EndIndex-1)]
        
        Write-SafeOutput "Processing batch $($CurrentBatch + 1) (folders $($StartIndex + 1) to $EndIndex of $TotalFolders)..." -Color Cyan
        
        # Process the current batch of folders
        $BatchResults = Start-FolderProcessing -Folders $CurrentFolders -MaxThreads $MaxThreads -SkipUniquenessCounting:$SkipUniquenessCounting
        
        # Display results for this batch
        foreach ($Result in $BatchResults.GetEnumerator()) {
            # Display result logic here
            // ...existing code...
        }
        
        $CurrentBatch++
        
        # Ask to continue if there are more folders to process
        if ($EndIndex -lt $TotalFolders) {
            Write-SafeOutput "`nProcessed $EndIndex of $TotalFolders folders." -Color Cyan
            $Continue = Read-Host "Continue to iterate? (Y/N)"
            if ($Continue -notmatch '^[Yy]') {
                Write-SafeOutput "Processing stopped by user after $EndIndex folders" -Color Yellow
                break
            }
        }
    }
    
    # Final summary
    Write-SafeOutput "`nProcessing completed for $EndIndex of $TotalFolders folders" -Color Green
    $TotalTime = ([DateTime]::Now) - $StartTime
    Write-SafeOutput "Total execution time: $($TotalTime.TotalSeconds.ToString('0.00')) seconds" -Color Green
    
    # Save remaining results
    Write-SafeOutput "Saving final results to log file..." -Color DarkGray
    Write-File-Safe -Content $OutputText.ToString() -FilePath $OutputLog -Encoding "UTF8"
    Write-SafeOutput "Results saved to: $OutputLog" -Color Green
}
catch {
    Write-SafeOutput "Error during folder processing: $($_.Exception.Message)" -Color Red
    Write-File-Safe -Content $OutputText.ToString() -FilePath $OutputLog -Encoding "UTF8"
}
