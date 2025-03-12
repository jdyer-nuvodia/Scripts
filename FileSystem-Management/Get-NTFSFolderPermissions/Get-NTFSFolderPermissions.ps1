# =============================================================================
# Script: Get-NTFSFolderPermissions.ps1
# Created: 2025-03-06 21:06:43 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-03-12 17:39:47 UTC
# Updated By: jdyer-nuvodia
# Version: 1.14.3
# Additional Info: Fixed empty permission set display in summary report
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
.PARAMETER EnableSIDDiagnostics
    Enable detailed diagnostic logging for SID resolution attempts.
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
.EXAMPLE
    .\Get-NTFSFolderPermissions.ps1 -FolderPath "C:\Data" -EnableSIDDiagnostics
    Processes permissions with detailed SID resolution logging for troubleshooting
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
    [switch]$SkipADResolution,
    
    [Parameter(Mandatory = $false)]
    [switch]$EnableSIDDiagnostics
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

# Initialize the global variable at script start
$Global:UseFallbackSIDResolution = $false

# Enhanced SID cache for performance
$Global:SIDCache = @{}

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
            # If even direct .NET methods fail, we cannot write to file
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
    # Skip AD module initialization and go straight to fallback mode
    Write-Log "AD module not available - using built-in SID resolution" "Yellow"
    Set-Variable -Name UseFallbackSIDResolution -Value $true -Scope Global
    return $false
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

# Test SID resolution capabilities at startup when diagnostics are enabled
if ($EnableSIDDiagnostics) {
    function Test-SIDResolution {
        # Common SIDs to test
        $testSIDs = @(
            'S-1-5-18',           # Local System
            'S-1-5-32-544',       # BUILTIN\Administrators
            'S-1-5-32-545',       # BUILTIN\Users
            'S-1-5-11'            # Authenticated Users
        )
        
        Write-Log "========== SID RESOLUTION DIAGNOSTIC TEST ==========" -Color Cyan
        Write-Log "Testing SID resolution capabilities..." -Color Cyan
        
        foreach ($sid in $testSIDs) {
            $resolved = Resolve-ADAccountFromSID -SID $sid -EnableDiagnostics
            $status = if ($resolved -eq $sid) { "FAILED" } else { "SUCCESS" }
            $color = if ($resolved -eq $sid) { "Red" } else { "Green" }
            Write-Log "Test SID: $sid -> $resolved ($status)" -Color $color
        }
        
        # Test network connectivity to domain if SIDs failed
        if ($testSIDs | ForEach-Object { Resolve-ADAccountFromSID -SID $_ } | Where-Object { $_ -match '^S-1-' }) {
            Write-Log "WARNING: Some SIDs failed to resolve. Testing network connectivity..." -Color Yellow
            
            try {
                $domainName = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name
                $dcTest = Test-Connection -ComputerName $domainName -Count 1 -Quiet
                Write-Log "Domain connectivity test: $($dcTest ? 'SUCCESS' : 'FAILED')" -Color ($dcTest ? 'Green' : 'Red')
            }
            catch {
                Write-Log "Could not determine current domain. Machine may not be domain-joined." -Color Yellow
            }
        }
        
        Write-Log "============== END DIAGNOSTIC TEST ==============" -Color Cyan
        Write-Log "" # Empty line for spacing
    }
    
    Test-SIDResolution
}

# Enhanced Resolve-ADAccountFromSID function with diagnostic capabilities
function Resolve-ADAccountFromSID {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SID,
        [switch]$EnableDiagnostics
    )
    
    try {
        # Skip if not a valid SID format
        if (-not ($SID -match '^S-\d-\d+(-\d+)+$')) {
            if ($EnableDiagnostics) { Write-Log "Not a valid SID format: $SID" -Color "Yellow" }
            return $SID
        }

        # Check cache first for performance
        if ($Global:SIDCache.ContainsKey($SID)) {
            if ($EnableDiagnostics) { Write-Log "Retrieved from cache: $($Global:SIDCache[$SID])" -Color "Green" }
            return $Global:SIDCache[$SID]
        }
        
        if ($EnableDiagnostics) {
            Write-Log "Attempting to resolve SID: $SID" -Color "Magenta"
        }

        # Well-known SIDs mapping
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
            'S-1-3-2'='Creator Owner Server'
            'S-1-3-3'='Creator Group Server'
            'S-1-3-4'='Owner Rights'
            'S-1-5-1'='Dialup'
            'S-1-5-2'='Network'
            'S-1-5-3'='Batch'
            'S-1-5-4'='Interactive'
            'S-1-5-6'='Service'
            'S-1-5-7'='Anonymous'
            'S-1-5-8'='Proxy'
            'S-1-5-9'='Enterprise Domain Controllers'
            'S-1-5-10'='Principal Self'
            'S-1-5-11'='Authenticated Users'
            'S-1-5-12'='Restricted Code'
            'S-1-5-13'='Terminal Server Users'
            'S-1-5-14'='Remote Interactive Logon'
            'S-1-5-15'='This Organization'
            'S-1-5-17'='IUSR'
            'S-1-5-18'='Local System'
            'S-1-5-19'='NT Authority\Local Service'
            'S-1-5-20'='NT Authority\Network Service'
            'S-1-5-32-544'='BUILTIN\Administrators'
            'S-1-5-32-545'='BUILTIN\Users'
            'S-1-5-32-546'='BUILTIN\Guests'
            'S-1-5-32-547'='BUILTIN\Power Users'
            'S-1-5-32-548'='BUILTIN\Account Operators'
            'S-1-5-32-549'='BUILTIN\Server Operators'
            'S-1-5-32-550'='BUILTIN\Print Operators'
            'S-1-5-32-551'='BUILTIN\Backup Operators'
            'S-1-5-32-552'='BUILTIN\Replicators'
            'S-1-5-32-554'='BUILTIN\Pre-Windows 2000 Compatible Access'
            'S-1-5-32-555'='BUILTIN\Remote Desktop Users'
            'S-1-5-32-556'='BUILTIN\Network Configuration Operators'
            'S-1-5-32-557'='BUILTIN\Incoming Forest Trust Builders'
            'S-1-5-32-558'='BUILTIN\Performance Monitor Users'
            'S-1-5-32-559'='BUILTIN\Performance Log Users'
            'S-1-5-32-560'='BUILTIN\Windows Authorization Access Group'
            'S-1-5-32-561'='BUILTIN\Terminal Server License Servers'
            'S-1-5-32-562'='BUILTIN\Distributed COM Users'
            'S-1-5-32-568'='BUILTIN\IIS_IUSRS'
            'S-1-5-32-569'='BUILTIN\Cryptographic Operators'
            'S-1-5-32-573'='BUILTIN\Event Log Readers'
            'S-1-5-32-574'='BUILTIN\Certificate Service DCOM Access'
            'S-1-5-80-0'='NT SERVICE\ALL SERVICES'
        }

        # Check well-known SIDs first
        if ($wellKnownSIDs.ContainsKey($SID)) {
            if ($EnableDiagnostics) { Write-Log "Resolved from well-known SIDs: $($wellKnownSIDs[$SID])" -Color "Green" }
            $Global:SIDCache[$SID] = $wellKnownSIDs[$SID]
            return $wellKnownSIDs[$SID]
        }

        # Method 1: Try .NET translation - most common approach
        try {
            $sidObj = [System.Security.Principal.SecurityIdentifier]::new($SID)
            $ntAccount = $sidObj.Translate([System.Security.Principal.NTAccount])
            if ($EnableDiagnostics) { Write-Log "Resolved via .NET primary method: $($ntAccount.Value)" -Color "Green" }
            $Global:SIDCache[$SID] = $ntAccount.Value
            return $ntAccount.Value
        }
        catch {
            if ($EnableDiagnostics) { Write-Log "Primary .NET translation failed: $($_.Exception.Message)" -Color "Yellow" }
            
            # Method 2: Try alternate constructor approach
            try {
                $accountName = [System.Security.Principal.NTAccount]::new($SID).Translate([System.Security.Principal.NTAccount]).Value
                if ($EnableDiagnostics) { Write-Log "Resolved via alternate constructor: $accountName" -Color "Green" }
                $Global:SIDCache[$SID] = $accountName
                return $accountName
            }
            catch {
                if ($EnableDiagnostics) { Write-Log "Alternate constructor failed: $($_.Exception.Message)" -Color "Yellow" }
                
                # Method 3: Try using DirectoryServices
                try {
                    $objSID = New-Object System.Security.Principal.SecurityIdentifier($SID)
                    $objUser = $objSID.Translate([System.Security.Principal.NTAccount])
                    if ($EnableDiagnostics) { Write-Log "Resolved via DirectoryServices: $($objUser.Value)" -Color "Green" }
                    $Global:SIDCache[$SID] = $objUser.Value
                    return $objUser.Value
                }
                catch {
                    if ($EnableDiagnostics) { Write-Log "DirectoryServices method failed: $($_.Exception.Message)" -Color "Yellow" }
                    
                    # Method 4: Try using WMI/CIM as last resort
                    try {
                        $query = "SELECT * FROM Win32_Account WHERE SID='$SID'"
                        $account = Get-CimInstance -Query $query -ErrorAction SilentlyContinue
                        
                        if ($account -and $account.Caption) {
                            if ($EnableDiagnostics) { Write-Log "Resolved via WMI: $($account.Caption)" -Color "Green" }
                            $Global:SIDCache[$SID] = $account.Caption
                            return $account.Caption
                        }
                    }
                    catch {
                        if ($EnableDiagnostics) { Write-Log "WMI method failed: $($_.Exception.Message)" -Color "Yellow" }
                    }

                    # If all methods fail, cache the failure and return original SID to avoid repeated lookup attempts
                    if ($EnableDiagnostics) { Write-Log "All SID resolution methods failed for $SID" -Color "Red" }
                    $Global:SIDCache[$SID] = "Unknown Account ($SID)"
                    return "Unknown Account ($SID)"
                }
            }
        }
    }
    catch {
        if ($EnableDiagnostics) { Write-Log "Critical error in SID resolution: $($_.Exception.Message)" -Color "Red" }
        return $SID
    }
}

# Fix for nested Add-Type statement
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public class AccountUtils {
    [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern bool LookupAccountSid(
        string lpSystemName,
        IntPtr Sid,
        StringBuilder lpName,
        ref uint cchName,
        StringBuilder ReferencedDomainName,
        ref uint cchReferencedDomainName,
        out int peUse);
        
    [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern bool ConvertStringSidToSid(
        string StringSid, 
        out IntPtr Sid);
        
    [DllImport("kernel32.dll")]
    public static extern IntPtr LocalFree(IntPtr hMem);
}
"@ -ErrorAction SilentlyContinue

# Continue with the SID resolution method implementation
function Resolve-Win32SID {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SID,
        [switch]$EnableDiagnostics
    )
    
    try {
        if ($EnableDiagnostics) { Write-Log "Attempting Win32 API resolution..." -Color "Magenta" }
        
        $sidPtr = [IntPtr]::Zero
        $success = [AccountUtils]::ConvertStringSidToSid($SID, [ref]$sidPtr)
        
        if ($success) {
            $nameLen = 0
            $domainLen = 0
            $accountType = 0
            
            # First call to get buffer sizes
            [void][AccountUtils]::LookupAccountSid($null, $sidPtr, $null, [ref]$nameLen, $null, [ref]$domainLen, [ref]$accountType)
            
            $name = New-Object System.Text.StringBuilder($nameLen)
            $domain = New-Object System.Text.StringBuilder($domainLen)
            
            # Second call to get actual values
            $result = [AccountUtils]::LookupAccountSid($null, $sidPtr, $name, [ref]$nameLen, $domain, [ref]$domainLen, [ref]$accountType)
            
            if ($result) {
                $resolvedAccount = if ($domain.Length -gt 0) { "$($domain)\$($name)" } else { "$($name)" }
                if ($EnableDiagnostics) { Write-Log "Resolved via Win32 API: $resolvedAccount" -Color "Green" }
                $Global:SIDCache[$SID] = $resolvedAccount
                return $resolvedAccount
            }
        }
        
        # Cleanup
        if ($sidPtr -ne [IntPtr]::Zero) {
            [void][AccountUtils]::LocalFree($sidPtr)
        }
    }
    catch {
        if ($EnableDiagnostics) { Write-Log "Win32 API method failed: $($_.Exception.Message)" -Color "Yellow" }
    }
    
    # Method 6: Network connectivity test and retry
    try {
        if ($EnableDiagnostics) { Write-Log "Testing network connectivity before final retry..." -Color "Magenta" }
        
        # Check if we can ping the domain controller
        $domainName = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name
        if (Test-Connection -ComputerName $domainName -Count 1 -Quiet) {
            if ($EnableDiagnostics) { Write-Log "Network is available, final retry with full SID..." -Color "Green" }
            
            # One last desperate attempt with full SID object creation
            $sidBytes = New-Object byte[] 28
            $sidObj = New-Object Security.Principal.SecurityIdentifier($SID)
            $sidObj.GetBinaryForm($sidBytes, 0)
            $ntAccount = New-Object Security.Principal.SecurityIdentifier($sidBytes, 0).Translate([Security.Principal.NTAccount])
            
            if ($ntAccount) {
                if ($EnableDiagnostics) { Write-Log "Resolved via binary SID method: $($ntAccount.Value)" -Color "Green" }
                $Global:SIDCache[$SID] = $ntAccount.Value
                return $ntAccount.Value
            }
        }
    }
    catch {
        if ($EnableDiagnostics) { Write-Log "Network test/final method failed: $($_.Exception.Message)" -Color "Red" }
    }
    
    # If all resolution methods fail, return the original SID
    if ($EnableDiagnostics) { Write-Log "All resolution methods failed, returning original SID" -Color "Yellow" }
    return "Unknown Account ($SID)"
}

# Get the permissions for a folder
function Get-FolderPermissions {
    param(
        [string]$FolderPath
    )
    
    try {
        # Use cross-platform compatible approach to get ACLs
        $acl = $null
        
        # Method 1: Try to use Get-Acl cmdlet which works in both PowerShell versions
        try {
            $acl = Get-Acl -Path $FolderPath -ErrorAction Stop
        }
        catch {
            Write-Log "Get-Acl failed for $FolderPath : $($_.Exception.Message)" -Color "Yellow"
            
            # Method 2: Try alternate .NET approach for PowerShell Core
            try {
                $dirInfo = [System.IO.DirectoryInfo]::new($FolderPath)
                if ($PSVersionTable.PSVersion.Major -ge 6) {
                    # PowerShell Core approach
                    $acl = [System.IO.FileSystemAclExtensions]::GetAccessControl($dirInfo)
                }
                else {
                    # Windows PowerShell 5.1 approach
                    $acl = [System.Security.AccessControl.DirectorySecurity]::new()
                    $acl = $dirInfo.GetAccessControl([System.Security.AccessControl.AccessControlSections]::Access)
                }
            }
            catch {
                Write-Log "Both ACL retrieval methods failed for $FolderPath : $($_.Exception.Message)" -Color "Yellow"
                throw
            }
        }
        
        if ($null -eq $acl) {
            throw "Could not retrieve ACL for $FolderPath"
        }
        
        # Helper function for converting inheritance flags to string descriptions
        function Convert-InheritanceToDescription {
            param (
                $InheritanceFlags,
                $PropagationFlags
            )
            
            # Start with an empty string
            $description = @()
            
            # Check the inheritance flags
            if ($InheritanceFlags -band [InheritanceFlags]::ContainerInherit) {
                $description += "Container inherit"
            }
            if ($InheritanceFlags -band [InheritanceFlags]::ObjectInherit) {
                $description += "Object inherit"
            }
            if ($InheritanceFlags -eq [InheritanceFlags]::None) {
                $description += "None"
            }
            
            # Check propagation flags
            if ($PropagationFlags -band [PropagationFlags]::NoPropagateInherit) {
                $description += "Do not propagate"
            }
            if ($PropagationFlags -band [PropagationFlags]::InheritOnly) {
                $description += "Inherit only"
            }
            
            # Return description
            return $description -join ", "
        }
        
        # Get all access rules
        $permissions = @()
        foreach ($rule in $acl.GetAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier])) {
            # Resolve SIDs to friendly names if possible
            $identity = $rule.IdentityReference.Value
            
            # Only try to resolve if it looks like a SID and resolution isn't disabled
            if (-not $SkipADResolution -and $identity -match '^S-\d-') {
                $resolved = Resolve-ADAccountFromSID -SID $identity -EnableDiagnostics:$EnableSIDDiagnostics
                $identity = "$resolved ($identity)"
            }
            
            # Create a permission object
            $permissions += [PSCustomObject]@{
                IdentityReference = $identity
                FileSystemRights = $rule.FileSystemRights.ToString()
                AccessControlType = $rule.AccessControlType.ToString()
                IsInherited = $rule.IsInherited
                InheritanceFlags = Convert-InheritanceToDescription -InheritanceFlags $rule.InheritanceFlags -PropagationFlags $rule.PropagationFlags
                PropagationFlags = $rule.PropagationFlags.ToString()
            }
        }
        
        # Validation - ensure we have permissions, or add a placeholder
        if ($permissions.Count -eq 0) {
            $permissions += [PSCustomObject]@{
                IdentityReference = "DEFAULT"
                FileSystemRights = "Unknown - Possible Access Issue"
                AccessControlType = "N/A"
                IsInherited = $true
                InheritanceFlags = "None"
                PropagationFlags = "None"
            }
            Write-Log "Warning: No permissions extracted for $FolderPath - using placeholder" -Color "Yellow"
        }
        
        # Generate a hash of the permissions for comparison
        $sortedPermissions = $permissions | Sort-Object IdentityReference, FileSystemRights, AccessControlType
        $permissionStrings = $sortedPermissions | ForEach-Object {
            "$($_.IdentityReference)|$($_.FileSystemRights)|$($_.AccessControlType)|$($_.IsInherited)"
        }
        $permissionHash = ($permissionStrings -join ';').GetHashCode()
        
        return @{
            Success = $true
            FolderPath = $FolderPath
            Permissions = $permissions
            Hash = $permissionHash
        }
    }
    catch {
        # Return an error object
        return @{
            Success = $false
            FolderPath = $FolderPath
            Error = $_.Exception.Message
            Permissions = @([PSCustomObject]@{
                IdentityReference = "ACCESS ERROR"
                FileSystemRights = "Error: $($_.Exception.Message)"
                AccessControlType = "N/A"
                IsInherited = $true
                InheritanceFlags = "N/A"
                PropagationFlags = "N/A"
            })
            Hash = -1
        }
    }
}

# Main script execution

try {
    # Build a list of all folders to process
    $startFolder = [System.IO.DirectoryInfo]::new($FolderPath)
    $folders = @($startFolder) + (Get-AllDirectoriesModule -Path $FolderPath -MaxDepth $MaxDepth)
    
    Write-Log "Found $($folders.Count) folders to process" -Color "Cyan"
    
    # Initialize results collection
    $allResults = @()
    $processedCount = 0
    
    # Process folders in batches using runspaces for parallelization
    $batchSize = [Math]::Min($MaxThreads, $folders.Count)
    $queue = [System.Collections.Queue]::new($folders)
    
    while ($queue.Count -gt 0) {
        $batch = @()
        for ($i = 0; $i -lt $batchSize -and $queue.Count -gt 0; $i++) {
            $batch += $queue.Dequeue()
        }
        
        $results = foreach ($folder in $batch) {
            $processedCount++
            Invoke-FolderProcessing -Path $folder.FullName -CurrentCount $processedCount -TotalCount $folders.Count
        }
        
        # Add results to our collection
        $allResults += $results | Where-Object { $_ -ne $null }
    }
    
    # Process the results
    Write-Progress -Activity "Processing Results" -Status "Creating summary report..." -PercentComplete 100
    
    # Group folders by permission hash to identify identical sets
    $permissionGroups = $allResults | Group-Object { $_.Hash }
    
    # Log total folders with unique permissions
    $uniquePermissionCount = $permissionGroups.Count
    $successfullyProcessed = ($allResults | Where-Object { $_.Success }).Count
    $failedProcessed = ($allResults | Where-Object { -not $_.Success }).Count
    
    Write-Log ""
    Write-Log "====== SUMMARY REPORT ======" -Color "Cyan"
    Write-Log "Folders processed: $successfullyProcessed" -Color "Green"
    Write-Log "Failed to process: $failedProcessed" -Color "Yellow"
    Write-Log "Unique permission sets found: $uniquePermissionCount" -Color "Cyan"
    Write-Log ""
    
    # Display and log unique permission sets
    $setCounter = 1
    foreach ($group in $permissionGroups | Sort-Object { $_.Group.Count } -Descending) {
        $folderCount = $group.Group.Count
        $firstFolder = $group.Group[0].FolderPath
        $permissions = $group.Group[0].Permissions
        
        # Determine the group name
        $folderDisplay = if ($folderCount -eq 1) {
            # For a single folder, just show its path
            $firstFolder
        }
        else {
            # For multiple folders, show the first one and the count
            "$firstFolder (and $($folderCount - 1) other folder$(if ($folderCount -gt 2) {'s'}))"
        }
        
        Write-Log ""
        Write-Log "Permission Set #$setCounter - Applied to $folderCount folder$(if($folderCount -ne 1){'s'})" -Color "Cyan"
        Write-Log "Example: $folderDisplay" -Color "Cyan"
        Write-Log "-------------------------" -Color "Cyan"
        
        # Display permissions - ensure we have permissions to display
        if ($null -eq $permissions -or $permissions.Count -eq 0) {
            Write-Log "No permissions found or error retrieving permissions" -Color "Yellow"
        } else {
            # Ensure we're processing the permissions collection properly
            foreach ($perm in $permissions | Sort-Object IdentityReference, FileSystemRights) {
                $color = switch ($perm.AccessControlType) {
                    "Allow" { "White" }
                    "Deny" { "Red" }
                    default { "Yellow" }
                }
                
                $inheritedText = if ($perm.IsInherited) { " (Inherited)" } else { "" }
                
                # Handle null or empty inheritance flags
                $inheritanceInfo = ""
                if ($perm.InheritanceFlags -and $perm.InheritanceFlags -ne "None") {
                    $inheritanceInfo = " [$($perm.InheritanceFlags)]"
                }
                
                # Ensure we have a complete and properly formatted output line
                $permissionLine = "$($perm.IdentityReference): $($perm.FileSystemRights) - $($perm.AccessControlType)$inheritedText$inheritanceInfo"
                Write-Log $permissionLine -Color $color
            }
        }
        
        # Increment counter
        $setCounter++
    }
    
    if ($failedProcessed -gt 0) {
        Write-Log ""
        Write-Log "====== FOLDERS WITH ERRORS ======" -Color "Yellow"
        foreach ($failedItem in $allResults | Where-Object { -not $_.Success }) {
            Write-Log "Failed: $($failedItem.FolderPath) - $($failedItem.Error)" -Color "Red"
        }
    }
    
    # Show execution time
    $endTime = [DateTime]::Now
    $duration = $endTime - $StartTime
    $formattedDuration = "{0:D2}:{1:D2}:{2:D2}.{3:D3}" -f $duration.Hours, $duration.Minutes, $duration.Seconds, $duration.Milliseconds
    
    Write-Log ""
    Write-Log "==============================" -Color "Cyan"
    Write-Log "Report generated: $(Get-Date)" -Color "Cyan"
    Write-Log "Total execution time: $formattedDuration" -Color "Cyan"
    Write-Log "Full report saved to: $OutputLog" -Color "Cyan"
    Write-Log "==============================" -Color "Cyan"
    
    # Ensure output is written to file
    try {
        [System.IO.File]::WriteAllText($OutputLog, $OutputText.ToString(), [System.Text.Encoding]::UTF8)
    }
    catch {
        Write-Log "Error saving report: $($_.Exception.Message)" -Color "Red"
    }
}
catch {
    $errorMsg = "Critical error: $($_.Exception.Message)`n$($_.InvocationInfo.PositionMessage)"
    Write-Log $errorMsg -Color "Red"
    Write-Output-Safe $errorMsg
    exit 1
}
finally {
    # Clean up any remaining runspaces
    if ($runspacePool) {
        $runspacePool.Close()
        $runspacePool.Dispose()
    }
}