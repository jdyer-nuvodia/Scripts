# =============================================================================
# Script: Get-NTFSPermissionsForUser.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-08 15:45:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.8.7
# Additional Info: Fixed PSScriptAnalyzer compliance issues for unused variables:
#   - Removed intermediate $baseLogName variable and directly constructed log file paths
#   - Replaced ForEach-Object pipeline with foreach loop to fix variable scoping issue with $foundPermissions
#   - Fixed string interpolation syntax for consistent variable usage throughout log file naming
# =============================================================================

<#
.SYNOPSIS
    Gets NTFS permissions for specified users or SIDs in a directory structure.
.DESCRIPTION
    This script recursively checks NTFS permissions for specified users or SIDs
    across all folders under a given root directory. It can accept either a direct
    user input or read from a SIDS.txt file. Results are displayed in the console
    and saved to a transcript file, with detailed scanning logged to a debug file.
.PARAMETER User
    The username to check permissions for. Must include domain name if on a domain (e.g., "DOMAIN\username")
.PARAMETER StartPath
    The starting folder path to begin the recursive permission check
.PARAMETER SIDFile
    Optional path to a text file containing SIDs to check (one per line)
.EXAMPLE
    .\Get-NTFSPermissionsForUser.ps1 -User "DOMAIN\jsmith" -StartPath "D:\"
    Checks permissions for user DOMAIN\jsmith starting from D:\ drive

    Example output when a permission is found:
    Identity: DOMAIN\jsmith
    Path: D:\Shared\Projects
    Access: FullControl
    Type: Allow
    Owner Status: Owner

.EXAMPLE
    .\Get-NTFSPermissionsForUser.ps1 -StartPath "D:\" -SIDFile "C:\SIDS.txt"
    Checks permissions for all SIDs listed in SIDS.txt starting from D:\ drive
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$User,

    [Parameter(Mandatory = $true)]
    [string]$StartPath,

    [Parameter(Mandatory = $false)]
    [string]$SIDFile
)

# Initialize Logging
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$systemName = $env:COMPUTERNAME

# Determine user string for filename
$userString = if ($User) {
    $userParts = $User.Split('\')
    # Take the last part after splitting on backslash
    $userParts[-1]
} elseif ($SIDFile) {
    "multiple_users"
} else {
    "no_user"
}

# Construct log file names
$debugLogFile = Join-Path $PSScriptRoot "NTFSPermissionsForUser_${systemName}_${userString}_${timestamp}_debug.log"
$transcriptFile = Join-Path $PSScriptRoot "NTFSPermissionsForUser_${systemName}_${userString}_${timestamp}_transcript.log"

# Start Transcript
Start-Transcript -Path $transcriptFile

try {
    # Function to write debug log with error handling
    function Write-DebugLog {
        param(
            [string]$Message,
            [switch]$IsError
        )
        try {
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logEntry = "$timestamp - $Message"
            if ($IsError) {
                $logEntry = "$logEntry [ERROR]"
                Write-Warning $Message
            }
            $logEntry | Out-File -FilePath $debugLogFile -Append -ErrorAction Stop
        } catch {
            Write-Warning "Failed to write to debug log: $_"
        }
    }

    # Add enhanced owner resolution function
    function Get-FolderOwner {
        param(
            [string]$FolderPath
        )
        try {
            $acl = Get-Acl -Path $FolderPath -ErrorAction Stop
            $ownerSid = $acl.Owner

            # If already in domain\user format, return as is
            if ($ownerSid -notmatch '^S-1-') {
                return $ownerSid
            }

            # Try to resolve the SID
            try {
                $objSID = New-Object System.Security.Principal.SecurityIdentifier($ownerSid)
                $objOwner = $objSID.Translate([System.Security.Principal.NTAccount])
                return $objOwner.Value
            } catch {
                Write-DebugLog "Could not resolve owner SID: $ownerSid - $_" -IsError
                return $ownerSid
            }
        } catch {
            Write-DebugLog "Error getting owner for $FolderPath : $_" -IsError
            return "Unknown"
        }
    }

    Write-Output "Initializing permission scan..."

    # Get list of identities to check
    $identities = @()
    if ($User) {
        $identities += $User
    }
    if ($SIDFile -and (Test-Path $SIDFile)) {
        # Load and clean up identities from file
        $identities += Get-Content $SIDFile | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() }
        Write-Output "Loaded $(($identities | Measure-Object).Count) identities from $SIDFile"
    } elseif ($SIDFile) {
        Write-Warning "SID file not found: $SIDFile"
    }

    if ($identities.Count -eq 0) {
        Write-Error "No identities specified. Please provide either a User or a SIDFile."
        Stop-Transcript
        return
    }

    # Display identities being checked
    Write-Output "`nChecking permissions for the following identities:"
    foreach ($identity in $identities) {
        Write-Output "  - $identity"
    }
    Write-Output "`nStarting path: $StartPath"
    Write-Output "Processing... Please wait...`n"

    # Get total folder count for progress bar
    $totalFolders = (Get-ChildItem -Directory -Path $StartPath -Recurse -Force | Measure-Object).Count

    # Function to check and report folder permissions
    function Test-FolderPermission {
        param(
            [string]$FolderPath,
            [array]$IdentityList,
            [ref]$ProcessedCount
        )
        try {
            $acl = Get-Acl -Path $FolderPath -ErrorAction Stop
            $foundPermissions = $false

            # Get owner SID, handling the "O:" prefix
            $ownerSid = $acl.Owner
            if ($ownerSid -match '^O:(?<sid>S-1-[0-9-]+)') {
                $ownerSid = $matches['sid']
            } elseif ($ownerSid -notmatch '^S-1-') {
                try {
                    $ntAccount = New-Object System.Security.Principal.NTAccount($ownerSid)
                    $ownerSid = $ntAccount.Translate([System.Security.Principal.SecurityIdentifier]).Value
                } catch {
                    Write-DebugLog "Could not convert owner to SID: $ownerSid - $_" -IsError
                }
            }

            # Filter out empty identities and process each valid one
            foreach ($identity in ($IdentityList | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
                $cleanIdentity = $identity.Trim()

                # Compare owner SID with identity
                $isOwner = $ownerSid -eq $cleanIdentity -or $acl.Owner -eq $cleanIdentity

                # Get access rules for this identity
                $accessRules = $acl.Access | Where-Object {
                    $_.IdentityReference.Value -eq $cleanIdentity
                }

                if ($isOwner) {
                    $foundPermissions = $true
                    Write-Output "`n[Found owner match!]"
                    Write-PermissionInfo -Identity $cleanIdentity `
                        -Path $FolderPath `
                        -Access "FullControl" `
                        -AccessType ([System.Security.AccessControl.AccessControlType]::Allow) `
                        -OwnerStatus "Owner"
                }

                if ($accessRules) {
                    $foundPermissions = $true
                    foreach ($rule in $accessRules) {
                        Write-PermissionInfo -Identity $cleanIdentity `
                            -Path $FolderPath `
                            -Access $rule.FileSystemRights `
                            -AccessType $rule.AccessControlType `
                            -OwnerStatus $(if ($isOwner) { "Owner" } else { "Not Owner" })
                    }
                }
            }

            if ($foundPermissions) {
                Write-DebugLog "Found permissions/ownership in: $FolderPath"
            } else {
                Write-DebugLog "No matching permissions found in: $FolderPath"
            }

            $ProcessedCount.Value++
            $percentComplete = [math]::Min(($ProcessedCount.Value / $totalFolders) * 100, 100)
            Write-Progress -Activity "Scanning folders for permissions and ownership" -Status "Processing: $FolderPath" -PercentComplete $percentComplete

            # Return the flag indicating if permissions were found
            return $foundPermissions
        } catch {
            Write-DebugLog "Error checking permissions/ownership for $FolderPath : $_" -IsError
            return $false
        }
    }

    # Enhanced folder scanning function with progress tracking
    function Find-Folder {
        param(
            [string]$FolderPath,
            [array]$IdentityList,
            [ref]$ProcessedCount
        )
        try {
            Write-DebugLog "Scanning folder: $FolderPath"
            Test-FolderPermission -FolderPath $FolderPath -IdentityList $IdentityList -ProcessedCount $ProcessedCount

            Get-ChildItem -Path $FolderPath -Directory -ErrorAction Stop | ForEach-Object {
                try {
                    Find-Folder -FolderPath $_.FullName -IdentityList $IdentityList -ProcessedCount $ProcessedCount
                } catch {
                    Write-Error "Error scanning subfolder $($_.FullName)"
                    Write-DebugLog "Error scanning subfolder $($_.FullName): $_" -IsError
                }
            }
        } catch {
            Write-Error "Error accessing folder $FolderPath"
            Write-DebugLog "Error accessing folder $FolderPath : $_" -IsError
        }
    }

    # Modify Write-PermissionInfo to properly handle and display owner information
    function Write-PermissionInfo {
        param(
            [string]$Identity,
            [string]$Path,
            [System.Security.AccessControl.FileSystemRights]$Access,
            [System.Security.AccessControl.AccessControlType]$AccessType,
            [string]$OwnerStatus = "Not Owner"
        )

        $owner = Get-FolderOwner -FolderPath $Path
        $permissionInfo = @"

Folder: $Path
Owner: $owner
Identity: $Identity
Access: $Access
Type: $AccessType
Owner Status: $OwnerStatus
"@

        Write-Output $permissionInfo
        Write-DebugLog $permissionInfo
    }

    # Function to check ownership and permissions
    function Test-FolderAccess {
        param(
            [string]$Path,
            [string[]]$Identities
        )

        try {
            $Acl = Get-Acl -Path $Path -ErrorAction Stop
            $Owner = $Acl.Owner

            foreach ($Identity in $Identities) {
                # Convert username to SID if it's not already a SID
                $Sid = if ($Identity -match '^S-1-') {
                    $Identity
                } else {
                    try {
                        $NTAccount = New-Object System.Security.Principal.NTAccount($Identity)
                        $NTAccount.Translate([System.Security.Principal.SecurityIdentifier]).Value
                    } catch {
                        Write-Warning "Could not convert $Identity to SID"
                        continue
                    }
                }

                # Check ownership - comparing both SID forms and string forms
                $OwnerSid = try {
                    $NTAccount = New-Object System.Security.Principal.NTAccount($Owner)
                    $NTAccount.Translate([System.Security.Principal.SecurityIdentifier]).Value
                } catch {
                    $Owner
                }

                if ($OwnerSid -eq $Sid -or $Owner -eq $Identity) {
                    Write-Output "Owner Status: Owner"
                    Write-Output "Path: $Path"
                    Write-Output "Identity: $Identity"
                    Write-Output "---"
                }
            }
        } catch {
            Write-Warning "Could not access $Path"
        }
    }

    # Function to get folder permissions and owner information
    function Get-FolderPermission {
        param (
            [Parameter(Mandatory = $true)]
            [string]$Folder,
            [string[]]$TargetIdentities
        )

        try {
            $acl = Get-Acl -Path $Folder
            $ownerSID = $acl.Owner
            if ($ownerSID -match '^O:(?<sid>S-1-[0-9-]+)') {
                $ownerSID = $matches['sid']
            } elseif ($ownerSID -match '^S-1-[0-9-]+') {
                $ownerSID = $ownerSID
            } else {
                try {
                    $ntAccount = [System.Security.Principal.NTAccount]$acl.Owner
                    $ownerSID = $ntAccount.Translate([System.Security.Principal.SecurityIdentifier]).Value
                } catch {
                    $ownerSID = "Unknown"
                }
            }

            # Check if owner is in target identities
            if ($TargetIdentities -contains $ownerSID) {
                Write-Output "`n$Folder"
                Write-Output "Owner: O:$ownerSID"
            }

            $permissions = @()
            foreach ($ace in $acl.Access) {
                if ($ace.IdentityReference.Value -in $TargetIdentities) {
                    $permissions += [PSCustomObject]@{
                        Path        = $Folder
                        Identity    = $ace.IdentityReference.Value
                        Type        = $ace.AccessControlType
                        Rights      = $ace.FileSystemRights
                        IsInherited = $ace.IsInherited
                    }
                }
            }

            if ($permissions.Count -gt 0) {
                foreach ($perm in $permissions) {
                    Write-Output "$($perm.Identity) has $($perm.Type) $($perm.Rights) $(if($perm.IsInherited) { '(Inherited)'})"
                }
            }

            return $permissions
        } catch {
            Write-Log "Error getting permissions for ${Folder}: $_" -Level 'ERROR' -Color "Red"
            return $null
        }
    }

    # Initialize progress tracking
    $processedFolders = [ref]0

    # Start the scan with progress tracking
    Write-Output "Beginning folder scan from $StartPath"
    Find-Folder -FolderPath $StartPath -IdentityList $identities -ProcessedCount $processedFolders

    # Output final statistics
    Write-Progress -Activity "Scanning folders for permissions" -Completed
    Write-Output "`nScan Complete"
    Write-Output "Scan Statistics:"
    Write-Output "Folders processed: $($processedFolders.Value)/$totalFolders"
    Write-Output "Debug log: $debugLogFile"
    Write-Output "Transcript: $transcriptFile"
} catch {
    Write-Error "An error occurred during script execution: $_"
    Write-DebugLog "Script execution error: $_" -IsError
} finally {
    # Clean up all progress bars
    Write-Progress -Activity "Analyzing and Updating Folder Permissions" -Id 0 -Completed
    Write-Progress -Activity "Current Folder Analysis" -Id 1 -Completed

    # Stop transcript
    Stop-Transcript
    Write-Output "`nScript execution completed"
}
