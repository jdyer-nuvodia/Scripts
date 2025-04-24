# =============================================================================
# Script: Apply-WIBRSPRegistryChange.ps1
# Created: 2025-04-24 18:10:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-24 22:17:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.3.9
# Additional Info: Fixed binary value comparison in verification process.
# =============================================================================

<#
.SYNOPSIS
Applies a specific registry change (OneDrive TimerAutoMount) to a user's registry hive.

.DESCRIPTION
This script applies a specific registry setting (OneDrive TimerAutoMount = 1) based on the
content originally from 'SharePoint_Auto_Mount.reg'. It handles applying the change
to the current user, a specified user who is logged off (by loading their NTUSER.DAT),
or a specified user who is currently logged in (by targeting their live HKEY_USERS hive).
The script includes detailed logging and -WhatIf support.

Requires running with sufficient privileges (like SYSTEM) to load/modify other users' hives.

.PARAMETER Username
The Windows username for which to apply the registry changes. If not specified,
changes will be applied to the current user running the script.

.PARAMETER LoadedHivePathForVerification
Internal parameter used when the hive was already loaded for the apply step.

.PARAMETER WhatIf
Shows what would happen if the script runs without actually applying the changes.

.EXAMPLE
.\Apply-WIBRSPRegistryChange.ps1
Applies the registry change to the current user.

.EXAMPLE
.\Apply-WIBRSPRegistryChange.ps1 -Username "jsmith"
Applies the registry change to the jsmith user profile (logged in or off).

.EXAMPLE
.\Apply-WIBRSPRegistryChange.ps1 -Username "jsmith" -WhatIf
Shows what registry changes would be applied to the jsmith user profile without making changes.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$Username,

    # Internal parameter to reuse already loaded hive for verification
    [Parameter(Mandatory = $false)]
    [string]$LoadedHivePathForVerification
)

# Set error action preference
$ErrorActionPreference = "Stop"

# Define log file path in the same directory as script
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$logFile = Join-Path -Path $scriptPath -ChildPath "Apply-WIBRSPRegistryChange_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Define specific registry values to apply
$regKeyRelativePath = "Software\Microsoft\OneDrive\Accounts\Business1"
$regValueName = "TimerAutoMount"
$regValueData = [byte[]]@(1,0,0,0,0,0,0,0)
$regValueType = [Microsoft.Win32.RegistryValueKind]::Binary

# Function to write to log file and console with color-coding
function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "PROCESS", "SUCCESS", "WARNING", "ERROR", "DEBUG", "DETAIL")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Write to log file
    Add-Content -Path $logFile -Value $logMessage
    
    # Write to console with color-coding
    switch ($Level) {
        "INFO"    { Write-Host $logMessage -ForegroundColor White }
        "PROCESS" { Write-Host $logMessage -ForegroundColor Cyan }
        "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        "ERROR"   { Write-Host $logMessage -ForegroundColor Red }
        "DEBUG"   { Write-Host $logMessage -ForegroundColor Magenta }
        "DETAIL"  { Write-Host $logMessage -ForegroundColor DarkGray }
    }
}

# Function to get User SID
function Get-UserSID {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Username
    )
    try {
        $account = New-Object System.Security.Principal.NTAccount($Username)
        $sid = $account.Translate([System.Security.Principal.SecurityIdentifier]).Value
        Write-Log "Found SID for user '$Username': $sid" "DETAIL"
        return $sid
    } catch {
        Write-Log "Failed to retrieve SID for user '$Username'. Error: $_" "ERROR"
        throw "Could not resolve SID for user '$Username'."
    }
}

# Function to check if user is logged in using quser or by checking current identity
function Test-UserLoggedIn {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Username
    )
    Write-Log "Checking login status for user '$Username'..." "DETAIL"
    $trimmedUsername = $Username.Trim() # Trim whitespace just in case

    # 1. Check if the target username matches the current user running the script
    $currentIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $currentFullName = $currentIdentity.Name.Trim() # e.g., NUVODIA\jdyer
    
    # Handle backslash correctly by escaping it properly in the split
    $splitName = $currentFullName -split '\\'
    $currentUserNameOnly = if ($splitName.Count -ge 2) { $splitName[$splitName.Count-1].Trim() } else { $currentFullName.Trim() }
    
    Write-Log "Current script identity: $currentFullName" "DEBUG"
    Write-Log "Current username only (after split): '$currentUserNameOnly'" "DEBUG"
    Write-Log "Comparing target (trimmed): '$trimmedUsername' with current name only: '$currentUserNameOnly'" "DEBUG"

    # --- Enhanced Debugging --- Start
    Write-Log "Debug: trimmedUsername = '$trimmedUsername' (Length: $($trimmedUsername.Length))" "DEBUG"
    Write-Log "Debug: currentUserNameOnly = '$currentUserNameOnly' (Length: $($currentUserNameOnly.Length))" "DEBUG"
    Write-Log "Debug: currentFullName = '$currentFullName' (Length: $($currentFullName.Length))" "DEBUG"
    
    # Use case-insensitive comparisons
    $comparison1 = $trimmedUsername -ieq $currentUserNameOnly
    $comparison2 = $trimmedUsername -ieq $currentFullName
    Write-Log "Debug: Case-insensitive comparison 1 ('$trimmedUsername' -ieq '$currentUserNameOnly') Result: $comparison1" "DEBUG"
    Write-Log "Debug: Case-insensitive comparison 2 ('$trimmedUsername' -ieq '$currentFullName') Result: $comparison2" "DEBUG"
    # --- Enhanced Debugging --- End

    # Use pre-calculated case-insensitive comparison results
    if ($comparison1 -or $comparison2) {
        Write-Log "Target user '$trimmedUsername' matches the current script user '$currentFullName'. Assuming logged in." "PROCESS"
        return $true
    }

    # 2. If not the current user, fall back to quser check
    Write-Log "Target user '$trimmedUsername' does not match current script user '$currentFullName'. Checking quser output..." "DETAIL"
    $quserOutput = try {
        # Ensure quser exists before trying to run it
        if (Get-Command quser -ErrorAction SilentlyContinue) {
             quser 2>&1 | Out-String
        } else {
            Write-Log "'quser.exe' not found. Assuming user is not logged in." "WARNING"
            return $false
        }
    } catch {
        Write-Log "Failed to execute quser. Error: $_" "WARNING"
        # Assume user is not logged in if quser fails (conservative approach)
        return $false
    }

    # Improved matching: handle domain\user format and ensure it's the username field
    # Example quser output line: ">consoleuser       console            1  Active    .        2024-04-24 10:00 AM"
    # Or for domain user:        " domain\user       rdp-tcp#0          2  Active    1+00:00  2024-04-24 11:00 AM"
    # Regex explanation:
    #   ^                 - Start of line
    #   >?                - Optional '>' character at the beginning
    #   \s*               - Zero or more whitespace characters
    #   (?:[^\s]+\\)?     - Optional non-capturing group for 'domain\'
    #   $([regex]::Escape($trimmedUsername)) - The literal username, properly escaped
    #   \s+               - One or more whitespace characters (separating username from session name)
    $regexPattern = "^\\s*>?\\s*(?:[^\\s]+\\\\)?$([regex]::Escape($trimmedUsername))\\s+" # Use trimmed username
    if ($quserOutput -match $regexPattern) {
        Write-Log "User '$trimmedUsername' appears to be logged in based on quser output." "PROCESS"
        return $true
    } else {
        Write-Log "User '$trimmedUsername' does not appear to be logged in based on quser output." "PROCESS"
        return $false
    }
} # End of Test-UserLoggedIn

# Function to test registry changes
function Test-RegistryChanges {
    param (
        [Parameter(Mandatory = $false)]
        [string]$Username,

        [Parameter(Mandatory = $false)]
        [string]$LoadedHivePathForVerification # Expects format like HKLM\TempHive
    )

    Write-Log "Verifying registry changes..." "PROCESS"

    $verificationPath = $null
    $userSID = $null

    try {
        if ($Username) {
            # Check if a pre-loaded hive path was provided
            if ($LoadedHivePathForVerification) {
                 # Hive already loaded by caller, use the provided path directly
                 # Construct the full PS provider path with proper backslash formatting
                 $verificationPath = "Registry::$LoadedHivePathForVerification\$regKeyRelativePath"
                 Write-Log "Verifying pre-loaded hive path: $verificationPath" "DETAIL"
            } else {
                # Hive not pre-loaded by caller, check if user is logged in to determine target
                $isUserLoggedInVerify = Test-UserLoggedIn -Username $Username
                if ($isUserLoggedInVerify) {
                    # User is logged in, check HKEY_USERS\<SID>
                    $userSID = Get-UserSID -Username $Username
                    if (-not $userSID) { return $false } # Get-UserSID throws on failure, but double-check
                    $verificationPath = "Registry::HKEY_USERS\$userSID\$regKeyRelativePath"
                    Write-Log "Verifying logged-in user's live hive path: $verificationPath" "DETAIL"
                } else {
                    # User is not logged in AND hive wasn't pre-loaded by the caller.
                    # This function instance cannot load the hive itself anymore.
                    Write-Log "Verification cannot proceed for logged-off user '$Username' without a pre-loaded hive path." "ERROR"
                    return $false
                }
            }
        } else {
            # No username specified, check current user (HKCU)
            $verificationPath = "Registry::HKEY_CURRENT_USER\$regKeyRelativePath"
            Write-Log "Verifying current user path: $verificationPath" "DETAIL"
        }

        # Perform the actual check
        if (Test-Path -Path $verificationPath) {
            try {
                $currentValue = Get-ItemProperty -Path $verificationPath -Name $regValueName -ErrorAction Stop | 
                                Select-Object -ExpandProperty $regValueName
                
                Write-Log "✓ Verification: Registry key exists and contains the value '$regValueName'." "DETAIL"
                
                if ($null -eq $currentValue) {
                    Write-Log "✗ Verification FAILED: Value exists but is null." "WARNING"
                    return $false
                } else {
                # Determine how to compare values based on type
                    if ($currentValue -is [byte[]]) {
                        # For byte arrays, compare them using a more reliable method
                        $isEqual = $true
                        
                        # First check if lengths match
                        if ($regValueData.Length -ne $currentValue.Length) {
                            $isEqual = $false
                        } else {
                            # Compare each byte individually
                            for ($i = 0; $i -lt $regValueData.Length; $i++) {
                                if ($regValueData[$i] -ne $currentValue[$i]) {
                                    $isEqual = $false
                                    break
                                }
                            }
                        }
                        
                        if (-not $isEqual) {
                            Write-Log "✗ Verification FAILED: Value data does not match expected." "WARNING"
                            Write-Log "  Expected: $($regValueData | ForEach-Object { '{0:X2}' -f $_ })" "DETAIL"
                            Write-Log "  Actual:   $($currentValue | ForEach-Object { '{0:X2}' -f $_ })" "DETAIL"
                            return $false
                        } else {
                            Write-Log "✓ Verification SUCCESSFUL: Value exists and data matches." "SUCCESS"
                            return $true
                        }
                    } else {
                        # For other types, use direct equality check
                        if ($regValueData -ne $currentValue) {
                            Write-Log "✗ Verification FAILED: Value does not match expected." "WARNING"
                            Write-Log "  Expected: $regValueData" "DETAIL"
                            Write-Log "  Actual:   $currentValue" "DETAIL"
                            return $false
                        } else {
                            Write-Log "✓ Verification SUCCESSFUL: Value exists and matches expected." "SUCCESS"
                            return $true
                        }
                    }
                }
            } catch {
                Write-Log "✗ Verification FAILED: Could not read value '$regValueName'. Error: $_" "ERROR"
                return $false
            }
        } else {
            Write-Log "✗ Verification FAILED: Registry key path does not exist: $verificationPath" "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "An error occurred during verification: $_" "ERROR"
        return $false
    }
} # End of Test-RegistryChanges

# Function alias with approved verb - redirects to Test-RegistryChanges
function Confirm-RegistryChanges {
    [Alias('Verify-RegistryChanges')]
    param (
        [Parameter(Mandatory = $false)]
        [string]$Username,

        [Parameter(Mandatory = $false)]
        [string]$LoadedHivePathForVerification # Pass the parameter through
    )

    # Call the properly named function, passing the new parameter along
    Test-RegistryChanges -Username $Username -LoadedHivePathForVerification $LoadedHivePathForVerification
} # End of Confirm-RegistryChanges

try {
    # Log script start
    Write-Log "Starting registry change application script" "INFO"
    Write-Log "Script version: 1.3.8" "DETAIL"
    
    # Check for Admin/SYSTEM privileges
    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($currentUser)
    $isAdmin = $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    Write-Log "Running as: $($currentUser.Name) (Elevated/SYSTEM: $isAdmin)" "DETAIL"
    if (-not $isAdmin) {
        throw "This script requires administrative privileges to modify registry hives."
    }

    # Define variables for target path construction
    $targetRegRootPath = $null
    $targetRegKeyPath = $null
    $userHiveLoadedForApply = $false
    $regExeFullKeyName = "HKLM\TempHive" # Full key path for reg.exe load/unload
    $psTempHiveRootPath = "Registry::$regExeFullKeyName" # Full path for PowerShell provider
    $userSID = $null
    $isUserLoggedIn = $false
    $applySuccess = $false # Flag to track if the apply step succeeded

    if ($Username) {
        Write-Log "Target user specified: $Username" "PROCESS"
        
        # Check if the user is logged in
        $isUserLoggedIn = Test-UserLoggedIn -Username $Username
        
        if ($isUserLoggedIn) {
            # User is logged in, target HKEY_USERS\<SID>
            $userSID = Get-UserSID -Username $Username
            if (-not $userSID) { throw "Could not get SID for logged-in user '$Username'." }
            
            # Use proper path format for PowerShell provider
            $targetRegRootPath = "Registry::HKEY_USERS\$userSID"
            $targetRegKeyPath = "$targetRegRootPath\$regKeyRelativePath"
            Write-Log "Targeting live user hive: $targetRegKeyPath" "PROCESS"
            
            # Apply change directly to HKEY_USERS
            if ($pscmdlet.ShouldProcess("registry key '$targetRegKeyPath'", "Set value '$regValueName'")) {
                # Ensure parent key exists
                if (-not (Test-Path -Path $targetRegKeyPath)) {
                    Write-Log "Creating parent key: $targetRegKeyPath" "DETAIL"
                    # Create parent keys recursively with proper path formatting
                    $null = New-Item -Path $targetRegRootPath -Name $regKeyRelativePath -Force
                }
                Set-ItemProperty -Path $targetRegKeyPath -Name $regValueName -Value $regValueData -Type $regValueType -Force
                Write-Log "Registry value '$regValueName' set successfully." "SUCCESS"
                $applySuccess = $true
            } else {
                 Write-Log "Skipped applying registry change due to -WhatIf." "INFO"
                 $applySuccess = $true # Assume success for WhatIf verification
            }
        } else {
            # User is not logged in, load NTUSER.DAT
            Write-Log "User not logged in. Attempting to load user hive..." "PROCESS"
            # Find user profile path and hive path
            $userProfile = Get-CimInstance -ClassName Win32_UserProfile | Where-Object { $_.LocalPath.EndsWith($Username) }
            if (-not $userProfile) { throw "User profile for '$Username' not found via WMI." }
            $userHivePath = Join-Path -Path $userProfile.LocalPath -ChildPath "NTUSER.DAT"
            if (-not (Test-Path -Path $userHivePath)) { throw "NTUSER.DAT not found at $userHivePath." }

            Write-Log "Loading user registry hive for $Username from $userHivePath into $regExeFullKeyName..." "PROCESS"
            if ($PSCmdlet.ShouldProcess("Registry Hive: $userHivePath", "Load into $regExeFullKeyName")) {
                 # Ensure TempHive is clear before loading
                if (Test-Path -Path $psTempHiveRootPath) { # Use PS path for Test-Path
                    Write-Log "Attempting to unload existing apply hive mount ($regExeFullKeyName)..." "DEBUG"
                    reg.exe unload $regExeFullKeyName 2>&1 | Out-Null # Use full reg.exe name
                    Start-Sleep -Seconds 1
                }
                $loadHiveOutput = reg.exe load $regExeFullKeyName $userHivePath 2>&1 # Use full reg.exe name
                if ($LASTEXITCODE -ne 0) { throw "Failed to load user registry hive into '$regExeFullKeyName': $loadHiveOutput" }
                $userHiveLoadedForApply = $true # Mark that hive was loaded for apply step
                Write-Log "User registry hive loaded successfully into $regExeFullKeyName" "SUCCESS"
            } else {
                 Write-Log "WhatIf: Would load registry hive $userHivePath into $regExeFullKeyName" "PROCESS"
                 # In WhatIf mode, we can't actually load, so we can't proceed to set value or verify
                 $applySuccess = $false # Mark apply as not successful in WhatIf for load scenario
            }
            # Construct target path using the PS provider path with proper backslash formatting
            $targetRegKeyPath = "$psTempHiveRootPath\$regKeyRelativePath"
            Write-Log "Targeting loaded user hive: $targetRegKeyPath" "PROCESS"
        } # End of if ($isUserLoggedIn)

        # Apply registry change using PowerShell cmdlets (only if hive load wasn't skipped by WhatIf)
        if ($isUserLoggedIn -or $userHiveLoadedForApply) {
            Write-Log "Applying registry change: Set '$regValueName' in '$targetRegKeyPath'" "PROCESS"
            if ($PSCmdlet.ShouldProcess($targetRegKeyPath, "Set registry value '$regValueName'")) {
                try {
                    # Ensure parent key exists
                    $parentPath = Split-Path -Path $targetRegKeyPath -Parent
                    if (-not (Test-Path -Path $parentPath)) {
                        Write-Log "Parent key '$parentPath' does not exist. Creating..." "DETAIL"
                        # Determine the root path and relative path for New-Item
                        if ($userHiveLoadedForApply) {
                            # For loaded hive, create key at root of mounted hive
                            $relativePath = $regKeyRelativePath
                            $null = New-Item -Path $psTempHiveRootPath -Name $relativePath -Force -ErrorAction Stop
                        } else {
                            # For logged-in user, create key at HKEY_USERS\SID
                            $relativePath = $regKeyRelativePath
                            $null = New-Item -Path $targetRegRootPath -Name $relativePath -Force -ErrorAction Stop
                        }
                        Write-Log "Parent key created." "DETAIL"
                    }

                    # Set the value
                    Set-ItemProperty -Path $targetRegKeyPath -Name $regValueName -Value $regValueData -Type $regValueType -Force -ErrorAction Stop
                    Write-Log "Registry value '$regValueName' set successfully." "SUCCESS"
                    $applySuccess = $true
                } catch {
                    Write-Log "Failed to set registry value '$regValueName' at '$targetRegKeyPath'. Error: $_" "ERROR"
                    $applySuccess = $false
                }
            } else {
                Write-Log "WhatIf: Would set registry value '$regValueName' at '$targetRegKeyPath'" "PROCESS"
                $applySuccess = $true # Assume success for WhatIf verification
            }
        }

    } else {
        # No username specified - Apply to current user (HKCU)
        # Use proper path format for PowerShell provider
        $targetRegRootPath = "Registry::HKEY_CURRENT_USER"
        $targetRegKeyPath = "$targetRegRootPath\$regKeyRelativePath"
        Write-Log "Targeting current user hive: $targetRegKeyPath" "PROCESS"
        
        # Apply change to HKCU
        if ($PSCmdlet.ShouldProcess("registry key '$targetRegKeyPath'", "Set value '$regValueName'")) {
             try {
                # Ensure parent key exists
                $parentPath = Split-Path -Path $targetRegKeyPath -Parent
                if (-not (Test-Path -Path $parentPath)) {
                    Write-Log "Parent key '$parentPath' does not exist. Creating..." "DETAIL"
                    $null = New-Item -Path $targetRegRootPath -Name $regKeyRelativePath -Force
                    Write-Log "Parent key created." "DETAIL"
                }
                # Set the value
                Set-ItemProperty -Path $targetRegKeyPath -Name $regValueName -Value $regValueData -Type $regValueType -Force
                Write-Log "Registry value '$regValueName' set successfully for current user." "SUCCESS"
                $applySuccess = $true
            } catch {
                Write-Log "Failed to set registry value '$regValueName' for current user at '$targetRegKeyPath'. Error: $_" "ERROR"
                throw "Failed to apply registry change for current user."
            }
        } else {
             Write-Log "WhatIf: Would set registry value '$regValueName' for current user at '$targetRegKeyPath'" "PROCESS"
        }
    } # End of if ($Username)

    # --- Verification Step ---
    # Only verify if the apply step was successful (or assumed successful via WhatIf)
    if ($applySuccess) {
        Write-Log "Starting verification of registry changes..." "PROCESS"
        $verificationResult = $false
        if ($userHiveLoadedForApply) {
            # Pass the loaded hive path for verification (format: HKLM\KeyName)
            $verificationHivePath = $regExeFullKeyName # Use the full key name HKLM\TempHive
            $verificationResult = Confirm-RegistryChanges -Username $Username -LoadedHivePathForVerification $verificationHivePath
        } else {
            # Verify logged-in user or current user normally
            $verificationResult = Confirm-RegistryChanges -Username $Username
        }

        if (-not $verificationResult) {
            Write-Log "Registry changes could not be verified. Please check the logs for details." "WARNING"
            # Consider throwing an error here if verification failure should halt the script or indicate overall failure
        }
    } else {
        Write-Log "Skipping verification because the apply step did not complete successfully or was skipped." "INFO"
    } # End of if ($applySuccess)

} catch {
    Write-Log "An error occurred: $($_.Exception.Message)" "ERROR"
    # Consider re-throwing the error if needed: throw $_
} finally {
    # Unload the hive if it was loaded during the apply step
    if ($userHiveLoadedForApply) {
        Write-Log "Unloading user registry hive ($regExeFullKeyName)..." "PROCESS"
        [System.GC]::Collect()
        Start-Sleep -Seconds 1
        $unloadOutput = reg.exe unload $regExeFullKeyName 2>&1 # Use full reg.exe name
        if ($LASTEXITCODE -ne 0) {
            Write-Log "Warning: Failed to unload user registry hive ($regExeFullKeyName): $unloadOutput" "WARNING"
        } else {
            Write-Log "User registry hive unloaded successfully." "SUCCESS"
        }
    } # End of if ($userHiveLoadedForApply)
    Write-Log "Script execution completed" "INFO"
} # End of try/catch/finally
