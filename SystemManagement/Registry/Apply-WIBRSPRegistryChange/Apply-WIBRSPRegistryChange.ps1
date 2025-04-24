# =============================================================================
# Script: Apply-WIBRSPRegistryChange.ps1
# Created: 2025-04-24 18:10:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-24 19:41:00 UTC
# Updated By: GitHub Copilot / jdyer-nuvodia
# Version: 1.3.0
# Additional Info: Added logic to handle logged-in users by targeting HKEY_USERS directly.
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
    [string]$Username
)

# Set error action preference
$ErrorActionPreference = "Stop"

# Define log file path in the same directory as script
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$logFile = Join-Path -Path $scriptPath -ChildPath "Apply-WIBRSPRegistryChange_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Define specific registry values to apply
$regKeyRelativePath = "Software\\Microsoft\\OneDrive\\Accounts\\Business1"
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

# Function to check if user is logged in using quser
function Test-UserLoggedIn {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Username
    )
    Write-Log "Checking login status for user '$Username' using quser..." "DETAIL"
    $quserOutput = try {
        quser 2>&1 | Out-String
    } catch {
        Write-Log "Failed to execute quser. Error: $_" "WARNING"
        # Assume user is not logged in if quser fails (conservative approach)
        return $false
    }

    if ($quserOutput -match ">$($Username)\s+") { # Match username at start of line after '>'
        Write-Log "User '$Username' appears to be logged in." "PROCESS"
        return $true
    } else {
        Write-Log "User '$Username' does not appear to be logged in." "PROCESS"
        return $false
    }
}


# Function to test registry changes
function Test-RegistryChanges {
    param (
        [Parameter(Mandatory = $false)]
        [string]$Username
    )
    
    Write-Log "Verifying registry changes..." "PROCESS"
    
    $verificationPath = $null
    $userHiveLoadedForVerify = $false
    $verifyTempHiveKeyName = "HKLM\VerifyTempHive" # Temporary mount point for verification only
    $userSID = $null

    try {
        if ($Username) {
            $isUserLoggedInVerify = Test-UserLoggedIn -Username $Username

            if ($isUserLoggedInVerify) {
                # User is logged in, check HKEY_USERS\<SID>
                $userSID = Get-UserSID -Username $Username
                if (-not $userSID) { return $false } # Get-UserSID throws on failure, but double-check
                $verificationPath = "Registry::HKEY_USERS\$userSID\$regKeyRelativePath"
                Write-Log "Verifying logged-in user's live hive path: $verificationPath" "DETAIL"
            } else {
                # User is not logged in, load hive temporarily for verification
                Write-Log "User not logged in. Loading hive for verification..." "DETAIL"
                $allProfiles = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*" |
                               Where-Object { $_.ProfileImagePath -like "*\$Username" }
                if (-not $allProfiles) { Write-Log "User profile for '$Username' not found." "ERROR"; return $false }
                $userProfilePath = $allProfiles.ProfileImagePath
                $userHivePath = Join-Path -Path $userProfilePath -ChildPath "NTUSER.DAT"
                if (-not (Test-Path -Path $userHivePath)) { Write-Log "NTUSER.DAT not found at $userHivePath." "ERROR"; return $false }

                # Ensure VerifyTempHive is clear before loading
                if (Test-Path -Path "Registry::$verifyTempHiveKeyName") {
                    Write-Log "Attempting to unload existing verification hive mount..." "DEBUG"
                    reg.exe unload $verifyTempHiveKeyName 2>&1 | Out-Null
                    Start-Sleep -Seconds 1
                }

                $loadHiveOutput = reg.exe load $verifyTempHiveKeyName $userHivePath 2>&1
                if ($LASTEXITCODE -ne 0) { Write-Log "Failed to load user hive for verification: $loadHiveOutput" "ERROR"; return $false }
                $userHiveLoadedForVerify = $true
                $verificationPath = "Registry::$verifyTempHiveKeyName\$regKeyRelativePath"
                Write-Log "Verifying loaded hive path: $verificationPath" "DETAIL"
            }
        } else {
            # No username specified, check current user (HKCU)
            $verificationPath = "Registry::HKEY_CURRENT_USER\$regKeyRelativePath"
            Write-Log "Verifying current user path: $verificationPath" "DETAIL"
        }

        # Perform the actual check
        if (Test-Path -Path $verificationPath) {
            $item = Get-ItemProperty -Path $verificationPath -Name $regValueName -ErrorAction SilentlyContinue
            if ($item) {
                $currentValue = $item.$regValueName
                if ($currentValue -is [byte[]] -and $currentValue.Length -eq $regValueData.Length) {
                    if (Compare-Object -ReferenceObject $regValueData -DifferenceObject $currentValue -SyncWindow 0 -Property @{Expression={$PSItem}}) {
                        Write-Log "✗ Verification FAILED: Value data does not match expected." "WARNING"
                        Write-Log "  Expected: $($regValueData | ForEach-Object { '{0:X2}' -f $_ })" "DETAIL"
                        Write-Log "  Actual:   $($currentValue | ForEach-Object { '{0:X2}' -f $_ })" "DETAIL"
                        return $false
                    } else {
                        Write-Log "✓ Verification SUCCESSFUL: Value exists and data matches." "SUCCESS"
                        return $true
                    }
                } else {
                    Write-Log "✗ Verification FAILED: Value exists but type or length is incorrect." "WARNING"
                    Write-Log "  Actual Type: $($currentValue.GetType().Name)" "DETAIL"
                    Write-Log "  Actual Length: $($currentValue.Length)" "DETAIL"
                    return $false
                }
            } else {
                Write-Log "✗ Verification FAILED: Value '$regValueName' does not exist in the key '$verificationPath'." "ERROR"
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
    finally {
        # Unload the hive if it was loaded for verification
        if ($userHiveLoadedForVerify) {
            Write-Log "Unloading verification registry hive ($verifyTempHiveKeyName)..." "DETAIL"
            [System.GC]::Collect()
            Start-Sleep -Seconds 1
            $unloadVerifyOutput = reg.exe unload $verifyTempHiveKeyName 2>&1
            if ($LASTEXITCODE -ne 0) {
                Write-Log "Warning: Failed to unload verification registry hive: $unloadVerifyOutput" "WARNING"
            }
        }
    }
}

# Function alias with approved verb - redirects to Test-RegistryChanges
function Confirm-RegistryChanges {
    [Alias('Verify-RegistryChanges')]
    param (
        [Parameter(Mandatory = $false)]
        [string]$Username
    )
    
    # Call the properly named function
    Test-RegistryChanges -Username $Username
}

try {
    # Log script start
    Write-Log "Starting registry change application script" "INFO"
    Write-Log "Script version: 1.3.0" "DETAIL"

    # Check if running as elevated/SYSTEM (needed for HKEY_USERS or reg load)
    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $isAdmin = (New-Object System.Security.Principal.WindowsPrincipal $currentUser).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    $isSystem = $currentUser.IsSystem
    if (-not ($isAdmin -or $isSystem)) {
         Write-Log "This script requires elevated privileges (Administrator or SYSTEM) to modify other users' registry hives." "ERROR"
         throw "Insufficient privileges."
    }
    Write-Log "Running as: $($currentUser.Name) (Elevated/SYSTEM: $($isAdmin -or $isSystem))" "DETAIL"

    # Define variables for target path construction
    $targetRegRootPath = $null
    $targetRegKeyPath = $null
    $userHiveLoadedForApply = $false
    $tempHiveKeyName = "HKLM\TempHive" # Temporary mount point for applying changes only
    $userSID = $null
    $isUserLoggedIn = $false

    if ($Username) {
        Write-Log "Target user specified: $Username" "PROCESS"

        # Check if user is logged in
        $isUserLoggedIn = Test-UserLoggedIn -Username $Username

        if ($isUserLoggedIn) {
            # User is logged in, target HKEY_USERS\<SID>
            $userSID = Get-UserSID -Username $Username
            # Get-UserSID throws on failure, script would exit
            $targetRegRootPath = "Registry::HKEY_USERS\$userSID"
            $targetRegKeyPath = Join-Path -Path $targetRegRootPath -ChildPath $regKeyRelativePath
            Write-Log "Targeting live user hive: $targetRegKeyPath" "PROCESS"
        } else {
            # User is not logged in, load NTUSER.DAT
            Write-Log "User not logged in. Attempting to load user hive..." "PROCESS"
            # Find user profile path
            $allProfiles = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*" |
                           Where-Object { $_.ProfileImagePath -like "*\$Username" }
            if (-not $allProfiles) { throw "User profile for '$Username' not found." }
            $userProfilePath = $allProfiles.ProfileImagePath
            $userHivePath = Join-Path -Path $userProfilePath -ChildPath "NTUSER.DAT"
            if (-not (Test-Path -Path $userHivePath)) { throw "User registry hive (NTUSER.DAT) not found at '$userHivePath'." }

            Write-Log "Loading user registry hive for $Username from $userHivePath into $tempHiveKeyName..." "PROCESS"
            if ($PSCmdlet.ShouldProcess("Registry", "Load user hive from $userHivePath into $tempHiveKeyName")) {
                 # Ensure TempHive is clear before loading
                if (Test-Path -Path "Registry::$tempHiveKeyName") {
                    Write-Log "Attempting to unload existing apply hive mount..." "DEBUG"
                    reg.exe unload $tempHiveKeyName 2>&1 | Out-Null
                    Start-Sleep -Seconds 1
                }
                $loadHiveOutput = reg.exe load $tempHiveKeyName $userHivePath 2>&1
                if ($LASTEXITCODE -ne 0) { throw "Failed to load user registry hive: $loadHiveOutput" }
                $userHiveLoadedForApply = $true
                Write-Log "User registry hive loaded successfully into $tempHiveKeyName" "SUCCESS"
            } else {
                 Write-Log "WhatIf: Would load user registry hive from: $userHivePath into $tempHiveKeyName" "PROCESS"
                 # In WhatIf mode, we can't proceed further with modification for offline user
                 Write-Log "WhatIf: Skipping modification as hive loading was skipped." "PROCESS"
                 # Exit cleanly in WhatIf for offline user if load is skipped
                 return
            }
            $targetRegRootPath = "Registry::$tempHiveKeyName"
            $targetRegKeyPath = Join-Path -Path $targetRegRootPath -ChildPath $regKeyRelativePath
            Write-Log "Targeting loaded user hive: $targetRegKeyPath" "PROCESS"
        }

        # Apply registry change using PowerShell cmdlets
        Write-Log "Applying registry change: Set '$regValueName' in '$targetRegKeyPath'" "PROCESS"
        if ($PSCmdlet.ShouldProcess($targetRegKeyPath, "Set registry value '$regValueName'")) {
            try {
                # Ensure parent key exists
                $parentPath = Split-Path -Path $targetRegKeyPath
                if (-not (Test-Path -Path $parentPath)) {
                    Write-Log "Parent key '$parentPath' does not exist. Creating..." "DETAIL"
                    New-Item -Path $parentPath -Force -ErrorAction Stop | Out-Null
                    Write-Log "Parent key created." "DETAIL"
                }

                # Set the value
                Set-ItemProperty -Path $targetRegKeyPath -Name $regValueName -Value $regValueData -Type $regValueType -Force -ErrorAction Stop
                Write-Log "Registry value '$regValueName' set successfully." "SUCCESS"
            } catch {
                Write-Log "Failed to set registry value '$regValueName' at '$targetRegKeyPath'. Error: $_" "ERROR"
                throw "Failed to apply registry change." # Throw to ensure cleanup runs if needed
            }
        } else {
            Write-Log "WhatIf: Would set registry value '$regValueName' at '$targetRegKeyPath'" "PROCESS"
        }

    } else {
        # No username specified - Apply to current user (HKCU)
        $targetRegRootPath = "Registry::HKEY_CURRENT_USER"
        $targetRegKeyPath = Join-Path -Path $targetRegRootPath -ChildPath $regKeyRelativePath
        Write-Log "Targeting current user hive: $targetRegKeyPath" "PROCESS"

        Write-Log "Applying registry change: Set '$regValueName' in '$targetRegKeyPath'" "PROCESS"
        if ($PSCmdlet.ShouldProcess($targetRegKeyPath, "Set registry value '$regValueName'")) {
             try {
                # Ensure parent key exists
                $parentPath = Split-Path -Path $targetRegKeyPath
                if (-not (Test-Path -Path $parentPath)) {
                    Write-Log "Parent key '$parentPath' does not exist. Creating..." "DETAIL"
                    New-Item -Path $parentPath -Force -ErrorAction Stop | Out-Null
                    Write-Log "Parent key created." "DETAIL"
                }
                # Set the value
                Set-ItemProperty -Path $targetRegKeyPath -Name $regValueName -Value $regValueData -Type $regValueType -Force -ErrorAction Stop
                Write-Log "Registry value '$regValueName' set successfully for current user." "SUCCESS"
            } catch {
                Write-Log "Failed to set registry value '$regValueName' for current user at '$targetRegKeyPath'. Error: $_" "ERROR"
                throw "Failed to apply registry change for current user."
            }
        } else {
             Write-Log "WhatIf: Would set registry value '$regValueName' for current user at '$targetRegKeyPath'" "PROCESS"
        }
    }

    # Verify the registry changes if not in WhatIf mode
    if (-not $WhatIfPreference) {
        Write-Log "Starting verification of registry changes..." "PROCESS"
        $verificationResult = Test-RegistryChanges -Username $Username
        
        if ($verificationResult) {
            Write-Log "Registry changes have been successfully verified!" "SUCCESS"
        } else {
            Write-Log "Registry changes could not be verified. Please check the logs for details." "WARNING"
        }
    }
}
catch {
    Write-Log "An error occurred: $_" "ERROR"
    Write-Log "Exception details: $($_.Exception)" "DEBUG"
    exit 1
}
finally {
    # Clean up - unload hive if it was loaded for applying changes
    if ($userHiveLoadedForApply) {
        Write-Log "Unloading user registry hive ($tempHiveKeyName)..." "PROCESS"
        if ($PSCmdlet.ShouldProcess("Registry", "Unload user hive from $tempHiveKeyName")) {
            [System.GC]::Collect()
            Start-Sleep -Seconds 1
            $unloadApplyOutput = reg.exe unload $tempHiveKeyName 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Log "User registry hive unloaded successfully." "SUCCESS"
            } else {
                Write-Log "Warning: Failed to unload user registry hive ($tempHiveKeyName). Exit code: $LASTEXITCODE" "WARNING"
                Write-Log "Error details: $unloadApplyOutput" "WARNING"
            }
        } else {
            Write-Log "WhatIf: Would unload user registry hive from $tempHiveKeyName" "PROCESS"
        }
    }

    Write-Log "Script execution completed" "INFO"
}
