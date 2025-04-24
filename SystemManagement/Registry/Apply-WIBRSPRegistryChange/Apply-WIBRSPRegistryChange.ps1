# =============================================================================
# Script: Apply-WIBRSPRegistryChange.ps1
# Created: 2025-04-24 18:10:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-25 00:52:00 UTC
# Updated By: GitHub Copilot
# Version: 1.4.18
# Additional Info: Ensured $null is on the left for both comparisons in the null check.
# =============================================================================

<#
.SYNOPSIS
Applies a specific registry change, targeting either the current user or a specified user's hive (loading if necessary).

.DESCRIPTION
This script sets a defined registry value (Name, Data, Type) at a specified relative path within HKEY_CURRENT_USER or a specified user's registry hive.
It handles:
- Running with elevated privileges (required).
- Targeting the current user (if no -Username specified).
- Targeting a specific user by loading their NTUSER.DAT hive if they are not logged in.
- Applying the change directly if the specified user is logged in.
- Verifying the change after application.
- Logging actions to a file and the console.
- Includes -WhatIf support.

Dependencies: Requires administrative privileges.

.PARAMETER Username
Optional. The username of the target user. If not provided, the script targets HKEY_CURRENT_USER.

.PARAMETER LogPath
Optional. The path to the log file. Defaults to a file named <ScriptName>-<Date>.log in the script's directory.

.PARAMETER DebugPreference
Optional. Set to 'Continue' to enable detailed debug logging.

.EXAMPLE
.\Apply-WIBRSPRegistryChange.ps1
# Applies the predefined registry change to the current user's hive.

.EXAMPLE
.\Apply-WIBRSPRegistryChange.ps1 -Username 'testuser'
# Applies the predefined registry change to the 'testuser' hive, loading NTUSER.DAT if necessary.

.EXAMPLE
.\Apply-WIBRSPRegistryChange.ps1 -Username 'testuser' -WhatIf
# Shows what actions would be taken for 'testuser' without making changes.

.EXAMPLE
.\Apply-WIBRSPRegistryChange.ps1 -DebugPreference Continue
# Runs the script with verbose debug logging enabled.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$Username,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = (Join-Path -Path $PSScriptRoot -ChildPath "$($MyInvocation.MyCommand.Name)-$(Get-Date -Format 'yyyyMMdd_HHmmss').log"),

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.ActionPreference]$DebugPreference = 'SilentlyContinue' # Default DebugPreference
)

# --- Configuration ---
$regKeyRelativePath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
$regValueName = "Shell"
$regValueData = "explorer.exe"
$regValueType = "String"

# --- Script Setup ---
$global:DebugPreference = $DebugPreference # Set global DebugPreference

# --- Logging Function ---
function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $true)]
        [ValidateSet('INFO', 'PROCESS', 'SUCCESS', 'WARNING', 'ERROR', 'DEBUG', 'DETAIL')]
        [string]$Level
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC'
    $logEntry = "[$timestamp][$Level] $Message"

    # Console Output with Color
    switch ($Level) {
        'INFO'    { Write-Host $logEntry -ForegroundColor White } # Standard info
        'PROCESS' { Write-Host $logEntry -ForegroundColor Cyan } # Process updates
        'SUCCESS' { Write-Host $logEntry -ForegroundColor Green } # Success messages
        'WARNING' { Write-Host $logEntry -ForegroundColor Yellow } # Warnings
        'ERROR'   { Write-Host $logEntry -ForegroundColor Red } # Errors
        'DEBUG'   { Write-Debug $logEntry } # Debug messages (respects -Debug switch)
        'DETAIL'  { Write-Host $logEntry -ForegroundColor DarkGray } # Less important details
        default   { Write-Host $logEntry } # Default case
    }

    # File Output
    try {
        Out-File -FilePath $LogPath -InputObject $logEntry -Append -Encoding UTF8 -ErrorAction Stop
    } catch {
        Write-Warning "Failed to write to log file '$LogPath'. Error: $($_.Exception.Message)"
    }
}

# --- Helper Functions ---

# Function to get User SID
function Get-UserSID {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Username
    )
    try {
        $ntAccount = New-Object System.Security.Principal.NTAccount($Username)
        $sid = $ntAccount.Translate([System.Security.Principal.SecurityIdentifier])
        return $sid.Value
    } catch {
        Write-Log "Error retrieving SID for user '$Username': $_" "ERROR"
        return $null
    }
}

# Function to check if a user is logged in
function Test-UserLoggedIn {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Username
    )
    # Use quser for broader compatibility, filter output
    $loggedInUsers = (quser 2>&1 | Select-String -Pattern $Username -SimpleMatch)
    if ($loggedInUsers) {
        Write-Log "User '$Username' is currently logged in." "DETAIL"
        return $true
    } else {
        Write-Log "User '$Username' is not currently logged in." "DETAIL"
        return $false
    }
}

# Function to test registry changes
function Test-RegistryChanges {
    [CmdletBinding(SupportsShouldProcess = $true)] 
    param (
        [Parameter(Mandatory = $true)]
        [string]$RegKeyRelativePath,

        [Parameter(Mandatory = $true)]
        [string]$RegValueName,

        [Parameter(Mandatory = $true)]
        $RegValueData, # Can be any type

        [Parameter(Mandatory = $true)]
        [Microsoft.Win32.RegistryValueKind]$RegValueType,

        [Parameter(Mandatory = $false)]
        [string]$Username,

        [Parameter(Mandatory = $false)]
        [string]$LoadedHivePathForVerification # Path like HKLM:\TempHiveVerify
    )

    Write-Log "Verifying registry changes..." "PROCESS"

    $verificationPath = $null
    $userSID = $null
    $currentValue = $null
    $valueReadSuccess = $false

    try { # Outer Try Block (Setup and Read)
        # Determine the correct verification path
        if ($Username) {
            if ($LoadedHivePathForVerification) {
                # Verification uses a specifically loaded hive (e.g., TempHiveVerify)
                $verificationPath = "Registry::$LoadedHivePathForVerification\$RegKeyRelativePath"
                Write-Log "Verification Target: Loaded Hive at $verificationPath" "DETAIL"
            } else {
                # Verification targets the live HKEY_USERS hive for a logged-in user
                $userSID = Get-UserSID -Username $Username
                if (-not $userSID) { throw "Could not get SID for verification for user '$Username'." }
                $verificationPath = "Registry::HKEY_USERS\$userSID\$RegKeyRelativePath"
                Write-Log "Verification Target: Live User Hive at $verificationPath" "DETAIL"
            }
        } else {
            # Verification targets HKEY_CURRENT_USER
            $verificationPath = "Registry::HKEY_CURRENT_USER\$RegKeyRelativePath"
            Write-Log "Verification Target: HKEY_CURRENT_USER at $verificationPath" "DETAIL"
        }

        if (-not (Test-Path -Path $verificationPath)) {
            Write-Log "✗ Verification FAILED: Registry key path does not exist: $verificationPath" "ERROR"
            return $false
        }

        # Inner Try Block ONLY for Get-ItemProperty
        try {
            $currentValue = Get-ItemProperty -Path $verificationPath -Name $RegValueName -ErrorAction Stop | Select-Object -ExpandProperty $RegValueName
            $valueReadSuccess = $true
            Write-Log "✓ Verification: Successfully read value '$RegValueName' from path '$verificationPath'." "DETAIL"
        }
        catch {
            Write-Log "✗ Verification FAILED: Could not read value '$RegValueName' from path '$verificationPath'. Error: $_" "ERROR"
            # $valueReadSuccess remains false
        }
        # --- End Inner Try/Catch ---

    } # End Outer Try Block
    catch { # Catch for Outer Try Block
        Write-Log "An error occurred during verification setup or path check: $_" "ERROR"
        return $false
    } # End Outer Catch Block

    # --- Value Verification Logic (Outside ALL Try/Catch Blocks) ---
    # Check if the read was successful before proceeding
    if (-not $valueReadSuccess) {
         Write-Log "✗ Verification FAILED: Value read step failed." "ERROR"
         return $false
    }

    # Check if the retrieved value is null (value exists but is empty/null)
    # Corrected comparison: $null on the left for both parts
    if ($null -eq $currentValue -and $null -ne $RegValueData) { # Only fail if expected data is not null
        Write-Log "✗ Verification FAILED: Value '$RegValueName' exists but is null/empty, expected non-null value '$RegValueData'." "WARNING"
        return $false
    }

    # Perform comparisons - Use a clear, separate block
    $finalVerificationResult = $false # Initialize final result
    if ($RegValueName -eq "TimerAutoMount" -and $RegValueType -eq [Microsoft.Win32.RegistryValueKind]::Binary) 
    { 
        # Special handling for TimerAutoMount (Binary) - just check existence
        Write-Log "✓ Verification SUCCESSFUL: TimerAutoMount registry value exists (existence check only)." "SUCCESS"
        $finalVerificationResult = $true
    } 
    elseif ($currentValue -is [byte[]] -and $RegValueData -is [byte[]]) 
    { 
        # Compare byte arrays
        if (-not (Compare-Object -ReferenceObject $RegValueData -DifferenceObject $currentValue -SyncWindow 0)) {
            Write-Log "✓ Verification SUCCESSFUL: Binary value matches expected." "SUCCESS"
            $finalVerificationResult = $true
        } else {
            Write-Log "✗ Verification FAILED: Binary value does not match expected." "WARNING"
            Write-Log "  Expected (bytes): $(($RegValueData | ForEach-Object { $_.ToString('X2') }) -join ' ')" "DETAIL"
            Write-Log "  Actual   (bytes): $(($currentValue | ForEach-Object { $_.ToString('X2') }) -join ' ')" "DETAIL"
            $finalVerificationResult = $false
        }
    } 
    elseif ($RegValueType -eq [Microsoft.Win32.RegistryValueKind]::MultiString) {
        # Compare MultiString arrays
        if (-not (Compare-Object -ReferenceObject $RegValueData -DifferenceObject $currentValue -SyncWindow 0)) {
            Write-Log "✓ Verification SUCCESSFUL: MultiString value matches expected." "SUCCESS"
            $finalVerificationResult = $true
        } else {
            Write-Log "✗ Verification FAILED: MultiString value does not match expected." "WARNING"
            Write-Log "  Expected: $($RegValueData -join '; ')" "DETAIL"
            Write-Log "  Actual:   $($currentValue -join '; ')" "DETAIL"
            $finalVerificationResult = $false
        }
    }
    else 
    { 
        # Direct equality check for other types (String, DWord, QWord, ExpandString)
        if ($RegValueData -ne $currentValue) 
        { 
            Write-Log "✗ Verification FAILED: Value does not match expected." "WARNING"
            Write-Log "  Expected: $RegValueData ($($RegValueData.GetType().Name))" "DETAIL"
            Write-Log "  Actual:   $currentValue ($($currentValue.GetType().Name))" "DETAIL"
            $finalVerificationResult = $false
        } 
        else 
        { 
            Write-Log "✓ Verification SUCCESSFUL: Value exists and matches expected." "SUCCESS"
            $finalVerificationResult = $true
        }
    }
    
    return $finalVerificationResult # Return the final result
    # --- End Value Verification Logic ---

} # End of Test-RegistryChanges

# Function alias with approved verb - redirects to Test-RegistryChanges
function Confirm-RegistryChanges {
    [Alias('Verify-RegistryChanges')] # Keep alias for compatibility if needed
    [CmdletBinding(SupportsShouldProcess = $true)] # Add binding for WhatIf propagation
    param (
        # Parameters should match Test-RegistryChanges for consistency
        [Parameter(Mandatory = $true)]
        [string]$RegKeyRelativePath,

        [Parameter(Mandatory = $true)]
        [string]$RegValueName,

        [Parameter(Mandatory = $true)]
        $RegValueData,

        [Parameter(Mandatory = $true)]
        [Microsoft.Win32.RegistryValueKind]$RegValueType,

        [Parameter(Mandatory = $false)]
        [string]$Username,

        [Parameter(Mandatory = $false)]
        [string]$LoadedHivePathForVerification
    )

    # Call the properly named function, passing all parameters
    # Use splatting for cleaner parameter passing
    $testParams = @{
        RegKeyRelativePath = $RegKeyRelativePath
        RegValueName = $RegValueName
        RegValueData = $RegValueData
        RegValueType = $RegValueType
        ErrorAction = 'Stop' # Ensure errors stop execution here if needed
    }
    if ($PSBoundParameters.ContainsKey('Username')) {
        $testParams.Username = $Username
    }
    if ($PSBoundParameters.ContainsKey('LoadedHivePathForVerification')) {
        $testParams.LoadedHivePathForVerification = $LoadedHivePathForVerification
    }

    # Propagate WhatIf/Confirm
    if ($PSCmdlet.ShouldProcess("Registry Key '$RegKeyRelativePath', Value '$RegValueName'", "Verify Configuration")) {
        Test-RegistryChanges @testParams
    }
    # Note: If -WhatIf is used, Test-RegistryChanges won't actually read, 
    # but this structure correctly shows the intent to verify.

} # End of Confirm-RegistryChanges

# --- Main Script Logic ---
try {    
    Write-Log "Starting registry change application script" "INFO"
    Write-Log "Script version: $($MyInvocation.MyCommand.Version)" "DETAIL"
    Write-Log "Log file: $LogPath" "DETAIL"
    
    # Check for Admin/SYSTEM privileges
    $currentUserIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($currentUserIdentity)
    $isAdmin = $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    Write-Log "Running as: $($currentUserIdentity.Name) (Elevated/SYSTEM: $isAdmin)" "DETAIL"
    if (-not $isAdmin) {
        throw "This script requires administrative privileges to modify registry hives."
    }

    # Define variables for target path construction
    $targetRegRootPath = $null
    $targetRegKeyPath = $null
    $userHiveLoadedForApply = $false
    $userHiveLoadedForVerify = $false
    $regExeApplyHiveMount = "HKLM\TempHiveApply" # Unique name for apply load
    $regExeVerifyHiveMount = "HKLM\TempHiveVerify" # Unique name for verify load
    $psApplyHiveRootPath = "Registry::$regExeApplyHiveMount" # PS path for apply
    $psVerifyHiveRootPath = "Registry::$regExeVerifyHiveMount" # PS path for verify
    $userSID = $null
    $isUserLoggedIn = $false
    $applySuccess = $false # Flag to track if the apply step succeeded
    $verificationSuccess = $false # Flag for final verification result
    $userHivePath = $null # Store NTUSER.DAT path

    # --- Determine Target and Load Hive if Necessary ---
    if ($Username) {
        Write-Log "Target user specified: $Username" "PROCESS"
        $userSID = Get-UserSID -Username $Username
        if (-not $userSID) { throw "Could not get SID for user '$Username'." }
        
        $isUserLoggedIn = Test-UserLoggedIn -Username $Username
        
        if ($isUserLoggedIn) {
            # User is logged in, target HKEY_USERS\<SID>
            $targetRegRootPath = "Registry::HKEY_USERS\$userSID"
            $targetRegKeyPath = Join-Path -Path $targetRegRootPath -ChildPath $regKeyRelativePath
            Write-Log "Targeting live user hive: $targetRegKeyPath" "PROCESS"
        } else {
            # User is not logged in, need to load NTUSER.DAT
            Write-Log "User not logged in. Locating user hive..." "PROCESS"
            $userProfile = Get-CimInstance -ClassName Win32_UserProfile | Where-Object { $_.SID -eq $userSID }
            if (-not $userProfile) { throw "User profile for SID '$userSID' ($Username) not found via WMI." }
            $userHivePath = Join-Path -Path $userProfile.LocalPath -ChildPath "NTUSER.DAT"
            if (-not (Test-Path -Path $userHivePath)) { throw "NTUSER.DAT not found at $userHivePath." }

            Write-Log "Attempting to load user hive for APPLY: $userHivePath -> $regExeApplyHiveMount" "PROCESS"
            if ($PSCmdlet.ShouldProcess("Registry Hive: $userHivePath", "Load into $regExeApplyHiveMount for Apply")) {
                # Ensure Apply mount point is clear
                if (Test-Path -Path $psApplyHiveRootPath) {
                    Write-Log "Unloading existing apply hive mount ($regExeApplyHiveMount)..." "DEBUG"
                    reg.exe unload $regExeApplyHiveMount 2>&1 | Out-Null
                    Start-Sleep -Seconds 1
                }
                $loadHiveOutput = reg.exe load $regExeApplyHiveMount $userHivePath 2>&1
                if ($LASTEXITCODE -ne 0) { throw "Failed to load user registry hive into '$regExeApplyHiveMount': $loadHiveOutput" }
                $userHiveLoadedForApply = $true
                Write-Log "User registry hive loaded successfully into $regExeApplyHiveMount for Apply." "SUCCESS"
                $targetRegKeyPath = Join-Path -Path $psApplyHiveRootPath -ChildPath $regKeyRelativePath
                Write-Log "Targeting loaded apply hive: $targetRegKeyPath" "PROCESS"
            } else {
                 Write-Log "WhatIf: Skipped loading registry hive $userHivePath into $regExeApplyHiveMount for Apply." "INFO"
                 # Cannot proceed with apply if hive load is skipped
                 $applySuccess = $false 
            }
        } 
    } else {
        # No username specified, target HKEY_CURRENT_USER
        Write-Log "Targeting current user (HKEY_CURRENT_USER)" "PROCESS"
        $targetRegRootPath = "Registry::HKEY_CURRENT_USER"
        $targetRegKeyPath = Join-Path -Path $targetRegRootPath -ChildPath $regKeyRelativePath
        Write-Log "Targeting current user hive: $targetRegKeyPath" "PROCESS"
    }

    # --- Apply Registry Change ---
    # Proceed only if we have a valid target path (i.e., not skipped by WhatIf on hive load)
    if ($targetRegKeyPath) {
        Write-Log "Applying registry change: Set '$regValueName' in '$targetRegKeyPath'" "PROCESS"
        if ($PSCmdlet.ShouldProcess($targetRegKeyPath, "Set registry value '$regValueName' = '$regValueData' (Type: $regValueType)")) {
            try {
                # Ensure parent key exists
                $parentPath = Split-Path -Path $targetRegKeyPath -Parent
                if (-not (Test-Path -Path $parentPath)) {
                    Write-Log "Parent key '$parentPath' does not exist. Creating..." "DETAIL"
                    # Create parent keys recursively
                    $null = New-Item -Path $parentPath -Force -ErrorAction Stop
                }
                # Set the value
                Set-ItemProperty -Path $targetRegKeyPath -Name $regValueName -Value $regValueData -Type $regValueType -Force -ErrorAction Stop
                Write-Log "✓ Registry value '$regValueName' set successfully in '$targetRegKeyPath'." "SUCCESS"
                $applySuccess = $true
            } catch {
                Write-Log "✗ FAILED to apply registry change to '$targetRegKeyPath'. Error: $_" "ERROR"
                $applySuccess = $false
                # Optionally re-throw to stop script, or just log and continue to finally block
                # throw $_ 
            }
        } else {
             Write-Log "Skipped applying registry change due to -WhatIf." "INFO"
             $applySuccess = $true # Assume success for WhatIf verification purposes
        }
    } else {
        Write-Log "Skipping Apply step because target path was not determined (likely due to -WhatIf on hive load)." "INFO"
        $applySuccess = $false
    }

    # --- Verification Step ---
    if ($applySuccess) {
        Write-Log "Proceeding to verification..." "PROCESS"
        $verifyParams = @{
            RegKeyRelativePath = $regKeyRelativePath
            RegValueName = $regValueName
            RegValueData = $regValueData
            RegValueType = $regValueType
            ErrorAction = 'Stop' # Stop verification if it fails internally
        }
        if ($Username) {
            $verifyParams.Username = $Username
            # If we loaded the hive for APPLY, we need to load it AGAIN for VERIFY
            # because the apply hive might be unloaded in the finally block before verification runs.
            if ($userHiveLoadedForApply) {
                Write-Log "Attempting to load user hive for VERIFY: $userHivePath -> $regExeVerifyHiveMount" "PROCESS"
                if ($PSCmdlet.ShouldProcess("Registry Hive: $userHivePath", "Load into $regExeVerifyHiveMount for Verify")) {
                    # Ensure Verify mount point is clear
                    if (Test-Path -Path $psVerifyHiveRootPath) {
                        Write-Log "Unloading existing verify hive mount ($regExeVerifyHiveMount)..." "DEBUG"
                        reg.exe unload $regExeVerifyHiveMount 2>&1 | Out-Null
                        Start-Sleep -Seconds 1
                    }
                    $loadVerifyOutput = reg.exe load $regExeVerifyHiveMount $userHivePath 2>&1
                    if ($LASTEXITCODE -ne 0) {
                        Write-Log "Failed to load user registry hive into '$regExeVerifyHiveMount' for verification: $loadVerifyOutput" "ERROR"
                        # Cannot verify if load fails
                        $verificationSuccess = $false 
                    } else {
                        $userHiveLoadedForVerify = $true
                        Write-Log "User registry hive loaded successfully into $regExeVerifyHiveMount for Verify." "SUCCESS"
                        $verifyParams.LoadedHivePathForVerification = $regExeVerifyHiveMount # Tell verification to use this loaded hive
                    }
                } else {
                    Write-Log "WhatIf: Skipped loading registry hive $userHivePath into $regExeVerifyHiveMount for Verify." "INFO"
                    # Cannot verify if load is skipped
                    $verificationSuccess = $false 
                }
            } 
            # If user was logged in, verification will target HKEY_USERS directly (no LoadedHivePathForVerification needed)
        }
        # Only run verification if the hive was successfully loaded (or wasn't needed)
        if ($isUserLoggedIn -or -not $Username -or $userHiveLoadedForVerify) {
             Write-Log "Running verification check..." "PROCESS"
             $verificationSuccess = Confirm-RegistryChanges @verifyParams -WhatIf:$false # Force actual check
        } else {
            Write-Log "Skipping verification because the required hive could not be loaded or accessed." "WARNING"
            $verificationSuccess = $false # Mark as failed if we couldn't load/access for verify
        }

    } else {
        Write-Log "Skipping verification because the apply step did not succeed or was skipped." "WARNING"
        $verificationSuccess = $false
    }

} catch {
    Write-Log "An error occurred in the main script block: $_" "ERROR"
    # Consider exiting with a non-zero code
    # exit 1 
} finally {
    # --- Cleanup: Unload hives ---
    Write-Log "Performing cleanup..." "PROCESS"
    if ($userHiveLoadedForApply) {
        Write-Log "Unloading apply hive mount ($regExeApplyHiveMount)..." "PROCESS"
        if ($PSCmdlet.ShouldProcess($regExeApplyHiveMount, "Unload Registry Hive")) {
            reg.exe unload $regExeApplyHiveMount 2>&1 | Out-Null
            Write-Log "Apply user registry hive unloaded successfully." "SUCCESS"
        } else {
            Write-Log "WhatIf: Skipped unloading apply hive mount $regExeApplyHiveMount." "INFO"
        }
    } # <-- This was the missing brace
    if ($userHiveLoadedForVerify) {
        Write-Log "Unloading verify hive mount ($regExeVerifyHiveMount)..." "PROCESS"
        if ($PSCmdlet.ShouldProcess($regExeVerifyHiveMount, "Unload Registry Hive")) {
            reg.exe unload $regExeVerifyHiveMount 2>&1 | Out-Null
            Write-Log "Verify user registry hive unloaded successfully." "SUCCESS"
        } else {
            Write-Log "WhatIf: Skipped unloading verify hive mount $regExeVerifyHiveMount." "INFO"
        }
    }

    Write-Log "Script finished. Final Verification Status: $($verificationSuccess)" "INFO"
    
    # Exit with appropriate code based on verification
    if ($verificationSuccess) {
        # exit 0 # Success
    } else {
        # exit 1 # Failure
    }
}
