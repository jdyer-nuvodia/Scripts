# =============================================================================
# Script: Apply-WIBRSPRegistryChange.ps1
# Created: 2025-04-24 18:10:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-25 23:30:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.5.2
# Additional Info: Complete rewrite for PowerShell 5.1 compatibility
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

# Initial console message to confirm script is running
Write-Host "Starting registry change script..." -ForegroundColor Cyan

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
    } 
    catch {
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
    }
    catch {
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
    }
    else {
        Write-Log "User '$Username' is not currently logged in." "DETAIL"
        return $false
    }
}

# Function to test registry changes - completely rewritten for PS 5.1 compatibility
function Test-RegistryChanges {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
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

    Write-Log "Verifying registry changes..." "PROCESS"
    
    # Initialize variables
    $verificationPath = $null
    $userSID = $null
    $currentValue = $null
    
    # STEP 1: Determine verification path - simple single purpose block
    if ($Username) {
        if ($LoadedHivePathForVerification) {
            $verificationPath = "Registry::$LoadedHivePathForVerification\$RegKeyRelativePath"
            Write-Log "Verification Target: Loaded Hive at $verificationPath" "DETAIL"
        }
        else {
            try {
                $userSID = Get-UserSID -Username $Username
                if (-not $userSID) { 
                    Write-Log "Could not get SID for verification for user '$Username'." "ERROR"
                    return $false
                }
                $verificationPath = "Registry::HKEY_USERS\$userSID\$RegKeyRelativePath"
                Write-Log "Verification Target: Live User Hive at $verificationPath" "DETAIL"
            }
            catch {
                Write-Log "Error getting user SID: $_" "ERROR"
                return $false
            }
        }
    } 
    else {
        $verificationPath = "Registry::HKEY_CURRENT_USER\$RegKeyRelativePath"
        Write-Log "Verification Target: HKEY_CURRENT_USER at $verificationPath" "DETAIL"
    }
    
    # STEP 2: Check if path exists
    if (-not (Test-Path -Path $verificationPath)) {
        Write-Log "✗ Verification FAILED: Registry key path does not exist: $verificationPath" "ERROR"
        return $false
    }
    
    # STEP 3: Read registry value
    try {
        $currentValue = Get-ItemProperty -Path $verificationPath -Name $RegValueName -ErrorAction Stop | 
                        Select-Object -ExpandProperty $RegValueName
        Write-Log "✓ Verification: Successfully read value '$RegValueName'." "DETAIL"
    }
    catch {
        Write-Log "✗ Verification FAILED: Could not read value '$RegValueName'. Error: $_" "ERROR"
        return $false
    }
    
    # STEP 4: Check for null values
    if ($null -eq $currentValue -and $null -ne $RegValueData) {
        Write-Log "✗ Verification FAILED: Value exists but is null/empty, expected non-null value." "WARNING"
        return $false
    }
    
    # STEP 5: Compare values based on type
    # Special case for TimerAutoMount
    if ($RegValueName -eq "TimerAutoMount" -and $RegValueType -eq [Microsoft.Win32.RegistryValueKind]::Binary) {
        Write-Log "✓ Verification SUCCESSFUL: TimerAutoMount registry value exists (existence check only)." "SUCCESS"
        return $true
    }
    
    # Binary value comparison
    if (($currentValue -is [byte[]]) -and ($RegValueData -is [byte[]])) {
        $comparisonResult = $true
        # Simple length check first
        if ($currentValue.Length -ne $RegValueData.Length) {
            $comparisonResult = $false
        }
        else {
            # Check each byte
            for ($i = 0; $i -lt $currentValue.Length; $i++) {
                if ($currentValue[$i] -ne $RegValueData[$i]) {
                    $comparisonResult = $false
                    break
                }
            }
        }
        
        if ($comparisonResult) {
            Write-Log "✓ Verification SUCCESSFUL: Binary value matches expected." "SUCCESS"
            return $true
        }
        else {
            Write-Log "✗ Verification FAILED: Binary value does not match expected." "WARNING"
            return $false
        }
    }

    # MultiString comparison
    if ($RegValueType -eq [Microsoft.Win32.RegistryValueKind]::MultiString) {
        $comparisonResult = $true
        # Simple length check first
        if (($currentValue -is [array]) -and ($RegValueData -is [array])) {
            if ($currentValue.Length -ne $RegValueData.Length) {
                $comparisonResult = $false
            }
            else {
                # Check each string
                for ($i = 0; $i -lt $currentValue.Length; $i++) {
                    if ($currentValue[$i] -ne $RegValueData[$i]) {
                        $comparisonResult = $false
                        break
                    }
                }
            }
        }
        else {
            # One of them is not an array
            $comparisonResult = ($currentValue -eq $RegValueData)
        }
        
        if ($comparisonResult) {
            Write-Log "✓ Verification SUCCESSFUL: MultiString value matches expected." "SUCCESS"
            return $true
        }
        else {
            Write-Log "✗ Verification FAILED: MultiString value does not match expected." "WARNING"
            return $false
        }
    }
    
    # For all other types - direct comparison
    if ($RegValueData -eq $currentValue) {
        Write-Log "✓ Verification SUCCESSFUL: Value exists and matches expected." "SUCCESS"
        return $true
    }
    else {
        Write-Log "✗ Verification FAILED: Value does not match expected." "WARNING"
        Write-Log "  Expected: $RegValueData ($($RegValueData.GetType().Name))" "DETAIL"
        Write-Log "  Actual:   $currentValue ($($currentValue.GetType().Name))" "DETAIL"
        return $false
    }
}

# Function alias with approved verb - redirects to Test-RegistryChanges
function Confirm-RegistryChanges {
    [CmdletBinding(SupportsShouldProcess = $true)]
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
}

# --- Main Script Logic - Simplified structure for PS 5.1 ---
# Step 1: Initialize
$ErrorActionPreference = "Stop"
$verificationSuccess = $false
$applySuccess = $false
$userHiveLoadedForApply = $false
$userHiveLoadedForVerify = $false
$userHivePath = $null
$targetRegKeyPath = $null
$targetRegRootPath = $null

try {
    # Log script start
    Write-Log "Starting registry change application script" "INFO"
    Write-Log "Script version: 1.5.2" "DETAIL"
    Write-Log "Log file: $LogPath" "DETAIL"
    
    # Check for Admin
    $currentUserIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($currentUserIdentity)
    $isAdmin = $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    Write-Log "Running as: $($currentUserIdentity.Name) (Elevated: $isAdmin)" "DETAIL"
    
    if (-not $isAdmin) {
        Write-Host "ERROR: This script requires administrative privileges." -ForegroundColor Red
        throw "This script requires administrative privileges."
    }
    
    # Setup variables
    $regExeApplyHiveMount = "HKLM\TempHiveApply"
    $regExeVerifyHiveMount = "HKLM\TempHiveVerify"
    $psApplyHiveRootPath = "Registry::$regExeApplyHiveMount"
    $psVerifyHiveRootPath = "Registry::$regExeVerifyHiveMount"
    
    # Target determination
    if ($Username) {
        Write-Log "Target user specified: $Username" "PROCESS"
        $userSID = Get-UserSID -Username $Username
        
        if (-not $userSID) { 
            throw "Could not get SID for user '$Username'."
        }
        
        $isUserLoggedIn = Test-UserLoggedIn -Username $Username
        
        if ($isUserLoggedIn) {
            $targetRegRootPath = "Registry::HKEY_USERS\$userSID"
            $targetRegKeyPath = Join-Path -Path $targetRegRootPath -ChildPath $regKeyRelativePath
            Write-Log "Targeting live user hive: $targetRegKeyPath" "PROCESS"
        }
        else {
            # Load user hive
            Write-Log "User not logged in. Locating user hive..." "PROCESS"
            $userProfile = Get-CimInstance -ClassName Win32_UserProfile | Where-Object { $_.SID -eq $userSID }
            
            if (-not $userProfile) { 
                throw "User profile for SID '$userSID' ($Username) not found." 
            }
            
            $userHivePath = Join-Path -Path $userProfile.LocalPath -ChildPath "NTUSER.DAT"
            
            if (-not (Test-Path -Path $userHivePath)) { 
                throw "NTUSER.DAT not found at $userHivePath." 
            }
            
            Write-Log "Loading user hive: $userHivePath" "PROCESS"
            
            if ($PSCmdlet.ShouldProcess("Registry Hive: $userHivePath", "Load into $regExeApplyHiveMount")) {
                # Clear the mount point if it exists
                if (Test-Path -Path $psApplyHiveRootPath) {
                    Write-Log "Unloading existing hive mount..." "DEBUG"
                    reg.exe unload $regExeApplyHiveMount 2>&1 | Out-Null
                    Start-Sleep -Seconds 1
                }
                
                # Load the hive
                $loadResult = reg.exe load $regExeApplyHiveMount $userHivePath 2>&1
                
                if ($LASTEXITCODE -ne 0) {
                    throw "Failed to load user registry hive: $loadResult"
                }
                
                $userHiveLoadedForApply = $true
                Write-Log "User registry hive loaded successfully." "SUCCESS"
                $targetRegKeyPath = Join-Path -Path $psApplyHiveRootPath -ChildPath $regKeyRelativePath
                Write-Log "Targeting loaded hive: $targetRegKeyPath" "PROCESS"
            }
            else {
                Write-Log "WhatIf: Would load registry hive" "INFO"
            }
        }
    }
    else {
        # No username specified
        Write-Log "Targeting current user" "PROCESS"
        $targetRegRootPath = "Registry::HKEY_CURRENT_USER"
        $targetRegKeyPath = Join-Path -Path $targetRegRootPath -ChildPath $regKeyRelativePath
    }
}
catch {
    Write-Host "Initialization Error: $_" -ForegroundColor Red
    Write-Log "Error in initialization: $_" "ERROR"
    # Cleanup will happen in Finally block
}

# Apply registry change
try {
    if ($targetRegKeyPath) {
        Write-Log "Applying registry change: $regValueName = $regValueData" "PROCESS"
        
        if ($PSCmdlet.ShouldProcess($targetRegKeyPath, "Set registry value")) {
            # Ensure parent key exists
            $parentPath = Split-Path -Path $targetRegKeyPath -Parent
            
            if (-not (Test-Path -Path $parentPath)) {
                Write-Log "Creating parent key..." "DETAIL"
                $null = New-Item -Path $parentPath -Force
            }
            
            # Set value
            Set-ItemProperty -Path $targetRegKeyPath -Name $regValueName -Value $regValueData -Type $regValueType -Force
            Write-Log "Registry value set successfully." "SUCCESS"
            $applySuccess = $true
        }
        else {
            Write-Log "WhatIf: Would set registry value" "INFO"
            $applySuccess = $true # For WhatIf mode
        }
    }
}
catch {
    Write-Host "Error applying registry change: $_" -ForegroundColor Red
    Write-Log "Error applying registry change: $_" "ERROR"
    $applySuccess = $false
}

# Verify registry change
try {
    if ($applySuccess) {
        Write-Log "Verifying registry change..." "PROCESS"
        
        # Set up verification parameters
        $verifyParams = @{
            RegKeyRelativePath = $regKeyRelativePath
            RegValueName = $regValueName
            RegValueData = $regValueData
            RegValueType = $regValueType
        }
        
        # Add username if specified
        if ($Username) {
            $verifyParams.Username = $Username
            
            # For logged-off users, we need to load the hive for verification
            if (-not $isUserLoggedIn -and $userHivePath) {
                if ($PSCmdlet.ShouldProcess("Registry Hive: $userHivePath", "Load for verification")) {
                    # Ensure verify mount point is clear
                    if (Test-Path -Path $psVerifyHiveRootPath) {
                        reg.exe unload $regExeVerifyHiveMount 2>&1 | Out-Null
                        Start-Sleep -Seconds 1
                    }
                    
                    # Load for verification
                    $loadVerifyResult = reg.exe load $regExeVerifyHiveMount $userHivePath 2>&1
                    
                    if ($LASTEXITCODE -eq 0) {
                        $userHiveLoadedForVerify = $true
                        $verifyParams.LoadedHivePathForVerification = $regExeVerifyHiveMount
                        Write-Log "Loaded hive for verification" "SUCCESS"
                    }
                    else {
                        Write-Log "Failed to load hive for verification: $loadVerifyResult" "WARNING"
                    }
                }
            }
        }
        
        # Perform verification if possible
        if ($isUserLoggedIn -or -not $Username -or $userHiveLoadedForVerify) {
            $verificationSuccess = Confirm-RegistryChanges @verifyParams -WhatIf:$false
        }
        else {
            Write-Log "Cannot verify - user not logged in and hive not loaded" "WARNING"
        }
    }
    else {
        Write-Log "Skipping verification due to failed apply" "WARNING"
    }
}
catch {
    Write-Host "Error in verification: $_" -ForegroundColor Red
    Write-Log "Error in verification: $_" "ERROR"
    $verificationSuccess = $false
}
finally {
    # Cleanup
    Write-Log "Performing cleanup..." "PROCESS"
    
    # Unload hives if loaded
    if ($userHiveLoadedForApply) {
        if ($PSCmdlet.ShouldProcess($regExeApplyHiveMount, "Unload Registry Hive")) {
            reg.exe unload $regExeApplyHiveMount 2>&1 | Out-Null
            Write-Log "Unloaded Apply hive" "SUCCESS"
        }
    }
    
    if ($userHiveLoadedForVerify) {
        if ($PSCmdlet.ShouldProcess($regExeVerifyHiveMount, "Unload Registry Hive")) {
            reg.exe unload $regExeVerifyHiveMount 2>&1 | Out-Null
            Write-Log "Unloaded Verify hive" "SUCCESS"
        }
    }
    
    # Final status
    Write-Log "Script finished. Success: $verificationSuccess" "INFO"
    Write-Host "Script execution complete. Success: $verificationSuccess" -ForegroundColor $(if ($verificationSuccess) { "Green" } else { "Yellow" })
}
