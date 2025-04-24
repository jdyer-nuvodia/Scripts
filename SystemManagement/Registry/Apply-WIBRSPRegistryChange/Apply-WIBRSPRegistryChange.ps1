# =============================================================================
# Script: Apply-WIBRSPRegistryChange.ps1
# Created: 2025-04-24 18:10:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-24 18:32:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2.1
# Additional Info: Fixed registry verification not finding applied changes
# =============================================================================

<#
.SYNOPSIS
Applies registry changes from a .reg file in the same directory to a specific user's registry hive.

.DESCRIPTION
This script imports registry settings from a file named 'SharePoint_Auto_Mount.reg' located
in the same directory as this script. It verifies the file exists, validates it has a .reg
extension, and applies the registry changes with detailed logging and error handling.
The script includes -WhatIf functionality to preview changes before applying them.

The script can apply changes to either the current user or a specified user profile
by loading that user's NTUSER.DAT hive temporarily.

.PARAMETER Username
The Windows username for which to apply the registry changes. If not specified,
changes will be applied to the current user.

.PARAMETER WhatIf
Shows what would happen if the script runs without actually applying the changes.

.EXAMPLE
.\Apply-WIBRSPRegistryChange.ps1
Applies registry changes from the SharePoint_Auto_Mount.reg file to the current user.

.EXAMPLE
.\Apply-WIBRSPRegistryChange.ps1 -Username "jsmith"
Applies registry changes from the SharePoint_Auto_Mount.reg file to the jsmith user profile.

.EXAMPLE
.\Apply-WIBRSPRegistryChange.ps1 -Username "jsmith" -WhatIf
Shows what registry changes would be applied to the jsmith user profile without actually making changes.
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
$regFilePath = Join-Path -Path $scriptPath -ChildPath "SharePoint_Auto_Mount.reg"

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

# Function to test registry changes
function Test-RegistryChanges {
    param (
        [Parameter(Mandatory = $false)]
        [string]$Username
    )
    
    Write-Log "Verifying registry changes..." "PROCESS"
    
    # Define expected registry path and value
    $regKeyPath = "Software\Microsoft\OneDrive\Accounts\Business1"
    $regValueName = "TimerAutoMount"
    
    if ($Username) {
        # Need to load the user's hive first to verify
        $userHiveLoaded = $false
        $verifyTempHiveKeyName = "HKLM\VerifyTempHive"
        
        try {
            # Find user profile path
            $allProfiles = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*" | 
                           Where-Object { $_.ProfileImagePath -like "*\$Username" }
            
            if (-not $allProfiles) {
                Write-Log "User profile for '$Username' not found on this system" "ERROR"
                return $false
            }
            
            $userProfilePath = $allProfiles.ProfileImagePath
            $userHivePath = Join-Path -Path $userProfilePath -ChildPath "NTUSER.DAT"
            
            if (-not (Test-Path -Path $userHivePath)) {
                Write-Log "User registry hive not found at: $userHivePath" "ERROR"
                return $false
            }
            
            # Load the hive
            Write-Log "Loading user registry hive for verification..." "DETAIL"
            $loadHiveOutput = reg.exe load $verifyTempHiveKeyName $userHivePath 2>&1
            
            if ($LASTEXITCODE -ne 0) {
                Write-Log "Failed to load user registry hive for verification: $loadHiveOutput" "ERROR"
                return $false
            }
            
            $userHiveLoaded = $true
            
            # Create a direct registry access for verification
            New-PSDrive -Name HKV -PSProvider Registry -Root HKEY_LOCAL_MACHINE\VerifyTempHive -ErrorAction SilentlyContinue | Out-Null
            
            # List available registry keys for debugging
            $allKeys = Get-ChildItem -Path "HKV:" -Recurse -ErrorAction SilentlyContinue | 
                       Where-Object { $_ -is [Microsoft.Win32.RegistryKey] } | 
                       Select-Object -ExpandProperty Name
            Write-Log "Available registry keys: $($allKeys.Count) keys found" "DEBUG"
            
            # Check the registry value
            $fullKeyPath = "HKV:\$regKeyPath"
            Write-Log "Checking registry key: $fullKeyPath" "DETAIL"
            
            if (Test-Path -Path $fullKeyPath) {
                if (Get-ItemProperty -Path $fullKeyPath -Name $regValueName -ErrorAction SilentlyContinue) {
                    $value = Get-ItemProperty -Path $fullKeyPath -Name $regValueName
                    $hexValue = "0x" + ($value.$regValueName | ForEach-Object { "{0:X2}" -f $_ }) -join " "
                    
                    Write-Log "Registry value exists in user $Username profile!" "SUCCESS"
                    Write-Log "Value Name: $regValueName" "DETAIL"
                    Write-Log "Value Type: Binary" "DETAIL"
                    Write-Log "Value Data: $hexValue" "DETAIL"
                    
                    if ($value.$regValueName[0] -eq 1) {
                        Write-Log "✓ Verification SUCCESSFUL: TimerAutoMount is set correctly (enabled)" "SUCCESS"
                        Remove-PSDrive -Name HKV -Force -ErrorAction SilentlyContinue
                        return $true
                    } else {
                        Write-Log "✗ Verification FAILED: TimerAutoMount is not set to enabled" "WARNING"
                        Remove-PSDrive -Name HKV -Force -ErrorAction SilentlyContinue
                        return $false
                    }
                } else {
                    Write-Log "✗ Value '$regValueName' does not exist in the registry key" "ERROR"
                    Remove-PSDrive -Name HKV -Force -ErrorAction SilentlyContinue
                    return $false
                }
            } else {
                Write-Log "✗ Registry key path does not exist: $fullKeyPath" "ERROR"
                
                # Debug: try to find similar keys
                $parentPath = Split-Path -Parent $regKeyPath
                $parentKeys = Get-ChildItem -Path "HKV:\$parentPath" -ErrorAction SilentlyContinue | 
                              Select-Object -ExpandProperty Name
                Write-Log "Available keys in parent path: $($parentKeys -join ', ')" "DEBUG"
                
                Remove-PSDrive -Name HKV -Force -ErrorAction SilentlyContinue
                return $false
            }
        }
        finally {
            # Remove PSDrive if it exists
            if (Get-PSDrive -Name HKV -ErrorAction SilentlyContinue) {
                Remove-PSDrive -Name HKV -Force -ErrorAction SilentlyContinue
            }
            
            # Unload the hive if loaded
            if ($userHiveLoaded) {
                Write-Log "Unloading verification registry hive..." "DETAIL"
                [System.GC]::Collect()
                Start-Sleep -Seconds 1
                
                $unloadHiveOutput = reg.exe unload $verifyTempHiveKeyName 2>&1
                if ($LASTEXITCODE -ne 0) {
                    Write-Log "Warning: Failed to unload verification registry hive: $unloadHiveOutput" "WARNING"
                }
            }
        }
    }
    else {
        # Verify current user registry
        $fullKeyPath = "HKCU:\$regKeyPath"
        
        if (Test-Path -Path $fullKeyPath) {
            if (Get-ItemProperty -Path $fullKeyPath -Name $regValueName -ErrorAction SilentlyContinue) {
                $value = Get-ItemProperty -Path $fullKeyPath -Name $regValueName
                $hexValue = "0x" + ($value.$regValueName | ForEach-Object { "{0:X2}" -f $_ }) -join " "
                
                Write-Log "Registry value exists in current user profile!" "SUCCESS"
                Write-Log "Value Name: $regValueName" "DETAIL"
                Write-Log "Value Type: Binary" "DETAIL"
                Write-Log "Value Data: $hexValue" "DETAIL"
                
                if ($value.$regValueName[0] -eq 1) {
                    Write-Log "✓ Verification SUCCESSFUL: TimerAutoMount is set correctly (enabled)" "SUCCESS"
                    return $true
                } else {
                    Write-Log "✗ Verification FAILED: TimerAutoMount is not set to enabled" "WARNING"
                    return $false
                }
            } else {
                Write-Log "✗ Value '$regValueName' does not exist in the registry key" "ERROR"
                return $false
            }
        } else {
            Write-Log "✗ Registry key path does not exist" "ERROR"
            return $false
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

try {    # Log script start
    Write-Log "Starting registry change application script" "INFO"
    Write-Log "Script version: 1.2.1" "DETAIL"
    
    # Check if registry file exists
    if (-not (Test-Path -Path $regFilePath)) {
        Write-Log "Registry file not found: $regFilePath" "ERROR"
        throw "Registry file 'SharePoint_Auto_Mount.reg' not found in script directory."
    }
    
    # Verify file has .reg extension
    if ([System.IO.Path]::GetExtension($regFilePath) -ne ".reg") {
        Write-Log "File does not have .reg extension: $regFilePath" "ERROR"
        throw "The specified file is not a registry (.reg) file."
    }
    
    # Log file details
    $fileInfo = Get-Item -Path $regFilePath
    Write-Log "Registry file found: $($fileInfo.Name)" "PROCESS"
    Write-Log "File size: $([Math]::Round($fileInfo.Length / 1KB, 2)) KB" "DETAIL"
    Write-Log "Last modified: $($fileInfo.LastWriteTime)" "DETAIL"
    
    # Preview first 5 lines of registry file (for information purposes)
    Write-Log "Registry file preview:" "DETAIL"
    Get-Content -Path $regFilePath -TotalCount 5 | ForEach-Object {
        Write-Log "  $_" "DETAIL"
    }
    Write-Log "..." "DETAIL"

    # Check if specified user exists if Username parameter was provided
    $userHiveLoaded = $false
    $tempHiveKeyName = "HKLM\TempHive"
    $modifiedRegFilePath = $null
    
    if ($Username) {
        Write-Log "Target user specified: $Username" "PROCESS"
        
        # Check if the user profile exists
        $userProfilePath = Join-Path -Path (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList").ProfilesDirectory -ChildPath $Username
        if (-not (Test-Path -Path $userProfilePath)) {
            Write-Log "User profile not found at expected location: $userProfilePath" "WARNING"
            
            # Try to find the profile by searching all profiles
            $allProfiles = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*" | 
                           Where-Object { $_.ProfileImagePath -like "*\$Username" }
            
            if ($allProfiles) {
                $userProfilePath = $allProfiles.ProfileImagePath
                Write-Log "Found user profile at: $userProfilePath" "PROCESS"
            }
            else {
                Write-Log "User profile for '$Username' not found on this system" "ERROR"
                throw "User profile for '$Username' not found on this system."
            }
        }
        
        # Path to user's registry hive
        $userHivePath = Join-Path -Path $userProfilePath -ChildPath "NTUSER.DAT"
        
        if (-not (Test-Path -Path $userHivePath)) {
            Write-Log "User registry hive not found at: $userHivePath" "ERROR"
            throw "User registry hive (NTUSER.DAT) not found for user '$Username'."
        }
          # Create temporary registry file with HKEY_CURRENT_USER replaced by HKEY_LOCAL_MACHINE\TempHive
        $regContent = Get-Content -Path $regFilePath -Raw
        $modifiedRegContent = $regContent -replace "HKEY_CURRENT_USER", "HKEY_LOCAL_MACHINE\\TempHive"
        $modifiedRegFilePath = "$regFilePath.temp"
        $modifiedRegContent | Out-File -FilePath $modifiedRegFilePath -Encoding Unicode
        Write-Log "Created temporary modified registry file for the target user" "DETAIL"
          # Create a direct PowerShell registry modification script to avoid reg.exe permission issues
        $psRegScriptPath = "$regFilePath.ps1"
        $psRegScript = @"
# PowerShell Registry Modification Script
`$ErrorActionPreference = 'Stop'
# Create a PSDrive for the loaded hive
New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_LOCAL_MACHINE\TempHive | Out-Null
"@
        
        # Parse the registry file to create PowerShell commands
        $regLines = Get-Content -Path $regFilePath
        $currentKey = $null
        foreach ($line in $regLines) {
            if ($line -match '^\s*\[HKEY_CURRENT_USER\\(.+)\]\s*$') {
                $subKey = $matches[1]
                $currentKey = "HKU:\$subKey"
                $psRegScript += "`n# Ensure the registry key path exists"
                $psRegScript += "`nWrite-Host 'Creating/verifying registry key: $currentKey'"
                $psRegScript += "`nNew-Item -Path '$currentKey' -Force -ErrorAction SilentlyContinue | Out-Null"
            }
            elseif ($line -match '^\s*"([^"]+)"=(.+)$' -and $currentKey) {
                $valueName = $matches[1]
                $valueData = $matches[2]
                
                if ($valueData -match '^"([^"]*)"$') {
                    # String value
                    $valueContent = $matches[1]
                    $psRegScript += "`nWrite-Host 'Setting string value: $valueName'"
                    $psRegScript += "`nSet-ItemProperty -Path '$currentKey' -Name '$valueName' -Value '$valueContent' -Type String"
                }
                elseif ($valueData -match '^dword:([0-9a-fA-F]+)$') {
                    # DWORD value
                    $valueContent = [int]"0x$($matches[1])"
                    $psRegScript += "`nWrite-Host 'Setting DWORD value: $valueName'"
                    $psRegScript += "`nSet-ItemProperty -Path '$currentKey' -Name '$valueName' -Value $valueContent -Type DWord"
                }
                elseif ($valueData -match '^hex\(([0-9a-fA-F]+)\):(.+)$' -or $valueData -match '^hex\(b\):(.+)$') {
                    # Binary or other hex data - this is complex, simplified for common use case
                    $hexValues = $valueData -replace '^hex\(([0-9a-fA-F]+)\):', '' -replace '^hex\(b\):', ''
                    $hexBytes = $hexValues -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
                    $psRegScript += "`nWrite-Host 'Setting binary value: $valueName'"
                    $psRegScript += "`n`$hexBytes = @($($hexBytes -join ','))"
                    $psRegScript += "`nSet-ItemProperty -Path '$currentKey' -Name '$valueName' -Value ([byte[]]`$hexBytes) -Type Binary"
                }
            }
        }
        
        $psRegScript += "`n# Verify the changes"
        $psRegScript += "`nWrite-Host 'Verifying registry changes:'"
        $psRegScript += "`n`$allKeys = Get-ChildItem -Path 'HKU:' -Recurse -ErrorAction SilentlyContinue | Where-Object { `$_ -is [Microsoft.Win32.RegistryKey] }"
        $psRegScript += "`nWrite-Host 'Found keys:' `$allKeys.Name"
        $psRegScript += "`nRemove-PSDrive -Name HKU -Force"
        $psRegScript | Out-File -FilePath $psRegScriptPath -Encoding UTF8
        Write-Log "Created PowerShell registry script to handle registry modifications" "DETAIL"
    }
    
    # Import registry file
    if ($PSCmdlet.ShouldProcess("Registry", "Apply changes from $regFilePath" + $(if($Username) {" to user $Username"} else {""}))) {
        try {
            if ($Username) {
                # Need to load the user's hive first
                Write-Log "Loading user registry hive for $Username..." "PROCESS"
                if ($PSCmdlet.ShouldProcess("Registry", "Load user hive from $userHivePath")) {
                    $loadHiveOutput = reg.exe load $tempHiveKeyName $userHivePath 2>&1
                    if ($LASTEXITCODE -ne 0) {
                        Write-Log "Failed to load user registry hive. Exit code: $LASTEXITCODE" "ERROR"
                        Write-Log "Error details: $loadHiveOutput" "ERROR"
                        throw "Failed to load user registry hive: $loadHiveOutput"
                    }
                    $userHiveLoaded = $true
                    Write-Log "User registry hive loaded successfully" "SUCCESS"
                } 
                else {
                    Write-Log "WhatIf: Would load user registry hive from: $userHivePath" "PROCESS"
                }
                  # Now import the modified registry file to the loaded hive
                Write-Log "Applying registry changes to user $Username..." "PROCESS"
                if ($PSCmdlet.ShouldProcess("Registry", "Apply changes from modified registry file to user hive")) {
                    # Use PowerShell script to apply registry changes instead of reg.exe
                    $scriptOutput = & powershell.exe -ExecutionPolicy Bypass -File $psRegScriptPath 2>&1
                    $scriptExitCode = $LASTEXITCODE
                    
                    if ($scriptExitCode -eq 0) {
                        Write-Log "Registry changes applied successfully to user $Username" "SUCCESS"
                    }
                    else {
                        Write-Log "Failed to apply registry changes. Exit code: $scriptExitCode" "ERROR"
                        Write-Log "Error details: $scriptOutput" "ERROR"
                        throw "Failed to apply registry changes: $scriptOutput"
                    }
                }
                else {
                    Write-Log "WhatIf: Would apply registry changes from modified file to user $Username" "PROCESS"
                }
            }
            else {
                # Apply to current user
                Write-Log "Applying registry changes to current user..." "PROCESS"
                $regImportOutput = reg.exe import $regFilePath 2>&1
                
                if ($LASTEXITCODE -eq 0) {
                    Write-Log "Registry changes applied successfully to current user" "SUCCESS"
                }
                else {
                    Write-Log "Failed to apply registry changes. Exit code: $LASTEXITCODE" "ERROR"
                    Write-Log "Error details: $regImportOutput" "ERROR"
                    throw "Failed to apply registry changes: $regImportOutput"
                }
            }
        }
        finally {
            # Clean up - unload hive if loaded
            if ($userHiveLoaded) {
                Write-Log "Unloading user registry hive..." "PROCESS"
                if ($PSCmdlet.ShouldProcess("Registry", "Unload user hive")) {
                    # Give system time to release locks on the hive
                    [System.GC]::Collect()
                    Start-Sleep -Seconds 1
                    
                    $unloadHiveOutput = reg.exe unload $tempHiveKeyName 2>&1
                    if ($LASTEXITCODE -eq 0) {
                        Write-Log "User registry hive unloaded successfully" "SUCCESS"
                    }
                    else {
                        Write-Log "Warning: Failed to unload user registry hive. Exit code: $LASTEXITCODE" "WARNING"
                        Write-Log "Error details: $unloadHiveOutput" "WARNING"
                    }
                }
                else {
                    Write-Log "WhatIf: Would unload user registry hive" "PROCESS"
                }
            }
              # Remove temporary registry file if created
            if ($modifiedRegFilePath -and (Test-Path -Path $modifiedRegFilePath)) {
                if ($PSCmdlet.ShouldProcess("File", "Remove temporary registry file")) {
                    Remove-Item -Path $modifiedRegFilePath -Force
                    Write-Log "Temporary registry file removed" "DETAIL"
                }
                else {
                    Write-Log "WhatIf: Would remove temporary registry file" "PROCESS"
                }
            }
            
            # Remove temporary PowerShell script if created
            if ($psRegScriptPath -and (Test-Path -Path $psRegScriptPath)) {
                if ($PSCmdlet.ShouldProcess("File", "Remove temporary script file")) {
                    Remove-Item -Path $psRegScriptPath -Force
                    Write-Log "Temporary script file removed" "DETAIL"
                }
                else {
                    Write-Log "WhatIf: Would remove temporary script file" "PROCESS"
                }
            }
        }
    }
    else {
        if ($Username) {
            Write-Log "WhatIf: Would apply registry changes from: $regFilePath to user $Username" "PROCESS"
        }
        else {
            Write-Log "WhatIf: Would apply registry changes from: $regFilePath to current user" "PROCESS"
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
    Write-Log "Script execution completed" "INFO"
}
