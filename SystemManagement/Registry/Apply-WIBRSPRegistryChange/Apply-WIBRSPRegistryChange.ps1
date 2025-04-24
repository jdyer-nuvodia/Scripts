# =============================================================================
# Script: Apply-WIBRSPRegistryChange.ps1
# Created: 2025-04-24 18:10:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-24 18:55:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1.1
# Additional Info: Fixed registry access issue when applying to another user
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

try {    # Log script start
    Write-Log "Starting registry change application script" "INFO"
    Write-Log "Script version: 1.1.1" "DETAIL"
    
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
New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_LOCAL_MACHINE\TempHive | Out-Null
"@
        
        # Parse the registry file to create PowerShell commands
        $regLines = Get-Content -Path $regFilePath
        $currentKey = $null
        foreach ($line in $regLines) {
            if ($line -match '^\s*\[HKEY_CURRENT_USER\\(.+)\]\s*$') {
                $subKey = $matches[1]
                $currentKey = "HKU:\$subKey"
                $psRegScript += "`nNew-Item -Path '$currentKey' -Force -ErrorAction SilentlyContinue | Out-Null"
            }
            elseif ($line -match '^\s*"([^"]+)"=(.+)$' -and $currentKey) {
                $valueName = $matches[1]
                $valueData = $matches[2]
                
                if ($valueData -match '^"([^"]*)"$') {
                    # String value
                    $valueContent = $matches[1]
                    $psRegScript += "`nSet-ItemProperty -Path '$currentKey' -Name '$valueName' -Value '$valueContent' -Type String"
                }
                elseif ($valueData -match '^dword:([0-9a-fA-F]+)$') {
                    # DWORD value
                    $valueContent = [int]"0x$($matches[1])"
                    $psRegScript += "`nSet-ItemProperty -Path '$currentKey' -Name '$valueName' -Value $valueContent -Type DWord"
                }
                elseif ($valueData -match '^hex\(([0-9a-fA-F]+)\):(.+)$' -or $valueData -match '^hex\(b\):(.+)$') {
                    # Binary or other hex data - this is complex, simplified for common use case
                    $hexValues = $valueData -replace '^hex\(([0-9a-fA-F]+)\):', '' -replace '^hex\(b\):', ''
                    $hexBytes = $hexValues -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
                    $psRegScript += "`n`$hexBytes = @($($hexBytes -join ','))"
                    $psRegScript += "`nSet-ItemProperty -Path '$currentKey' -Name '$valueName' -Value ([byte[]]`$hexBytes) -Type Binary"
                }
            }
        }
        
        $psRegScript += "`nRemove-PSDrive -Name HKU"
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
}
catch {
    Write-Log "An error occurred: $_" "ERROR"
    Write-Log "Exception details: $($_.Exception)" "DEBUG"
    exit 1
}
finally {
    Write-Log "Script execution completed" "INFO"
}
