# =============================================================================
# Script: Get-GroupPolicyStatus.ps1
# Created: 2024-03-17 17:35:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2024-03-17 18:10:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1.0
# Additional Info: Added support for non-DC environments using gpresult.exe
# =============================================================================

<#
.SYNOPSIS
Retrieves and displays active Group Policy settings for both computer and user configurations.

.DESCRIPTION
This script analyzes the current Group Policy settings applied to the local computer
and all users. It shows which policies are active and their current values. The script
uses native PowerShell commands and the GroupPolicy module to gather this information.

.EXAMPLE
.\Get-GroupPolicyStatus.ps1
Returns a detailed report of all active Group Policy settings

.NOTES
Requires administrative privileges to run
Requires GroupPolicy module
#>

#Requires -RunAsAdministrator

# Function to format output with colors based on status
function Write-StatusMessage {
    param(
        [string]$Message,
        [string]$Type = "Info"
    )
    
    switch ($Type) {
        "Info"    { Write-Host $Message -ForegroundColor White }
        "Process" { Write-Host $Message -ForegroundColor Cyan }
        "Success" { Write-Host $Message -ForegroundColor Green }
        "Warning" { Write-Host $Message -ForegroundColor Yellow }
        "Error"   { Write-Host $Message -ForegroundColor Red }
        "Debug"   { Write-Host $Message -ForegroundColor Magenta }
        "Detail"  { Write-Host $Message -ForegroundColor DarkGray }
    }
}

# Function to check if running on a Domain Controller
function Test-IsDomainController {
    return (Get-WmiObject Win32_ComputerSystem).DomainRole -ge 4
}

# Function to get GP status using gpresult
function Get-GPStatusWithGpresult {
    param (
        [string]$ReportType = "Both"
    )
    
    $tempFile = [System.IO.Path]::GetTempFileName() + ".html"
    
    Write-StatusMessage "Generating GPResult report..." -Type "Process"
    $null = gpresult.exe /H "$tempFile" /F
    
    if (Test-Path $tempFile) {
        Start-Process $tempFile
        Write-StatusMessage "GPResult report generated and opened in default browser" -Type "Success"
    } else {
        Write-StatusMessage "Failed to generate GPResult report" -Type "Error"
    }
}

# Initialize log file
$LogPath = Join-Path $PSScriptRoot "$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Start-Transcript -Path $LogPath

try {
    Write-StatusMessage "Starting Group Policy analysis..." -Type "Process"
    
    if (Test-IsDomainController) {
        if (Get-Module -ListAvailable -Name GroupPolicy) {
            Import-Module GroupPolicy
        
            # Get Computer Policy Settings
            Write-StatusMessage "Analyzing Computer Policy Settings..." -Type "Process"
            $computerPolicies = Get-GPResultantSetOfPolicy -ReportType Computer -Computer $env:COMPUTERNAME
            
            Write-StatusMessage "`nComputer Policies:" -Type "Info"
            $computerPolicies.ComputerResults.ExtensionData | ForEach-Object {
                $extension = $_
                Write-StatusMessage "Category: $($extension.Name)" -Type "Success"
                $extension.Extension.Policy | ForEach-Object {
                    Write-StatusMessage "  Policy: $($_.Name)" -Type "Detail"
                    Write-StatusMessage "  State: $($_.State)" -Type "Detail"
                    Write-StatusMessage "  Setting: $($_.Setting)" -Type "Detail"
                    Write-StatusMessage "" -Type "Detail"
                }
            }
            
            # Get User Policy Settings for all users
            Write-StatusMessage "`nAnalyzing User Policy Settings..." -Type "Process"
            $users = Get-LocalUser | Where-Object Enabled -eq $true
            
            foreach ($user in $users) {
                Write-StatusMessage "`nUser Policies for: $($user.Name)" -Type "Info"
                try {
                    $userPolicies = Get-GPResultantSetOfPolicy -ReportType User -User $user.Name
                    
                    $userPolicies.UserResults.ExtensionData | ForEach-Object {
                        $extension = $_
                        Write-StatusMessage "Category: $($extension.Name)" -Type "Success"
                        $extension.Extension.Policy | ForEach-Object {
                            Write-StatusMessage "  Policy: $($_.Name)" -Type "Detail"
                            Write-StatusMessage "  State: $($_.State)" -Type "Detail"
                            Write-StatusMessage "  Setting: $($_.Setting)" -Type "Detail"
                            Write-StatusMessage "" -Type "Detail"
                        }
                    }
                }
                catch {
                    Write-StatusMessage "Unable to retrieve policies for user: $($user.Name)" -Type "Warning"
                    Write-StatusMessage $_.Exception.Message -Type "Error"
                }
            }
        } else {
            Write-StatusMessage "GroupPolicy module not available. Falling back to gpresult." -Type "Warning"
            Get-GPStatusWithGpresult
        }
    } else {
        Write-StatusMessage "Not running on a Domain Controller. Using gpresult for Group Policy analysis." -Type "Info"
        Get-GPStatusWithGpresult
    }
}
catch {
    Write-StatusMessage "An error occurred while analyzing Group Policy settings" -Type "Error"
    Write-StatusMessage $_.Exception.Message -Type "Error"
}
finally {
    Write-StatusMessage "`nGroup Policy analysis complete. Log file saved to: $LogPath" -Type "Success"
    Stop-Transcript
}