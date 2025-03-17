# =============================================================================
# Script: Get-GroupPolicyStatus.ps1
# Created: 2024-03-17 17:35:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2024-03-17 18:14:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1.1
# Additional Info: Fixed gpresult access issues and improved error handling
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
    
    $tempFolder = Join-Path $env:TEMP "GPReport"
    if (-not (Test-Path $tempFolder)) {
        New-Item -ItemType Directory -Path $tempFolder | Out-Null
    }
    
    $reportFile = Join-Path $tempFolder "GPReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    
    Write-StatusMessage "Generating GPResult report..." -Type "Process"
    
    try {
        # Run gpresult with explicit scope
        $process = Start-Process -FilePath "gpresult.exe" -ArgumentList "/H `"$reportFile`"", "/F" -Wait -NoNewWindow -PassThru
        
        if ($process.ExitCode -eq 0 -and (Test-Path $reportFile)) {
            Write-StatusMessage "Report generated successfully at: $reportFile" -Type "Success"
            Start-Process $reportFile
        } else {
            # Fall back to text-based report if HTML fails
            Write-StatusMessage "HTML report generation failed, attempting text report..." -Type "Warning"
            $textReport = gpresult.exe /R
            Write-StatusMessage "`nGroup Policy Report (Text Format):" -Type "Info"
            $textReport | ForEach-Object {
                Write-StatusMessage $_ -Type "Detail"
            }
        }
    }
    catch {
        Write-StatusMessage "Error generating Group Policy report: $($_.Exception.Message)" -Type "Error"
        Write-StatusMessage "Attempting to run with scope..." -Type "Warning"
        try {
            # Final fallback - try user scope only
            $textReport = gpresult.exe /SCOPE USER /R
            Write-StatusMessage "`nUser Group Policy Report:" -Type "Info"
            $textReport | ForEach-Object {
                Write-StatusMessage $_ -Type "Detail"
            }
        }
        catch {
            Write-StatusMessage "Failed to generate any Group Policy report: $($_.Exception.Message)" -Type "Error"
        }
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