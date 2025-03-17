# =============================================================================
# Script: Get-GroupPolicyStatus.ps1
# Created: 2024-03-17 17:35:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2024-03-17 18:26:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.3.0
# Additional Info: Added better handling for workgroup computers and non-domain scenarios
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

[CmdletBinding()]
param(
    [ValidateSet('HTML', 'Text')]
    [string]$OutputFormat = 'HTML'
)

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
        [string]$ReportType = "Both",
        [string]$OutputFormat = $script:OutputFormat
    )
    
    $tempFolder = Join-Path $env:TEMP "GPReport"
    if (-not (Test-Path $tempFolder)) {
        New-Item -ItemType Directory -Path $tempFolder | Out-Null
    }

    if ($OutputFormat -eq 'HTML') {
        $reportFile = Join-Path $tempFolder "GPReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
        Write-StatusMessage "Generating HTML GPResult report..." -Type "Process"
        
        try {
            $process = Start-Process -FilePath "gpresult.exe" -ArgumentList "/H `"$reportFile`"", "/F" -Wait -NoNewWindow -PassThru
            
            if ($process.ExitCode -eq 0 -and (Test-Path $reportFile)) {
                Write-StatusMessage "Report generated successfully at: $reportFile" -Type "Success"
                Start-Process $reportFile
                return
            }
        }
        catch {
            Write-StatusMessage "HTML report generation failed: $($_.Exception.Message)" -Type "Warning"
        }
    }

    # Fallback or primary text report generation
    Write-StatusMessage "Generating text-based report..." -Type "Process"
    try {
        $textReport = gpresult.exe /R
        if ($textReport) {
            Write-StatusMessage "`nGroup Policy Report (Text Format):" -Type "Info"
            $textReport | Where-Object { $_ -notmatch "ERROR:|INFO:" } | ForEach-Object {
                if ($_.Trim()) {
                    Write-StatusMessage $_ -Type "Detail"
                }
            }
        } else {
            Write-StatusMessage "No Group Policy settings found" -Type "Warning"
        }
    }
    catch {
        Write-StatusMessage "Failed to generate Group Policy report: $($_.Exception.Message)" -Type "Error"
    }
}

# Get system and domain information
$computerSystem = Get-WmiObject Win32_ComputerSystem
$computerName = $computerSystem.Name
$domainName = if ($computerSystem.PartOfDomain) { $computerSystem.Domain } else { "WORKGROUP" }

# Initialize log file with system info
$LogPath = Join-Path $PSScriptRoot "$computerName`_$($domainName.Split('.')[0])_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Start-Transcript -Path $LogPath

try {
    Write-StatusMessage "Starting Group Policy analysis..." -Type "Process"
    Write-StatusMessage "System Name: $computerName" -Type "Info"
    Write-StatusMessage "Domain: $domainName" -Type "Info"
    Write-StatusMessage "----------------------------------------" -Type "Info"
    
    if ($domainName -eq "WORKGROUP") {
        Write-StatusMessage "Computer is in a workgroup. Limited Group Policy information will be available." -Type "Warning"
        Get-GPStatusWithGpresult -OutputFormat $OutputFormat
    }
    elseif (Test-IsDomainController) {
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
                    
                    if ($userPolicies.UserResults -and $userPolicies.UserResults.ExtensionData) {
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
                    } else {
                        Write-StatusMessage "  No policy data available for this user" -Type "Warning"
                    }
                }
                catch {
                    Write-StatusMessage "Unable to retrieve policies for user: $($user.Name)" -Type "Warning"
                    Write-StatusMessage $_.Exception.Message -Type "Error"
                }
            }
        } else {
            Write-StatusMessage "GroupPolicy module not available. Falling back to gpresult." -Type "Warning"
            Get-GPStatusWithGpresult -OutputFormat $OutputFormat
        }
    } else {
        Write-StatusMessage "Computer is domain-joined but not a Domain Controller. Using gpresult for analysis." -Type "Info"
        Get-GPStatusWithGpresult -OutputFormat $OutputFormat
    }
}
catch {
    Write-StatusMessage "An error occurred while analyzing Group Policy settings" -Type "Error"
    Write-StatusMessage $_.Exception.Message -Type "Error"
}
finally {
    Write-StatusMessage "`nAnalysis Summary:" -Type "Info"
    Write-StatusMessage "System Name: $computerName" -Type "Detail"
    Write-StatusMessage "Domain: $domainName" -Type "Detail"
    Write-StatusMessage "Log file saved to: $LogPath" -Type "Success"
    Stop-Transcript
}