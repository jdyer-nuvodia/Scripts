# =============================================================================
# Script: Get-SecurityConfigurationStatus.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.9.2
# Additional Info: Fixed header metadata for workflow validation
# =============================================================================

<#
.SYNOPSIS
    Comprehensive analysis of system security settings and Group Policy configurations.
.DESCRIPTION
    This script performs a detailed analysis of security configurations including:
    - Group Policy settings (computer and user)
    - Security policy settings
    - Audit policies
    - System access controls
    - Security templates
    - Security database settings
    - Advanced registry security settings

    Dependencies:
    - Administrative privileges
    - GroupPolicy PowerShell module (optional)
    - secedit.exe
    - auditpol.exe
    - Access to system registry

    The script generates both console output and a detailed log file
    with color-coded status indicators for different types of information.
.PARAMETER OutputFormat
    The format for the GPResult report output
    Valid values: 'HTML', 'Text'
    Default: 'HTML'
.EXAMPLE
    .\Get-SecurityConfigurationStatus.ps1
    Runs the analysis with default HTML output format
.EXAMPLE
    .\Get-SecurityConfigurationStatus.ps1 -OutputFormat Text
    Runs the analysis and outputs results in text format
.NOTES
    Security Level: High
    Required Permissions: Local Administrator
    Validation Requirements:
    - Must run with administrative privileges
    - Requires access to system security settings
    - Domain connection for Group Policy analysis
#>

#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('HTML', 'Text')]
    [string]$OutputFormat = 'HTML'
)

<#
.SYNOPSIS
    Writes a formatted status message using appropriate PowerShell streams.
.DESCRIPTION
    Outputs messages using the appropriate PowerShell output streams based on message type:
    - Info: Write-Output (standard information)
    - Process: Write-Verbose (processing updates)
    - Success: Write-Output (successful operations)
    - Warning: Write-Warning (warning messages)
    - Error: Write-Error (error messages)
    - Debug: Write-Debug (debug information)
    - Detail: Write-Verbose (detailed/verbose information)
.PARAMETER Message
    The message text to display
.PARAMETER Type
    The type of message determining the output stream
    Valid values: Info, Process, Success, Warning, Error, Debug, Detail
.EXAMPLE
    Write-StatusMessage "Operation completed" "Success"
    Outputs "Operation completed" using Write-Output
#>
function Write-StatusMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet("Info", "Process", "Success", "Warning", "Error", "Debug", "Detail")]
        [string]$Type = "Info"
    )

    switch ($Type) {
        "Info" { Write-Output $Message }
        "Process" { Write-Verbose $Message }
        "Success" { Write-Output $Message }
        "Warning" { Write-Warning $Message }
        "Error" { Write-Error $Message }
        "Debug" { Write-Debug $Message }
        "Detail" { Write-Verbose $Message }
    }
}

<#
.SYNOPSIS
    Checks if the current system is a Domain Controller.
.DESCRIPTION
    Determines if the current system is a Domain Controller by checking
    the system's domain role value. Domain Controllers have a role
    value of 4 or 5.
.EXAMPLE
    if (Test-IsDomainController) { Write-Output "Running on DC" }
.OUTPUTS
    System.Boolean
    Returns True if running on a Domain Controller, False otherwise
#>
function Test-IsDomainController {
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    return (Get-CimInstance Win32_ComputerSystem).DomainRole -ge 4
}

<#
.SYNOPSIS
    Generates a Group Policy status report.
.DESCRIPTION
    Generates a detailed Group Policy report using gpresult.
    Can output in either HTML or text format.
.PARAMETER OutputFormat
    Format of the report (HTML or Text)
.EXAMPLE
    Get-GPStatusWithGpresult -OutputFormat HTML
    Generates an HTML report for group policies
#>
function Get-GPStatusWithGpresult {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [ValidateSet("HTML", "Text")]
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
        } catch {
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
    } catch {
        Write-StatusMessage "Failed to generate Group Policy report: $($_.Exception.Message)" -Type "Error"
    }
}

# Function to get detailed security and password policy settings
function Get-SecurityPolicySetting {
    Write-StatusMessage "Analyzing Security Policy Settings..." -Type "Process"

    try {
        # Create a unique temporary file for secedit export
        $secpolPath = Join-Path $env:TEMP "secpol_$(Get-Random).cfg"

        # Export security policy settings
        $null = secedit /export /cfg $secpolPath

        if (Test-Path $secpolPath) {
            $securitySettings = Get-Content $secpolPath | Where-Object {
                $_ -match "Password|MinimumPasswordAge|MaximumPasswordAge|PasswordComplexity|LockoutBadCount|ResetLockoutCount"
            }
            Remove-Item $secpolPath -Force

            Write-StatusMessage "`nPassword Policy Settings:" -Type "Info"
            foreach ($setting in $securitySettings) {
                $name = ($setting -split '=')[0].Trim()
                $value = ($setting -split '=')[1].Trim()
                Write-StatusMessage "  $name : $value" -Type "Detail"
            }

            # Get additional security settings using PowerShell
            Write-StatusMessage "`nAdditional Security Settings:" -Type "Info"

            # Account Policies
            $accountPolicies = net accounts
            Write-StatusMessage "Account Policies:" -Type "Success"
            $accountPolicies | Where-Object { $_ -match ":" } | ForEach-Object {
                Write-StatusMessage "  $_" -Type "Detail"
            }

            # User Rights Assignment (if on domain)
            if ($env:USERDOMAIN -ne $env:COMPUTERNAME) {
                Write-StatusMessage "`nUser Rights Assignment:" -Type "Success"
                $userRightsPath = Join-Path $env:TEMP "userrights_$(Get-Random).cfg"
                $null = secedit /export /areas USER_RIGHTS /cfg $userRightsPath

                if (Test-Path $userRightsPath) {
                    $rightsSettings = Get-Content $userRightsPath | Where-Object {
                        $_ -match "SeSecurityPrivilege|SeBackupPrivilege|SeRestorePrivilege"
                    }
                    Remove-Item $userRightsPath -Force
                    foreach ($right in $rightsSettings) {
                        $name = ($right -split '=')[0].Trim()
                        $value = ($right -split '=')[1].Trim()
                        Write-StatusMessage "  $name : $value" -Type "Detail"
                    }
                }
            }
        }
    } catch {
        Write-StatusMessage "Error retrieving security settings: $($_.Exception.Message)" -Type "Error"
    }
}

# Function to get audit policy settings
function Get-AuditPolicySetting {
    Write-StatusMessage "Analyzing Audit Policy Settings..." -Type "Process"

    try {
        $auditPol = auditpol /get /category:* /r | ConvertFrom-Csv

        Write-StatusMessage "`nAudit Policy Settings:" -Type "Info"
        foreach ($policy in $auditPol) {
            Write-StatusMessage "Category: $($policy.'Subcategory')" -Type "Success"
            Write-StatusMessage "  Setting: $($policy.'Inclusion Setting')" -Type "Detail"
        }
    } catch {
        Write-StatusMessage "Error retrieving audit policy settings: $($_.Exception.Message)" -Type "Error"
    }
}

# Function to get system access control settings
function Get-SystemAccessControl {
    Write-StatusMessage "Analyzing System Access Control Settings..." -Type "Process"

    try {
        $sysctrlPath = Join-Path $env:TEMP "sysctrl_$(Get-Random).cfg"
        $null = secedit /export /cfg $sysctrlPath

        if (Test-Path $sysctrlPath) {
            $systemSettings = Get-Content $sysctrlPath | Where-Object {
                $_ -match "EnableAdminAccount|EnableGuestAccount|LSAAnonymousNameLookup|RestrictAnonymousSAM"
            }
            Remove-Item $sysctrlPath -Force

            Write-StatusMessage "`nSystem Access Control Settings:" -Type "Info"
            foreach ($setting in $systemSettings) {
                $name = ($setting -split '=')[0].Trim()
                $value = ($setting -split '=')[1].Trim()
                Write-StatusMessage "  $name : $value" -Type "Detail"
            }

            # Get Registry Security Settings
            Write-StatusMessage "`nRegistry Security Settings:" -Type "Info"
            $registryPaths = @(
                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System",
                "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa"
            )

            foreach ($StartPath in $registryPaths) {
                if (Test-Path $StartPath) {
                    Write-StatusMessage "Registry Path: $StartPath" -Type "Success"
                    Get-ItemProperty -Path $StartPath |
                        Select-Object -Property * -ExcludeProperty PS* |
                        ForEach-Object {
                            $_.PSObject.Properties | ForEach-Object {
                                Write-StatusMessage "  $($_.Name): $($_.Value)" -Type "Detail"
                            }
                        }
                }
            }
        }
    } catch {
        Write-StatusMessage "Error retrieving system access control settings: $($_.Exception.Message)" -Type "Error"
    }
}

# Function to get security template settings
function Get-SecurityTemplateSetting {
    Write-StatusMessage "Analyzing Security Template Settings..." -Type "Process"

    try {
        # Create temporary security template
        $templatePath = Join-Path $env:TEMP "security_template.inf"
        secedit /export /cfg $templatePath | Out-Null

        if (Test-Path $templatePath) {
            $templateContent = Get-Content $templatePath

            Write-StatusMessage "`nSecurity Template Settings:" -Type "Info"

            # Analyze different sections
            $currentSection = ""
            foreach ($line in $templateContent) {
                if ($line -match '^\[(.+)\]') {
                    $currentSection = $matches[1]
                    Write-StatusMessage "`nSection: $currentSection" -Type "Success"
                } elseif ($line -match '^(.+?)\s*=\s*(.+)$') {
                    $setting = $matches[1].Trim()
                    $value = $matches[2].Trim()
                    Write-StatusMessage "  $setting = $value" -Type "Detail"
                }
            }

            Remove-Item $templatePath -Force
        }
    } catch {
        Write-StatusMessage "Error analyzing security template: $($_.Exception.Message)" -Type "Error"
    }
}

# Function to get security database settings
function Get-SecurityDatabaseSetting {
    Write-StatusMessage "Analyzing Security Database Settings..." -Type "Process"

    try {
        # Check Security Configuration and Analysis settings
        $scaPath = Join-Path $env:TEMP "sca_analysis"

        # Create new security database
        secedit /export /cfg "$scaPath.inf" | Out-Null
        secedit /configure /db "$scaPath.sdb" /cfg "$scaPath.inf" /quiet

        if (Test-Path "$scaPath.inf") {
            Write-StatusMessage "`nSecurity Database Analysis:" -Type "Info"

            # Get security areas
            $areas = @("Account Policies", "Local Policies", "Event Log", "Restricted Groups",
                "System Services", "Registry", "File System")

            foreach ($area in $areas) {
                Write-StatusMessage "`nAnalyzing $area..." -Type "Success"
                secedit /areas $area /export /cfg "$scaPath`_$($area -replace '\s', '_').inf" /quiet

                if (Test-Path "$scaPath`_$($area -replace '\s', '_').inf") {
                    $content = Get-Content "$scaPath`_$($area -replace '\s', '_').inf"
                    $content | Where-Object { $_ -match '=' } | ForEach-Object {
                        Write-StatusMessage "  $_" -Type "Detail"
                    }
                    Remove-Item "$scaPath`_$($area -replace '\s', '_').inf" -Force
                }
            }

            # Cleanup
            Remove-Item "$scaPath.inf" -Force -ErrorAction SilentlyContinue
            Remove-Item "$scaPath.sdb" -Force -ErrorAction SilentlyContinue
            Remove-Item "$scaPath.jfm" -Force -ErrorAction SilentlyContinue
        }
    } catch {
        Write-StatusMessage "Error analyzing security database: $($_.Exception.Message)" -Type "Error"
    }
}

# Function to get advanced registry settings
function Get-AdvancedRegistrySetting {
    Write-StatusMessage "Analyzing Advanced Registry Security Settings..." -Type "Process"

    try {
        $registryPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System",
            "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa",
            "HKLM:\SYSTEM\CurrentControlSet\Services\LanManServer\Parameters",
            "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest",
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\Kerberos\Parameters",
            "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\kernel",
            "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Installer"
        )

        Write-StatusMessage "`nAdvanced Registry Security Settings:" -Type "Info"

        foreach ($StartPath in $registryPaths) {
            if (Test-Path $StartPath) {
                Write-StatusMessage "`nRegistry Path: $StartPath" -Type "Success"
                try {
                    $properties = Get-ItemProperty -Path $StartPath -ErrorAction Stop
                    $properties.PSObject.Properties |
                        Where-Object { $_.Name -notlike 'PS*' } |
                        ForEach-Object {
                            $value = if ($_.Value -is [byte[]]) {
                                [System.BitConverter]::ToString($_.Value)
                            } else {
                                $_.Value
                            }
                            Write-StatusMessage "  $($_.Name): $value" -Type "Detail"
                        }
                } catch {
                    Write-StatusMessage "  Error reading properties: $($_.Exception.Message)" -Type "Warning"
                }
            }
        }
    } catch {
        Write-StatusMessage "Error analyzing registry settings: $($_.Exception.Message)" -Type "Error"
    }
}

# Get system and domain information
$computerSystem = Get-CimInstance Win32_ComputerSystem
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

    # Get all security-related settings
    Get-SecurityPolicySetting
    Get-AuditPolicySetting
    Get-SystemAccessControl
    Get-SecurityTemplateSetting
    Get-SecurityDatabaseSetting
    Get-AdvancedRegistrySetting

    if ($domainName -eq "WORKGROUP") {
        Write-StatusMessage "Computer is in a workgroup. Limited Group Policy information will be available." -Type "Warning"
        Get-GPStatusWithGpresult -OutputFormat $OutputFormat
    } elseif (Test-IsDomainController) {
        if (Get-Module -ListAvailable -Name GroupPolicy) {
            Import-Module GroupPolicy

            # Get Computer Policy Settings
            Write-StatusMessage "Analyzing Computer Policy Settings..." -Type "Process"
            try {
                $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
                $userPolicies = Get-GPResultantSetOfPolicy -ReportType User -User $currentUser

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
            } catch {
                Write-StatusMessage "Unable to retrieve policies for current user" -Type "Warning"
                Write-StatusMessage $_.Exception.Message -Type "Error"
            }
        } else {
            Write-StatusMessage "GroupPolicy module not available. Falling back to gpresult." -Type "Warning"
            Get-GPStatusWithGpresult -OutputFormat $OutputFormat
        }
    } else {
        Write-StatusMessage "Computer is domain-joined but not a Domain Controller. Using gpresult for analysis." -Type "Info"
        Get-GPStatusWithGpresult -OutputFormat $OutputFormat
    }
} catch {
    Write-StatusMessage "An error occurred while analyzing Group Policy settings" "Error"
    Write-StatusMessage $_.Exception.Message "Error"
} finally {
    Write-StatusMessage "`nAnalysis Summary:" "Info"
    Write-StatusMessage "System Name: $computerName" "Detail"
    Write-StatusMessage "Domain: $domainName" "Detail"
    Write-StatusMessage "Log file saved to: $LogPath" "Success"
    Stop-Transcript
}
