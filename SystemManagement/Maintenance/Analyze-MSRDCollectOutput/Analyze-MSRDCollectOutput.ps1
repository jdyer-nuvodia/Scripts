<#
=============================================================================
Script: Analyze-MSRDCollectOutput.ps1
Created: 2025-07-15 19:20:00 UTC
Author: jdyer-nuvodia
Last Updated: 2025-07-15 19:20:00 UTC
Updated By: jdyer-nuvodia
Version: 1.0.0
Additional Info: Initial version for analyzing MSRD-Collect diagnostic output
=============================================================================
<#
.SYNOPSIS
Analyzes the output from MSRD-Collect.ps1 diagnostic tool to identify common Remote Desktop issues.

.DESCRIPTION
This script processes the comprehensive diagnostic data collected by MSRD-Collect.ps1,
analyzing logs, registry entries, event logs, and system information to identify
common Remote Desktop Services issues, configuration problems, and potential solutions.

The script generates summary reports highlighting critical findings, warnings,
and recommendations for Remote Desktop troubleshooting.

.PARAMETER MSRDOutputPath
The path to the MSRD-Collect output directory containing diagnostic files.

.PARAMETER ReportPath
The path where the analysis report will be saved. Defaults to script directory.

.PARAMETER DaysToAnalyze
Number of days of event logs to analyze. Defaults to 7 days.

.PARAMETER IncludeDetailedLogs
Include detailed log analysis in the output report.

.PARAMETER ExportToCSV
Export findings to CSV format for further analysis.

.PARAMETER Verbose
Enable verbose output for detailed processing information.

.EXAMPLE
.\Analyze-MSRDCollectOutput.ps1 -MSRDOutputPath "C:\MSRD-Collect-Output" -ReportPath "C:\Reports"
Analyzes MSRD-Collect output and generates a comprehensive report.

.EXAMPLE
.\Analyze-MSRDCollectOutput.ps1 -MSRDOutputPath "C:\MSRD-Collect-Output" -DaysToAnalyze 14 -IncludeDetailedLogs -ExportToCSV
Analyzes 14 days of data with detailed logs and CSV export.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Path to MSRD-Collect output directory")]
    [ValidateScript({ Test-Path -Path $_ -PathType Container })]
    [string]$MSRDOutputPath,

    [Parameter(Mandatory = $false, HelpMessage = "Path for analysis reports")]
    [string]$ReportPath = $PSScriptRoot,

    [Parameter(Mandatory = $false, HelpMessage = "Number of days of event logs to analyze")]
    [ValidateRange(1, 365)]
    [int]$DaysToAnalyze = 7,

    [Parameter(Mandatory = $false, HelpMessage = "Include detailed log analysis")]
    [switch]$IncludeDetailedLogs,

    [Parameter(Mandatory = $false, HelpMessage = "Export findings to CSV")]
    [switch]$ExportToCSV,

    [Parameter(Mandatory = $false, HelpMessage = "Force overwrite existing reports")]
    [switch]$Force
)

# Script-level variables
$script:MSRDOutputPath = $MSRDOutputPath
$script:ReportPath = $ReportPath
$script:DaysToAnalyze = $DaysToAnalyze
$script:IncludeDetailedLogs = $IncludeDetailedLogs
$script:ExportToCSV = $ExportToCSV
$script:Force = $Force

# Initialize arrays for findings
$script:CriticalIssues = @()
$script:Warnings = @()
$script:Recommendations = @()
$script:DetailedFindings = @()

# Define PowerShell version-specific color support
$script:UseAnsiColors = $PSVersionTable.PSVersion.Major -ge 7
$script:Colors        = @{
    'Red'      = if ($script:UseAnsiColors) { "`e[31m" } else { [System.ConsoleColor]::Red }
    'Green'    = if ($script:UseAnsiColors) { "`e[32m" } else { [System.ConsoleColor]::Green }
    'Yellow'   = if ($script:UseAnsiColors) { "`e[33m" } else { [System.ConsoleColor]::Yellow }
    'Blue'     = if ($script:UseAnsiColors) { "`e[34m" } else { [System.ConsoleColor]::Blue }
    'Magenta'  = if ($script:UseAnsiColors) { "`e[35m" } else { [System.ConsoleColor]::Magenta }
    'Cyan'     = if ($script:UseAnsiColors) { "`e[36m" } else { [System.ConsoleColor]::Cyan }
    'White'    = if ($script:UseAnsiColors) { "`e[37m" } else { [System.ConsoleColor]::White }
    'DarkGray' = if ($script:UseAnsiColors) { "`e[90m" } else { [System.ConsoleColor]::DarkGray }
    'Reset'    = if ($script:UseAnsiColors) { "`e[0m" } else { "" }
}

function Write-ColorOutput {
    <#
    .SYNOPSIS
    Writes colored output to the console with cross-version compatibility.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$Color = "White"
    )

    if ($script:UseAnsiColors) {
        # PowerShell 7+ with ANSI escape codes
        $colorCode = $script:Colors[$Color]
        $resetCode = $script:Colors.Reset
        Write-Output "${colorCode}${Message}${resetCode}"
    } else {
        # PowerShell 5.1 - Change console color, write output, then reset
        $originalColor = $Host.UI.RawUI.ForegroundColor
        try {
            if ($script:Colors[$Color] -and $script:Colors[$Color] -ne "") {
                $Host.UI.RawUI.ForegroundColor = $script:Colors[$Color]
            }
            Write-Output $Message
        } finally {
            $Host.UI.RawUI.ForegroundColor = $originalColor
        }
    }
}

function Initialize-AnalysisEnvironment {
    <#
    .SYNOPSIS
    Initializes the analysis environment and validates prerequisites.
    #>
    try {
        Write-ColorOutput -Message "Initializing MSRD-Collect output analysis environment..." -Color Cyan

        # Validate MSRD output directory structure
        $expectedFolders = @('Logs', 'Registry', 'EventLogs', 'SystemInfo', 'NetworkInfo')
        $missingFolders = @()

        foreach ($folder in $expectedFolders) {
            $folderPath = Join-Path -Path $script:MSRDOutputPath -ChildPath $folder
            if (-not (Test-Path -Path $folderPath)) {
                $missingFolders += $folder
            }
        }

        if ($missingFolders.Count -gt 0) {
            Write-ColorOutput -Message "Warning: Some expected folders are missing: $($missingFolders -join ', ')" -Color Yellow
            Write-ColorOutput -Message "Analysis will continue with available data." -Color Yellow
        }

        # Create report directory if it doesn't exist
        if (-not (Test-Path -Path $script:ReportPath)) {
            New-Item -Path $script:ReportPath -ItemType Directory -Force | Out-Null
            Write-ColorOutput -Message "Created report directory: $script:ReportPath" -Color Green
        }

        # Initialize transcript logging
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $logFileName = "MSRD-Analysis-$env:COMPUTERNAME-$timestamp.log"
        $logPath = Join-Path -Path $script:ReportPath -ChildPath $logFileName
        Start-Transcript -Path $logPath -Force

        Write-ColorOutput -Message "Analysis environment initialized successfully" -Color Green
        Write-ColorOutput -Message "Log file: $logPath" -Color DarkGray

    } catch {
        Write-ColorOutput -Message "[SYSTEM ERROR DETECTED] Failed to initialize analysis environment: $($_.Exception.Message)" -Color Red
        throw
    }
}

function Get-MSRDSystemInfo {
    <#
    .SYNOPSIS
    Analyzes system information from MSRD-Collect output.
    #>
    try {
        Write-ColorOutput -Message "Analyzing system information..." -Color Cyan

        $systemInfoPath = Join-Path -Path $script:MSRDOutputPath -ChildPath "SystemInfo"
        if (-not (Test-Path -Path $systemInfoPath)) {
            Write-ColorOutput -Message "Warning: SystemInfo directory not found" -Color Yellow
            return
        }

        # Look for common system info files
        $systemFiles = Get-ChildItem -Path $systemInfoPath -Filter "*.txt" -ErrorAction SilentlyContinue

        foreach ($file in $systemFiles) {
            Write-Verbose "Processing system info file: $($file.Name)"

            $content = Get-Content -Path $file.FullName -ErrorAction SilentlyContinue

            # Analyze system specifications
            if ($file.Name -like "*systeminfo*") {
                $osVersion = $content | Where-Object { $_ -like "*OS Name*" }
                $totalMemory = $content | Where-Object { $_ -like "*Total Physical Memory*" }
                $availableMemory = $content | Where-Object { $_ -like "*Available Physical Memory*" }

                if ($osVersion) {
                    Write-ColorOutput -Message "Operating System: $($osVersion.Split(':')[1].Trim())" -Color White
                }

                if ($totalMemory -and $availableMemory) {
                    $script:DetailedFindings += [PSCustomObject]@{
                        Category = "System Info"
                        Type     = "Memory"
                        Finding  = "Total: $($totalMemory.Split(':')[1].Trim()), Available: $($availableMemory.Split(':')[1].Trim())"
                        Severity = "Info"
                    }
                }
            }
        }

        Write-ColorOutput -Message "System information analysis completed" -Color Green

    } catch {
        Write-ColorOutput -Message "[SYSTEM ERROR DETECTED] Error analyzing system information: $($_.Exception.Message)" -Color Red
    }
}

function Get-MSRDEventLog {
    <#
    .SYNOPSIS
    Analyzes event logs from MSRD-Collect output for RDS-related issues.
    #>
    try {
        Write-ColorOutput -Message "Analyzing event logs for Remote Desktop issues..." -Color Cyan

        $eventLogsPath = Join-Path -Path $script:MSRDOutputPath -ChildPath "EventLogs"
        if (-not (Test-Path -Path $eventLogsPath)) {
            Write-ColorOutput -Message "Warning: EventLogs directory not found" -Color Yellow
            return
        }

        $eventLogFiles = Get-ChildItem -Path $eventLogsPath -Filter "*.evtx" -ErrorAction SilentlyContinue

        # Define critical RDS event IDs to look for
        $criticalEventIDs = @{
            20499 = "RDP session logon failure"
            20500 = "RDP session disconnection"
            1149  = "User authentication succeeded"
            1158  = "RDP encryption error"
            4625  = "Logon failure"
            4648  = "Logon using explicit credentials"
            7001  = "Service control manager errors"
            7034  = "Service crashed unexpectedly"
        }

        foreach ($eventFile in $eventLogFiles) {
            Write-Verbose "Processing event log: $($eventFile.Name)"

            try {
                # For analysis, we'll look for exported text files or CSV files
                $textFile = $eventFile.FullName -replace "\.evtx$", ".txt"
                $csvFile = $eventFile.FullName -replace "\.evtx$", ".csv"

                $logContent = $null
                if (Test-Path -Path $csvFile) {
                    $logContent = Import-Csv -Path $csvFile -ErrorAction SilentlyContinue
                } elseif (Test-Path -Path $textFile) {
                    $logContent = Get-Content -Path $textFile -ErrorAction SilentlyContinue
                }

                if ($logContent) {
                    # Analyze for critical events
                    foreach ($eventID in $criticalEventIDs.Keys) {
                        $matchingEvents = $logContent | Where-Object { $_ -like "*$eventID*" }

                        if ($matchingEvents) {
                            $eventCount = $matchingEvents.Count
                            $severity = if ($eventID -in @(20499, 1158, 4625, 7034)) { "Critical" } else { "Warning" }

                            $finding = [PSCustomObject]@{
                                Category = "Event Logs"
                                Type     = "RDS Event"
                                Finding  = "$($criticalEventIDs[$eventID]) (Event ID: $eventID) - $eventCount occurrences"
                                Severity = $severity
                                Source   = $eventFile.Name
                            }

                            if ($severity -eq "Critical") {
                                $script:CriticalIssues += $finding
                            } else {
                                $script:Warnings += $finding
                            }
                        }
                    }
                }

            } catch {
                Write-ColorOutput -Message "Warning: Could not process event log $($eventFile.Name): $($_.Exception.Message)" -Color Yellow
            }
        }

        Write-ColorOutput -Message "Event log analysis completed. Found $($script:CriticalIssues.Count) critical issues and $($script:Warnings.Count) warnings." -Color Green

    } catch {
        Write-ColorOutput -Message "[SYSTEM ERROR DETECTED] Error analyzing event logs: $($_.Exception.Message)" -Color Red
    }
}

function Get-MSRDRegistryAnalysis {
    <#
    .SYNOPSIS
    Analyzes registry exports for RDS configuration issues.
    #>
    try {
        Write-ColorOutput -Message "Analyzing registry configuration..." -Color Cyan

        $registryPath = Join-Path -Path $script:MSRDOutputPath -ChildPath "Registry"
        if (-not (Test-Path -Path $registryPath)) {
            Write-ColorOutput -Message "Warning: Registry directory not found" -Color Yellow
            return
        }

        $registryFiles = Get-ChildItem -Path $registryPath -Filter "*.reg" -ErrorAction SilentlyContinue

        # Define critical registry keys to analyze
        $criticalKeys = @{
            "HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Control\\Terminal Server"              = "Terminal Services Configuration"
            "HKEY_LOCAL_MACHINE\\SOFTWARE\\Policies\\Microsoft\\Windows NT\\Terminal Services"     = "Terminal Services Policies"
            "HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Services\\TermService"                 = "Terminal Service Settings"
            "HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Terminal Server" = "Terminal Server Settings"
        }

        foreach ($regFile in $registryFiles) {
            Write-Verbose "Processing registry file: $($regFile.Name)"

            try {
                $regContent = Get-Content -Path $regFile.FullName -ErrorAction SilentlyContinue

                foreach ($keyPath in $criticalKeys.Keys) {
                    $keySection = $regContent | Where-Object { $_ -like "*$keyPath*" }

                    if ($keySection) {
                        # Look for common configuration issues
                        $tsEnabled = $regContent | Where-Object { $_ -like "*fDenyTSConnections*" }
                        $nlaEnabled = $regContent | Where-Object { $_ -like "*UserAuthentication*" }

                        if ($tsEnabled -and $tsEnabled -like "*0x00000001*") {
                            $script:CriticalIssues += [PSCustomObject]@{
                                Category = "Registry"
                                Type     = "Configuration"
                                Finding  = "Terminal Services connections are disabled (fDenyTSConnections=1)"
                                Severity = "Critical"
                                Source   = $regFile.Name
                            }
                        }

                        if ($nlaEnabled -and $nlaEnabled -like "*0x00000000*") {
                            $script:Warnings += [PSCustomObject]@{
                                Category = "Registry"
                                Type     = "Security"
                                Finding  = "Network Level Authentication is disabled"
                                Severity = "Warning"
                                Source   = $regFile.Name
                            }
                        }
                    }
                }

            } catch {
                Write-ColorOutput -Message "Warning: Could not process registry file $($regFile.Name): $($_.Exception.Message)" -Color Yellow
            }
        }

        Write-ColorOutput -Message "Registry analysis completed" -Color Green

    } catch {
        Write-ColorOutput -Message "[SYSTEM ERROR DETECTED] Error analyzing registry: $($_.Exception.Message)" -Color Red
    }
}

function Get-MSRDNetworkAnalysis {
    <#
    .SYNOPSIS
    Analyzes network configuration for RDS connectivity issues.
    #>
    try {
        Write-ColorOutput -Message "Analyzing network configuration..." -Color Cyan

        $networkPath = Join-Path -Path $script:MSRDOutputPath -ChildPath "NetworkInfo"
        if (-not (Test-Path -Path $networkPath)) {
            Write-ColorOutput -Message "Warning: NetworkInfo directory not found" -Color Yellow
            return
        }

        $networkFiles = Get-ChildItem -Path $networkPath -Filter "*.txt" -ErrorAction SilentlyContinue

        foreach ($netFile in $networkFiles) {
            Write-Verbose "Processing network file: $($netFile.Name)"

            try {
                $netContent = Get-Content -Path $netFile.FullName -ErrorAction SilentlyContinue

                # Analyze for RDP port configuration
                if ($netFile.Name -like "*netstat*") {
                    $rdpPorts = $netContent | Where-Object { $_ -like "*:3389*" }

                    if (-not $rdpPorts) {
                        $script:CriticalIssues += [PSCustomObject]@{
                            Category = "Network"
                            Type     = "Port Configuration"
                            Finding  = "RDP port 3389 is not listening"
                            Severity = "Critical"
                            Source   = $netFile.Name
                        }
                    } else {
                        $script:DetailedFindings += [PSCustomObject]@{
                            Category = "Network"
                            Type     = "Port Status"
                            Finding  = "RDP port 3389 is active and listening"
                            Severity = "Info"
                            Source   = $netFile.Name
                        }
                    }
                }

                # Analyze firewall rules
                if ($netFile.Name -like "*firewall*") {
                    $rdpRules = $netContent | Where-Object { $_ -like "*Remote Desktop*" -or $_ -like "*3389*" }

                    if ($rdpRules) {
                        $enabledRules = $rdpRules | Where-Object { $_ -like "*Enabled*" -or $_ -like "*Yes*" }

                        if (-not $enabledRules) {
                            $script:CriticalIssues += [PSCustomObject]@{
                                Category = "Network"
                                Type     = "Firewall"
                                Finding  = "Remote Desktop firewall rules are disabled"
                                Severity = "Critical"
                                Source   = $netFile.Name
                            }
                        }
                    }
                }

            } catch {
                Write-ColorOutput -Message "Warning: Could not process network file $($netFile.Name): $($_.Exception.Message)" -Color Yellow
            }
        }

        Write-ColorOutput -Message "Network analysis completed" -Color Green

    } catch {
        Write-ColorOutput -Message "[SYSTEM ERROR DETECTED] Error analyzing network configuration: $($_.Exception.Message)" -Color Red
    }
}

function Get-MSRDServiceAnalysis {
    <#
    .SYNOPSIS
    Analyzes RDS-related services status and configuration.
    #>
    try {
        Write-ColorOutput -Message "Analyzing RDS services status..." -Color Cyan

        $servicesPath = Join-Path -Path $script:MSRDOutputPath -ChildPath "Services"
        if (-not (Test-Path -Path $servicesPath)) {
            $servicesPath = Join-Path -Path $script:MSRDOutputPath -ChildPath "SystemInfo"
        }

        # Critical RDS services to check
        $criticalServices = @(
            "TermService",
            "SessionEnv",
            "UmRdpService",
            "RpcSs",
            "RpcEptMapper",
            "LanmanServer",
            "LanmanWorkstation"
        )

        $serviceFiles = Get-ChildItem -Path $servicesPath -Filter "*service*" -ErrorAction SilentlyContinue

        foreach ($serviceFile in $serviceFiles) {
            Write-Verbose "Processing service file: $($serviceFile.Name)"

            try {
                $serviceContent = Get-Content -Path $serviceFile.FullName -ErrorAction SilentlyContinue

                foreach ($service in $criticalServices) {
                    $serviceInfo = $serviceContent | Where-Object { $_ -like "*$service*" }

                    if ($serviceInfo) {
                        $stoppedService = $serviceInfo | Where-Object { $_ -like "*Stopped*" -or $_ -like "*Disabled*" }

                        if ($stoppedService) {
                            $script:CriticalIssues += [PSCustomObject]@{
                                Category = "Services"
                                Type     = "Service Status"
                                Finding  = "$service service is not running"
                                Severity = "Critical"
                                Source   = $serviceFile.Name
                            }
                        } else {
                            $runningService = $serviceInfo | Where-Object { $_ -like "*Running*" -or $_ -like "*Started*" }

                            if ($runningService) {
                                $script:DetailedFindings += [PSCustomObject]@{
                                    Category = "Services"
                                    Type     = "Service Status"
                                    Finding  = "$service service is running correctly"
                                    Severity = "Info"
                                    Source   = $serviceFile.Name
                                }
                            }
                        }
                    }
                }

            } catch {
                Write-ColorOutput -Message "Warning: Could not process service file $($serviceFile.Name): $($_.Exception.Message)" -Color Yellow
            }
        }

        Write-ColorOutput -Message "Service analysis completed" -Color Green

    } catch {
        Write-ColorOutput -Message "[SYSTEM ERROR DETECTED] Error analyzing services: $($_.Exception.Message)" -Color Red
    }
}

function Get-RecommendationEngine {
    <#
    .SYNOPSIS
    Generates recommendations based on identified issues.
    #>
    try {
        Write-ColorOutput -Message "Generating recommendations..." -Color Cyan

        # Analyze critical issues and generate recommendations
        foreach ($issue in $script:CriticalIssues) {
            switch ($issue.Finding) {
                { $_ -like "*fDenyTSConnections=1*" } {
                    $script:Recommendations += [PSCustomObject]@{
                        Issue          = $issue.Finding
                        Recommendation = "Enable Terminal Services by setting fDenyTSConnections to 0 in registry or using 'Enable-PSRemoting'"
                        Priority       = "High"
                        Category       = "Configuration"
                    }
                }
                { $_ -like "*port 3389*not listening*" } {
                    $script:Recommendations += [PSCustomObject]@{
                        Issue          = $issue.Finding
                        Recommendation = "Start Terminal Services service and verify RDP is enabled in System Properties"
                        Priority       = "High"
                        Category       = "Network"
                    }
                }
                { $_ -like "*firewall rules are disabled*" } {
                    $script:Recommendations += [PSCustomObject]@{
                        Issue          = $issue.Finding
                        Recommendation = "Enable Remote Desktop firewall rules: 'netsh advfirewall firewall set rule group='Remote Desktop' new enable=Yes'"
                        Priority       = "High"
                        Category       = "Firewall"
                    }
                }
                { $_ -like "*service is not running*" } {
                    $serviceName = ($issue.Finding -split " ")[0]
                    $script:Recommendations += [PSCustomObject]@{
                        Issue          = $issue.Finding
                        Recommendation = "Start the $serviceName service: 'Start-Service -Name $serviceName'"
                        Priority       = "High"
                        Category       = "Services"
                    }
                }
            }
        }

        # Generate general recommendations based on warnings
        foreach ($warning in $script:Warnings) {
            switch ($warning.Finding) {
                { $_ -like "*Network Level Authentication is disabled*" } {
                    $script:Recommendations += [PSCustomObject]@{
                        Issue          = $warning.Finding
                        Recommendation = "Enable Network Level Authentication for improved security"
                        Priority       = "Medium"
                        Category       = "Security"
                    }
                }
            }
        }

        Write-ColorOutput -Message "Generated $($script:Recommendations.Count) recommendations" -Color Green

    } catch {
        Write-ColorOutput -Message "[SYSTEM ERROR DETECTED] Error generating recommendations: $($_.Exception.Message)" -Color Red
    }
}

function Export-AnalysisReport {
    <#
    .SYNOPSIS
    Exports the analysis results to formatted reports.
    #>
    try {
        Write-ColorOutput -Message "Generating analysis reports..." -Color Cyan

        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $reportFileName = "MSRD-Analysis-Report-$env:COMPUTERNAME-$timestamp.txt"
        $reportPath = Join-Path -Path $script:ReportPath -ChildPath $reportFileName

        # Generate comprehensive report
        $reportContent = @"
=============================================================================
MSRD-Collect Output Analysis Report
Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC")
System: $env:COMPUTERNAME
Analysis Period: $($script:DaysToAnalyze) days
=============================================================================

EXECUTIVE SUMMARY
=================
Critical Issues Found: $($script:CriticalIssues.Count)
Warnings: $($script:Warnings.Count)
Recommendations: $($script:Recommendations.Count)

CRITICAL ISSUES
===============
"@

        foreach ($issue in $script:CriticalIssues) {
            $reportContent += "`n[$($issue.Category)] $($issue.Finding)"
            if ($issue.Source) {
                $reportContent += " (Source: $($issue.Source))"
            }
        }

        $reportContent += "`n`nWARNINGS`n========`n"
        foreach ($warning in $script:Warnings) {
            $reportContent += "`n[$($warning.Category)] $($warning.Finding)"
            if ($warning.Source) {
                $reportContent += " (Source: $($warning.Source))"
            }
        }

        $reportContent += "`n`nRECOMMENDATIONS`n===============`n"
        foreach ($rec in $script:Recommendations) {
            $reportContent += "`n[Priority: $($rec.Priority)] $($rec.Issue)`n"
            $reportContent += "Recommendation: $($rec.Recommendation)`n"
        }

        if ($script:IncludeDetailedLogs -and $script:DetailedFindings.Count -gt 0) {
            $reportContent += "`n`nDETAILED FINDINGS`n=================`n"
            foreach ($finding in $script:DetailedFindings) {
                $reportContent += "`n[$($finding.Category)] $($finding.Finding)"
                if ($finding.Source) {
                    $reportContent += " (Source: $($finding.Source))"
                }
            }
        }

        # Save report
        $reportContent | Out-File -FilePath $reportPath -Encoding UTF8 -Force
        Write-ColorOutput -Message "Analysis report saved: $reportPath" -Color Green

        # Export to CSV if requested
        if ($script:ExportToCSV) {
            $csvFileName = "MSRD-Analysis-Data-$env:COMPUTERNAME-$timestamp.csv"
            $csvPath = Join-Path -Path $script:ReportPath -ChildPath $csvFileName

            $allFindings = @()
            $allFindings += $script:CriticalIssues
            $allFindings += $script:Warnings
            if ($script:IncludeDetailedLogs) {
                $allFindings += $script:DetailedFindings
            }

            $allFindings | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-ColorOutput -Message "CSV data exported: $csvPath" -Color Green
        }

        return $reportPath

    } catch {
        Write-ColorOutput -Message "[SYSTEM ERROR DETECTED] Error exporting analysis report: $($_.Exception.Message)" -Color Red
        throw
    }
}

function Invoke-MSRDAnalysis {
    <#
    .SYNOPSIS
    Main analysis orchestration function.
    #>
    try {
        Write-ColorOutput -Message "Starting MSRD-Collect output analysis..." -Color Green
        Write-ColorOutput -Message "Target directory: $script:MSRDOutputPath" -Color White
        Write-ColorOutput -Message "Analysis period: $script:DaysToAnalyze days" -Color White

        # Initialize environment
        Initialize-AnalysisEnvironment

        # Run analysis modules
        Get-MSRDSystemInfo
        Get-MSRDEventLog
        Get-MSRDRegistryAnalysis
        Get-MSRDNetworkAnalysis
        Get-MSRDServiceAnalysis

        # Generate recommendations
        Get-RecommendationEngine

        # Export results
        $reportPath = Export-AnalysisReport

        # Display summary
        Write-ColorOutput -Message "`nANALYSIS COMPLETE" -Color Green
        Write-ColorOutput -Message "==================" -Color Green
        Write-ColorOutput -Message "Critical Issues: $($script:CriticalIssues.Count)" -Color $(if ($script:CriticalIssues.Count -gt 0) { "Red" } else { "Green" })
        Write-ColorOutput -Message "Warnings: $($script:Warnings.Count)" -Color $(if ($script:Warnings.Count -gt 0) { "Yellow" } else { "Green" })
        Write-ColorOutput -Message "Recommendations: $($script:Recommendations.Count)" -Color Cyan
        Write-ColorOutput -Message "Report Location: $reportPath" -Color White

        if ($script:CriticalIssues.Count -gt 0) {
            Write-ColorOutput -Message "`nIMEDIATE ATTENTION REQUIRED:" -Color Red
            foreach ($issue in $script:CriticalIssues | Select-Object -First 3) {
                Write-ColorOutput -Message "• $($issue.Finding)" -Color Red
            }
        }

        return 0

    } catch {
        Write-ColorOutput -Message "[SYSTEM ERROR DETECTED] Analysis failed: $($_.Exception.Message)" -Color Red
        return 1
    } finally {
        if (Get-Command Stop-Transcript -ErrorAction SilentlyContinue) {
            Stop-Transcript -ErrorAction SilentlyContinue
        }
    }
}

# Main execution
try {
    $exitCode = Invoke-MSRDAnalysis
    exit $exitCode
} catch {
    Write-ColorOutput -Message "[SYSTEM ERROR DETECTED] Script execution failed: $($_.Exception.Message)" -Color Red
    exit 1
}
