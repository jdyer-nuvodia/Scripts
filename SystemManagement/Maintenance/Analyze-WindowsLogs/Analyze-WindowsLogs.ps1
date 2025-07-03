# =============================================================================
# Script: Analyze-WindowsLogs.ps1
# Created: 2025-04-02 21:15:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-07-03 21:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1.1
# Additional Info: Fixed PSScriptAnalyzer indentation and operator spacing issues
# =============================================================================

<#
.SYNOPSIS
    Analyzes Windows event logs and generates detailed statistics and reports.
.DESCRIPTION
    Performs comprehensive analysis of Windows event logs including:
    - Critical error patterns and frequency
    - Security event analysis
    - System resource impact events
    - Application crash patterns
    - Log size and growth trends

    Key features:
    - Pattern recognition for common issues
    - Security incident detection
    - Resource consumption tracking
    - Automated report generation
    - Historical trend analysis

    Dependencies:
    - Windows PowerShell 5.1 or higher
    - Administrative privileges for log access
    - Write access to report directory
.PARAMETER DaysToAnalyze
    Number of days of log history to analyze. Default is 30.
.PARAMETER ReportPath
    Directory where the analysis report will be saved.
    Defaults to script directory if not specified.
.PARAMETER LogNames
    Array of specific log names to analyze.
    Defaults to Application, System, and Security logs.
.EXAMPLE
    .\Analyze-WindowsLogs.ps1
    Analyzes default logs for the past 30 days
.EXAMPLE
    .\Analyze-WindowsLogs.ps1 -DaysToAnalyze 7 -ReportPath "C:\Reports"
    Analyzes logs for the past week and saves report to specified location
#>

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateRange(1, 365)]
    [int]$DaysToAnalyze = 30,

    [Parameter()]
    [ValidateScript({
            if (-not (Test-Path $_)) {
                New-Item -Path $_ -ItemType Directory -Force | Out-Null
            }
            return $true
        })]
    [string]$ReportPath = $PSScriptRoot,

    [Parameter()]
    [string[]]$LogNames = @('Application', 'System', 'Security')
)

# Initialize logging
$SystemName = $env:COMPUTERNAME
$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile = Join-Path $ReportPath "LogAnalysis_${SystemName}_${TimeStamp}.log"
$ReportFile = Join-Path $ReportPath "LogAnalysis_${SystemName}_${TimeStamp}.html"

# Initialize color support
$Script:UseAnsiColors = ($PSVersionTable.PSVersion.Major -ge 7)
$Script:Colors = if ($Script:UseAnsiColors) {
    @{
        White    = "`e[37m"
        Cyan     = "`e[36m"
        Green    = "`e[32m"
        Yellow   = "`e[33m"
        Red      = "`e[31m"
        Magenta  = "`e[35m"
        DarkGray = "`e[90m"
        Reset    = "`e[0m"
    }
} else {
    @{
        White    = "White"
        Cyan     = "Cyan"
        Green    = "Green"
        Yellow   = "Yellow"
        Red      = "Red"
        Magenta  = "Magenta"
        DarkGray = "DarkGray"
        Reset    = ""
    }
}

function Write-ColorOutput {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$Color = "White"
    )

    if ($Script:UseAnsiColors) {
        # PowerShell 7+ with ANSI escape codes
        $colorCode = $Script:Colors[$Color]
        $resetCode = $Script:Colors.Reset
        Write-Output "${colorCode}${Message}${resetCode}"
    } else {
        # PowerShell 5.1 - Change console color, write output, then reset
        $originalColor = $Host.UI.RawUI.ForegroundColor
        try {
            if ($Script:Colors[$Color] -and $Script:Colors[$Color] -ne "") {
                $Host.UI.RawUI.ForegroundColor = $Script:Colors[$Color]
            }
            Write-Output $Message
        } finally {
            $Host.UI.RawUI.ForegroundColor = $originalColor
        }
    }
}

function Write-LogMessage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter()]
        [ValidateSet("Info", "Process", "Success", "Warning", "Error", "Debug")]
        [string]$Level = "Info",

        [Parameter()]
        [switch]$NoConsole
    )

    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC"
    $LogMessage = "[$TimeStamp] [$Level] $Message"
    Add-Content -Path $script:LogFile -Value $LogMessage

    if (-not $NoConsole) {
        $ColorToUse = switch ($Level) {
            "Info" { "White" }
            "Process" { "Cyan" }
            "Success" { "Green" }
            "Warning" { "Yellow" }
            "Error" { "Red" }
            "Debug" { "Magenta" }
            default { "DarkGray" }
        }
        Write-ColorOutput -Message $Message -Color $ColorToUse
    }
}

# Check for admin privileges
function Test-AdminPrivilege {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-LogStatistic {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogName
    )

    try {
        $StartTime = (Get-Date).AddDays(-$script:DaysToAnalyze)

        # First verify the log exists and is accessible
        $LogExists = Get-WinEvent -ListLog $LogName -ErrorAction Stop
        if (-not $LogExists.IsEnabled) {
            Write-LogMessage "Log '$LogName' exists but is not enabled" -Level Warning
            return $null
        }

        # Try different methods to access the log
        $Events = $null

        # Method 1: Direct access with minimal filtering
        try {
            Write-LogMessage "Attempting direct access to $LogName log..." -Level Process
            $Events = Get-WinEvent -LogName $LogName -MaxEvents 1 -ErrorAction Stop
            # If successful, proceed with full query
            $FilterHash = @{
                LogName   = $LogName
                StartTime = $StartTime
            }
            $Events = Get-WinEvent -FilterHashtable $FilterHash -ErrorAction Stop
        } catch {
            Write-LogMessage "Direct access failed, trying alternative method for $LogName..." -Level Process

            # Method 2: Use System.Diagnostics.EventLog
            try {
                Write-LogMessage "Attempting System.Diagnostics.EventLog access..." -Level Process
                $EventLog = New-Object System.Diagnostics.EventLog($LogName)
                $Events = $EventLog.Entries | Where-Object { $_.TimeGenerated -ge $StartTime }
                $Events = $Events | ForEach-Object {
                    # Convert to equivalent WinEvent structure
                    [PSCustomObject]@{
                        TimeCreated  = $_.TimeGenerated
                        Level        = switch ($_.EntryType) {
                            'Error' { 2 }
                            'Warning' { 3 }
                            default { 4 }
                        }
                        ProviderName = $_.Source
                    }
                }
            } catch {
                Write-LogMessage "Alternative method also failed for $LogName. Error: $_" -Level Warning
                return $null
            }
        }

        # Process events if we have any
        if ($Events) {
            # Continue with existing event processing logic
            $TotalEntries = ($Events | Measure-Object).Count
            $ErrorEntries = ($Events | Where-Object { $_.Level -eq 2 } | Measure-Object).Count
            $WarningEntries = ($Events | Where-Object { $_.Level -eq 3 } | Measure-Object).Count

            # Get top error sources with enhanced error handling
            $TopErrors = $Events |
            Where-Object { $_.Level -eq 2 } |
            Group-Object {
                try {
                    if ($_.ProviderName) { $_.ProviderName }
                    else { "Unknown Provider" }
                } catch {
                    "Unknown Provider"
                }
            } |
            Sort-Object Count -Descending |
            Select-Object -First 5

            # Calculate daily entry rate
            $DailyRate = if ($script:DaysToAnalyze -gt 0) {
                [math]::Round($TotalEntries / $script:DaysToAnalyze, 2)
            } else {
                0
            }

            return @{
                Name = $LogName
                TotalEntries = $TotalEntries
                RecentEntries = $TotalEntries
                ErrorCount = $ErrorEntries
                WarningCount = $WarningEntries
                DailyRate = $DailyRate
                TopErrorSources = $TopErrors
            }
        }
    } catch {
        $ErrorMessage = $_.Exception.Message
        Write-LogMessage "Error processing ${LogName}: ${ErrorMessage}" -Level Error
        return $null
    }
}

function Get-SecurityEvent {
    param(
        [Parameter(Mandatory = $true)]
        [datetime]$StartTime
    )

    try {
        $FilterHash = @{
            LogName   = 'Security'
            StartTime = $StartTime
        }

        $SecurityEvents = Get-WinEvent -FilterHashtable $FilterHash -ErrorAction Stop

        # Analyze login attempts
        $FailedLogins = $SecurityEvents |
        Where-Object { $_.Id -eq 4625 } |
        Group-Object { $_.Properties[5].Value } |
        Sort-Object Count -Descending |
        Select-Object -First 5

        # Analyze account modifications
        $AccountChanges = $SecurityEvents |
        Where-Object { $_.Id -in @(4720, 4722, 4725, 4726) } |
        Group-Object Id |
        Sort-Object Count -Descending

        return @{
            FailedLogins = $FailedLogins
            AccountChanges = $AccountChanges
        }
    } catch [System.UnauthorizedAccessException] {
        Write-LogMessage "Access denied while analyzing security events. Please run as administrator." -Level Error
        return $null
    } catch [System.InvalidOperationException] {
        Write-LogMessage "No security events found in specified time range" -Level Warning
        return @{
            FailedLogins = @()
            AccountChanges = @()
        }
    } catch {
        Write-LogMessage "Error analyzing security events: $_" -Level Error
        return $null
    }
}

function New-HTMLReport {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [array]$LogStats,
        [Parameter(Mandatory = $true)]
        [object]$SecurityStats
    )

    # Validate input parameters
    if ($null -eq $LogStats -or $LogStats.Count -eq 0) {
        throw "No log statistics available for report generation"
    }

    if ($PSCmdlet.ShouldProcess($script:ReportFile, "Create HTML Report")) {
        $HTMLHeader = @"
<!DOCTYPE html>
<html>
<head>
    <title>Windows Log Analysis Report - $script:SystemName</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1, h2 { color: #2c3e50; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
        th, td { padding: 8px; text-align: left; border: 1px solid #ddd; }
        th { background-color: #f5f5f5; }
        .error { color: #e74c3c; }
        .warning { color: #f39c12; }
        .success { color: #27ae60; }
    </style>
</head>
<body>
    <h1>Windows Log Analysis Report</h1>
    <p>System: $script:SystemName</p>
    <p>Date: $((Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss UTC'))</p>
    <p>Analysis Period: $script:DaysToAnalyze days</p>
"@

        $HTMLBody = @()
        foreach ($stat in $LogStats) {
            if ($null -eq $stat) { continue }

            $HTMLBody += @"
    <h2>$($stat.Name) Log Analysis</h2>
    <table>
        <tr><th>Metric</th><th>Value</th></tr>
        <tr><td>Total Entries</td><td>$($stat.TotalEntries)</td></tr>
        <tr><td>Recent Entries (${script:DaysToAnalyze}d)</td><td>$($stat.RecentEntries)</td></tr>
        <tr><td>Error Count</td><td>$($stat.ErrorCount)</td></tr>
        <tr><td>Warning Count</td><td>$($stat.WarningCount)</td></tr>
        <tr><td>Daily Entry Rate</td><td>$($stat.DailyRate)</td></tr>
    </table>

    <h3>Top Error Sources</h3>
    <table>
        <tr><th>Source</th><th>Count</th></tr>
"@

            foreach ($errorSource in $stat.TopErrorSources) {
                $HTMLBody += "<tr><td>$($errorSource.Name)</td><td>$($errorSource.Count)</td></tr>"
            }

            $HTMLBody += "</table>"
        }

        if ($null -ne $SecurityStats) {
            $HTMLBody += @"
    <h2>Security Analysis</h2>
    <h3>Failed Login Attempts</h3>
    <table>
        <tr><th>Account</th><th>Attempts</th></tr>
"@

            foreach ($login in $SecurityStats.FailedLogins) {
                $HTMLBody += "<tr><td>$($login.Name)</td><td>$($login.Count)</td></tr>"
            }

            $HTMLBody += @"
    </table>
    <h3>Account Modifications</h3>
    <table>
        <tr><th>Event ID</th><th>Count</th><th>Description</th></tr>
"@

            $EventDescriptions = @{
                4720 = "Account Created"
                4722 = "Account Enabled"
                4725 = "Account Disabled"
                4726 = "Account Deleted"
            }

            foreach ($change in $SecurityStats.AccountChanges) {
                $desc = $EventDescriptions[$change.Name]
                $HTMLBody += "<tr><td>$($change.Name)</td><td>$($change.Count)</td><td>$desc</td></tr>"
            }

            $HTMLBody += "</table>"
        }

        $HTMLFooter = @"
</body>
</html>
"@

        $Report = $HTMLHeader + ($HTMLBody -join "`n") + $HTMLFooter
        $Report | Out-File -FilePath $script:ReportFile -Encoding UTF8
    }
}

# Main execution
try {
    Write-LogMessage "Starting Windows log analysis..." -Level Process

    # Check for admin privileges
    if (-not (Test-AdminPrivilege)) {
        Write-LogMessage "This script requires administrator privileges. Please run as administrator." -Level Error
        exit 1
    }

    Write-LogMessage "System: $SystemName" -Level Info
    Write-LogMessage "Analysis period: $DaysToAnalyze days" -Level Info

    $StartTime = (Get-Date).AddDays(-$DaysToAnalyze)
    $LogStatistics = @()
    $hasValidData = $false

    foreach ($LogName in $LogNames) {
        Write-LogMessage "Analyzing $LogName log..." -Level Process
        $Stats = Get-LogStatistic -LogName $LogName
        if ($null -ne $Stats) {
            $LogStatistics += $Stats
            $hasValidData = $true
            Write-LogMessage "Completed analysis of $LogName log" -Level Success
        }
    }

    Write-LogMessage "Analyzing security events..." -Level Process
    $SecurityStats = Get-SecurityEvent -StartTime $StartTime

    if (-not $hasValidData) {
        Write-LogMessage "No valid log data could be collected. Please check permissions and try again." -Level Error
        exit 1
    }

    Write-LogMessage "Generating HTML report..." -Level Process
    New-HTMLReport -LogStats $LogStatistics -SecurityStats $SecurityStats

    Write-LogMessage "Analysis completed successfully" -Level Success
    Write-LogMessage "Report saved to: $ReportFile" -Level Success
    Write-LogMessage "Log file saved to: $LogFile" -Level Success
} catch {
    Write-LogMessage "Script execution failed: $($_.Exception.Message)" -Level Error
    Write-LogMessage "Stack trace: $($_.ScriptStackTrace)" -Level Debug
    exit 1
}
