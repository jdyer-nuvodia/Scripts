# =============================================================================
# Script: Analyze-WindowsLogs.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-11-10 22:22:42 UTC
# Updated By: jdyer-nuvodia
# Version: 1.3.1
# Additional Info: Fixed System.Web.HttpUtility dependency error by implementing native PowerShell HTML encoding
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
.PARAMETER LogDirectory
    Optional directory containing saved event log files (.evtx or .evt).
    When provided, the script analyzes these offline logs instead of live Windows logs.
.PARAMETER Recurse
    When using -LogDirectory, recurse into subfolders to find log files.
.PARAMETER TargetName
    Optional name to use in the report and output filenames when analyzing offline logs
    (e.g., the machine name the logs were collected from). Defaults to the current computer name.
.EXAMPLE
    .\Analyze-WindowsLogs.ps1
    Analyzes default logs for the past 30 days
.EXAMPLE
    .\Analyze-WindowsLogs.ps1 -DaysToAnalyze 7 -ReportPath "C:\Reports"
    Analyzes logs for the past week and saves report to specified location
.EXAMPLE
    .\Analyze-WindowsLogs.ps1 -LogDirectory "C:\CollectedLogs" -Recurse -TargetName "WS-1234"
    Analyzes all .evtx files found under C:\CollectedLogs and generates the same HTML report and log file,
    labeling results for WS-1234.
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
    [string[]]$LogNames = @('Application', 'System', 'Security'),

    [Parameter()]
    [ValidateScript({ Test-Path $_ })]
    [string]$LogDirectory,

    [Parameter()]
    [switch]$Recurse,

    [Parameter()]
    [string]$TargetName
)

# System name is resolved later (offline logs may specify a different machine)
$script:SystemName = if ([string]::IsNullOrWhiteSpace($TargetName)) { $env:COMPUTERNAME } else { $TargetName }
$script:LogFile     = $null
$script:ReportFile  = $null

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
        Write-Output "$colorCode$Message$resetCode"
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
    if ($null -ne $script:LogFile) {
        Add-Content -Path $script:LogFile -Value $LogMessage
    }

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

# Convert string to HTML-safe format
function ConvertTo-HtmlEncodedString {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$InputString
    )

    if ([string]::IsNullOrEmpty($InputString)) {
        return ""
    }

    $InputString = $InputString.Replace("&", "&amp;")
    $InputString = $InputString.Replace("<", "&lt;")
    $InputString = $InputString.Replace(">", "&gt;")
    $InputString = $InputString.Replace('"', "&quot;")
    $InputString = $InputString.Replace("'", "&#39;")

    return $InputString
}

# Attempt to resolve the machine name from offline log files when a directory is provided
function Resolve-TargetSystemName {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Directory,
        [Parameter()]
        [switch]$Recurse
    )

    try {
        $sampleFile = Get-ChildItem -Path (Join-Path -Path $Directory -ChildPath '*') -Include *.evtx,*.evt -File -Recurse:$Recurse | Select-Object -First 1
        if ($null -eq $sampleFile) { return $null }

        $sampleEvent = Get-WinEvent -Path $sampleFile.FullName -MaxEvents 1 -ErrorAction Stop
        if ($sampleEvent -and $sampleEvent.MachineName) {
            return $sampleEvent.MachineName
        }
    } catch {
        # Log and fall through to null if detection fails
        Write-LogMessage -Message "Failed to resolve target system name from offline logs: $_" -Level Debug -NoConsole
    }
    return $null
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
            Write-LogMessage "Attempting direct access to $LogName log..." -Level Process -NoConsole
            $Events = Get-WinEvent -LogName $LogName -MaxEvents 1 -ErrorAction Stop
            # If successful, proceed with full query
            $FilterHash = @{
                LogName = $LogName
                StartTime = $StartTime
            }
            $Events = Get-WinEvent -FilterHashtable $FilterHash -ErrorAction Stop
        } catch {
            Write-LogMessage "Direct access failed, trying alternative method for $LogName..." -Level Process -NoConsole

            # Method 2: Use System.Diagnostics.EventLog
            try {
                Write-LogMessage "Attempting System.Diagnostics.EventLog access..." -Level Process -NoConsole
                $EventLog = New-Object System.Diagnostics.EventLog($LogName)
                $Events = $EventLog.Entries | Where-Object { $_.TimeGenerated -ge $StartTime }
                $Events = $Events | ForEach-Object {
                    # Convert to equivalent WinEvent structure
                    [PSCustomObject]@{
                        TimeCreated = $_.TimeGenerated
                        Level = switch ($_.EntryType) {
                            'Error' { 2 }
                            'Warning' { 3 }
                            default { 4 }
                        }
                        ProviderName = $_.Source
                    }
                }
            } catch {
                Write-LogMessage "Alternative method also failed for $LogName. Error: $_" -Level Warning -NoConsole
                return $null
            }
        }

        if ($Events) {
            return Convert-EventToStatistic -LogName $LogName -Events $Events
        }
    } catch {
        Write-LogMessage "Error processing $LogName : $_" -Level Error
        return $null
    }
}

function Get-EventDetailKey {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Event
    )

    $provider = $Event.ProviderName
    $eventId  = $Event.Id
    $message  = $Event.Message

    $detail = $null

    try {
        switch ($provider) {
            'Service Control Manager' {
                if ($message) {
                    $patterns = @(
                        'The\s+(?<name>[^\.]*)\s+service\s',
                        'service name:\s*(?<name>[^\r\n\.,]+)'
                    )
                    foreach ($p in $patterns) {
                        $m = [regex]::Match($message, $p, 'IgnoreCase')
                        if ($m.Success) { $detail = $m.Groups['name'].Value.Trim(); break }
                    }
                }
                if (-not $detail -and $Event.Properties.Count -gt 0) { $detail = [string]$Event.Properties[0].Value }
            }
            '.NET Runtime' {
                if ($message) {
                    $patterns = @(
                        'Application:\s*(?<name>[^\r\n]+)',
                        'Description:\s*(?<name>[^\r\n]+)'
                    )
                    foreach ($p in $patterns) {
                        $m = [regex]::Match($message, $p, 'IgnoreCase')
                        if ($m.Success) { $detail = $m.Groups['name'].Value.Trim(); break }
                    }
                }
                if (-not $detail -and $Event.Properties.Count -gt 0) { $detail = [string]$Event.Properties[0].Value }
            }
            'Application Error' {
                if ($message) {
                    $m = [regex]::Match($message, 'Faulting application name:\s*(?<name>[^,\r\n]+)', 'IgnoreCase')
                    if ($m.Success) { $detail = $m.Groups['name'].Value.Trim() }
                }
                if (-not $detail -and $Event.Properties.Count -gt 0) { $detail = [string]$Event.Properties[0].Value }
            }
            'MsiInstaller' {
                if ($message) {
                    $m = [regex]::Match($message, 'Product:\s*(?<name>[^\-\r\n]+)', 'IgnoreCase')
                    if ($m.Success) { $detail = $m.Groups['name'].Value.Trim() }
                }
            }
            default {
                if ($Event.Properties -and $Event.Properties.Count -gt 0) {
                    $detail = [string]$Event.Properties[0].Value
                }
                if (-not $detail -and $message) {
                    $detail = ($message -split "\r?\n")[0]
                    if ($detail.Length -gt 80) { $detail = $detail.Substring(0, 80) + '...' }
                }
            }
        }
    } catch {
        $detail = $null
    }

    if ([string]::IsNullOrWhiteSpace($detail)) { $detail = 'Unspecified' }
    $key = "$provider ($eventId): $detail"
    return [PSCustomObject]@{ Key = $key; Provider = $provider; EventId = $eventId; Detail = $detail }
}

function Get-EventDetailCount {
    [CmdletBinding()]
    [OutputType([System.Object[]])]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Events,
        [Parameter(Mandatory = $true)]
        [int]$Level,
        [Parameter()]
        [int]$TopN = 10
    )

    $filtered = $Events | Where-Object { $_.Level -eq $Level }
    if (-not $filtered) { return @() }

    $keys = $filtered | ForEach-Object { Get-EventDetailKey -Event $_ }
    $groups = $keys | Group-Object -Property Key | Sort-Object -Property Count -Descending | Select-Object -First $TopN
    $results = foreach ($g in $groups) {
        $sample = $keys | Where-Object { $_.Key -eq $g.Name } | Select-Object -First 1
        [PSCustomObject]@{
            Provider = $sample.Provider
            EventId  = $sample.EventId
            Detail   = $sample.Detail
            Count    = $g.Count
        }
    }
    return $results
}

function Convert-EventToStatistic {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogName,
        [Parameter(Mandatory = $true)]
        [array]$Events
    )

    # Continue with existing event processing logic
    $TotalEntries   = ($Events | Measure-Object).Count
    $ErrorEntries   = ($Events | Where-Object { $_.Level -eq 2 } | Measure-Object).Count
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

    # Top warning sources
    $TopWarnings = $Events |
        Where-Object { $_.Level -eq 3 } |
        Group-Object {
            try {
                if ($_.ProviderName) { $_.ProviderName }
                else { "Unknown Provider" }
            } catch {
                "Unknown Provider"
            }
        } |
        Sort-Object -Property Count -Descending |
        Select-Object -First 5

    # Detailed aggregation for recurring issues (errors/warnings)
    $TopErrorDetails   = Get-EventDetailCount -Events $Events -Level 2 -TopN 10
    $TopWarningDetails = Get-EventDetailCount -Events $Events -Level 3 -TopN 10

    # Calculate daily entry rate
    $DailyRate = if ($script:DaysToAnalyze -gt 0) {
        [math]::Round($TotalEntries / $script:DaysToAnalyze, 2)
    } else {
        0
    }

    return @{
        Name               = $LogName
        TotalEntries       = $TotalEntries
        RecentEntries      = $TotalEntries
        ErrorCount         = $ErrorEntries
        WarningCount       = $WarningEntries
        DailyRate          = $DailyRate
        TopErrorSources    = $TopErrors
        TopWarningSources  = $TopWarnings
        TopErrorDetails    = $TopErrorDetails
        TopWarningDetails  = $TopWarningDetails
    }
}

function Get-SecurityEvent {
    [CmdletBinding(DefaultParameterSetName = 'Live')]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'Live')]
        [datetime]$StartTime,

        [Parameter(Mandatory = $true, ParameterSetName = 'FromEvents')]
        [array]$Events
    )

    try {
        $SecurityEvents = $null
        if ($PSCmdlet.ParameterSetName -eq 'FromEvents') {
            $SecurityEvents = $Events
        } else {
            $FilterHash = @{
                LogName   = 'Security'
                StartTime = $StartTime
            }
            $SecurityEvents = Get-WinEvent -FilterHashtable $FilterHash -ErrorAction Stop
        }

        if (-not $SecurityEvents) {
            throw [System.InvalidOperationException]::new('No security events available')
        }

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
            FailedLogins   = $FailedLogins
            AccountChanges = $AccountChanges
        }
    } catch [System.UnauthorizedAccessException] {
        Write-LogMessage "Access denied while analyzing security events. Please run as administrator." -Level Error
        return $null
    } catch [System.InvalidOperationException] {
        Write-LogMessage "No security events found in specified time range" -Level Warning
        return @{
            FailedLogins   = @()
            AccountChanges = @()
        }
    } catch {
        Write-LogMessage "Error analyzing security events: $_" -Level Error
        return $null
    }
}

function Get-DirectoryLogEvent {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Directory,
        [Parameter(Mandatory = $true)]
        [datetime]$StartTime,
        [Parameter()]
        [switch]$Recurse
    )

    Write-LogMessage "Scanning directory for log files: $Directory (Recurse: $($Recurse.IsPresent))" -Level Process -NoConsole

    # Use -Path with wildcard so -Include works; honor -Recurse switch
    $files = Get-ChildItem -Path (Join-Path $Directory '*') -Include *.evtx,*.evt -File -Recurse:$Recurse 2>$null
    if (-not $files -or $files.Count -eq 0) {
        Write-LogMessage "No .evtx or .evt files found in: $Directory" -Level Warning -NoConsole
        return @{}
    }

    $eventsByLog = @{}
    foreach ($file in $files) {
        try {
            Write-LogMessage "Reading events from file: $($file.FullName)" -Level Debug -NoConsole
            $fileEvents = Get-WinEvent -Path $file.FullName -ErrorAction Stop | Where-Object { $_.TimeCreated -ge $StartTime }
            if (-not $fileEvents) { continue }

            # Group by LogName if available; otherwise use filename as fallback
            $grouped = $fileEvents | Group-Object { if ($_.LogName) { $_.LogName } else { [System.IO.Path]::GetFileNameWithoutExtension($file.Name) } }
            foreach ($grp in $grouped) {
                if ($eventsByLog.ContainsKey($grp.Name)) {
                    $eventsByLog[$grp.Name] += $grp.Group
                } else {
                    $eventsByLog[$grp.Name] = @($grp.Group)
                }
            }
        } catch {
            Write-LogMessage "Failed to read events from file: $($file.FullName). Error: $_" -Level Warning -NoConsole
            continue
        }
    }

    return $eventsByLog
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
        h1, h2, h3 { color: #2c3e50; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
        th, td { padding: 8px; text-align: left; border: 1px solid #ddd; }
        th { background-color: #f5f5f5; }
        .error { color: #e74c3c; }
        .warning { color: #f39c12; }
        .success { color: #27ae60; }
        .muted { color: #7f8c8d; }
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
        <tr><td>Recent Entries ($script:DaysToAnalyze d)</td><td>$($stat.RecentEntries)</td></tr>
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

            # Top Warning Sources
            $HTMLBody += @"
    <h3>Top Warning Sources</h3>
    <table>
        <tr><th>Source</th><th>Count</th></tr>
"@
            foreach ($warnSource in $stat.TopWarningSources) {
                $HTMLBody += "<tr><td>$($warnSource.Name)</td><td>$($warnSource.Count)</td></tr>"
            }
            $HTMLBody += "</table>"

            # Detailed Error Breakdown
            $HTMLBody += @"
    <h3>Top Error Details</h3>
    <table>
        <tr><th>Provider</th><th>Event ID</th><th>Detail</th><th>Count</th></tr>
"@
            foreach ($err in $stat.TopErrorDetails) {
                $HTMLBody += "<tr><td>$($err.Provider)</td><td>$($err.EventId)</td><td>$(ConvertTo-HtmlEncodedString -InputString $err.Detail)</td><td>$($err.Count)</td></tr>"
            }
            $HTMLBody += "</table>"

            # Detailed Warning Breakdown
            $HTMLBody += @"
    <h3>Top Warning Details</h3>
    <table>
        <tr><th>Provider</th><th>Event ID</th><th>Detail</th><th>Count</th></tr>
"@
            foreach ($wd in $stat.TopWarningDetails) {
                $HTMLBody += "<tr><td>$($wd.Provider)</td><td>$($wd.EventId)</td><td>$(ConvertTo-HtmlEncodedString -InputString $wd.Detail)</td><td>$($wd.Count)</td></tr>"
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
    # Resolve final system name and initialize log/report files before first log write
    $TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"

    if ($PSBoundParameters.ContainsKey('LogDirectory') -and -not [string]::IsNullOrWhiteSpace($LogDirectory)) {
        # If no explicit TargetName was provided, try to read machine name from offline logs
        if (-not $PSBoundParameters.ContainsKey('TargetName') -or [string]::IsNullOrWhiteSpace($TargetName)) {
            $resolved = Resolve-TargetSystemName -Directory $LogDirectory -Recurse:$Recurse
            if ($resolved) { $script:SystemName = $resolved }
        }
    }

    $script:LogFile    = Join-Path -Path $ReportPath -ChildPath ("LogAnalysis_{0}_{1}.log" -f $script:SystemName, $TimeStamp)
    $script:ReportFile = Join-Path -Path $ReportPath -ChildPath ("LogAnalysis_{0}_{1}.html" -f $script:SystemName, $TimeStamp)

    Write-LogMessage -Message "Starting Windows log analysis..." -Level Process

    $StartTime = (Get-Date).AddDays(-$DaysToAnalyze)
    $LogStatistics = @()
    $hasValidData = $false
    $SecurityStats = $null

    if ($PSBoundParameters.ContainsKey('LogDirectory') -and -not [string]::IsNullOrWhiteSpace($LogDirectory)) {
        # Directory (offline) mode - skip admin requirement
        Write-LogMessage -Message "Running in offline directory mode against: $LogDirectory" -Level Info

        $eventsByLog = Get-DirectoryLogEvent -Directory $LogDirectory -StartTime $StartTime -Recurse:$Recurse
        if (-not $eventsByLog.Keys -or $eventsByLog.Keys.Count -eq 0) {
            Write-LogMessage "No events found in the specified directory and time range." -Level Error
            exit 1
        }

        # Determine which logs to analyze
        $logsToAnalyze = @()
        if ($PSBoundParameters.ContainsKey('LogNames')) {
            $logsToAnalyze = $LogNames | Where-Object { $eventsByLog.ContainsKey($_) }
            $missing = $LogNames | Where-Object { -not $eventsByLog.ContainsKey($_) }
            foreach ($m in $missing) { Write-LogMessage "Requested log '$m' not found in directory data; skipping." -Level Warning }
        } else {
            $logsToAnalyze = $eventsByLog.Keys
        }

        foreach ($log in $logsToAnalyze) {
            $logEvents = $eventsByLog[$log]
            if (-not $logEvents -or (($logEvents | Measure-Object).Count -eq 0)) {
                Write-LogMessage "No events available for '$log' within the specified time range; skipping." -Level Warning
                continue
            }

            Write-LogMessage -Message "Analyzing $log (offline) ..." -Level Process
            $stats = Convert-EventToStatistic -LogName $log -Events $logEvents
            if ($null -ne $stats) {
                $LogStatistics += $stats
                $hasValidData   = $true
                Write-LogMessage -Message "Completed analysis of $log (offline)" -Level Success
            }
        }

        # Security analysis from offline events if available
        if ($eventsByLog.ContainsKey('Security')) {
            $secEvents = $eventsByLog['Security']
            if ($secEvents -and (($secEvents | Measure-Object).Count -gt 0)) {
                Write-LogMessage -Message "Analyzing security events (offline) ..." -Level Process
                $SecurityStats = Get-SecurityEvent -Events $secEvents
            } else {
                Write-LogMessage -Message "Security log present but no events in the specified time range; skipping security analysis." -Level Warning
            }
        } else {
            Write-LogMessage -Message "No 'Security' log data found in directory; skipping security analysis." -Level Warning
        }
    } else {
        # Live mode - requires admin
        if (-not (Test-AdminPrivilege)) {
            Write-LogMessage -Message "This script requires administrator privileges. Please run as administrator." -Level Error
            exit 1
        }

        Write-LogMessage -Message "System: $script:SystemName" -Level Info
        Write-LogMessage -Message "Analysis period: $DaysToAnalyze days" -Level Info

        foreach ($LogName in $LogNames) {
            Write-LogMessage -Message "Analyzing $LogName log..." -Level Process
            $Stats = Get-LogStatistic -LogName $LogName
            if ($null -ne $Stats) {
                $LogStatistics += $Stats
                $hasValidData = $true
                Write-LogMessage -Message "Completed analysis of $LogName log" -Level Success
            }
        }

        Write-LogMessage -Message "Analyzing security events..." -Level Process
        $SecurityStats = Get-SecurityEvent -StartTime $StartTime
    }

    if (-not $hasValidData) {
        Write-LogMessage -Message "No valid log data could be collected. Please check inputs and try again." -Level Error
        exit 1
    }

    Write-LogMessage -Message "Generating HTML report..." -Level Process
    if ($null -eq $SecurityStats) {
        $SecurityStats = @{ FailedLogins = @(); AccountChanges = @() }
    }
    New-HTMLReport -LogStats $LogStatistics -SecurityStats $SecurityStats

    Write-LogMessage -Message "Analysis completed successfully" -Level Success
    Write-LogMessage -Message "Report saved to: $script:ReportFile" -Level Success
    Write-LogMessage -Message "Log file saved to: $script:LogFile" -Level Success
} catch {
    Write-LogMessage -Message "Script execution failed: $($_.Exception.Message)" -Level Error
    Write-LogMessage -Message "Stack trace: $($_.ScriptStackTrace)" -Level Debug
    exit 1
}
