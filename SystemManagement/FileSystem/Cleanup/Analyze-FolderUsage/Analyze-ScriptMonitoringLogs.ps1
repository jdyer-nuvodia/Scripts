# =============================================================================
# Script: Analyze-ScriptMonitoringLogs.ps1
# Created: 2025-06-20 18:05:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-06-20 18:05:00 UTC
# Updated By: jdyer-nuvodia
# Version: 2.0.0
# Additional Info: Updated for work-stealing queue system compatibility
# =============================================================================

<#
.SYNOPSIS
Comprehensive analysis tool for Analyze-FolderUsage.ps1 monitoring logs with work-stealing queue support.

.DESCRIPTION
This script analyzes the enhanced central debug logs from Analyze-FolderUsage.ps1 v5.0.0+ to provide
detailed insights into the work-stealing queue system performance, thread utilization, job distribution,
and overall efficiency. It provides specialized analysis for the new parallel processing enhancements.

Key Analysis Areas:
- Work-stealing queue performance and job distribution
- Thread utilization efficiency and peak performance
- Job completion timing and bottleneck identification
- Error detection and accessibility issues
- Memory and performance optimization insights

.PARAMETER LogFile
Path to the central debug log file to analyze. If not specified, uses the most recent log file.

.PARAMETER ShowDetailed
Switch to show detailed analysis including individual job information.

.PARAMETER ExportResults
Switch to export analysis results to a CSV file.

.EXAMPLE
.\Analyze-ScriptMonitoringLogs.ps1
Analyzes the most recent central debug log file with standard reporting.

.EXAMPLE
.\Analyze-ScriptMonitoringLogs.ps1 -LogFile "Analyze-FolderUsage_CENTRAL_LT-JBDYER_2025-06-20_12-04-24.log" -ShowDetailed
Analyzes the specified log file with detailed job-level information.

.EXAMPLE
.\Analyze-ScriptMonitoringLogs.ps1 -ExportResults
Analyzes the most recent log and exports results to CSV for further analysis.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$LogFile,
    [Parameter(Mandatory = $false)]
    [switch]$ShowDetailed,
    [Parameter(Mandatory = $false)]
    [switch]$ExportResults
)

# Determine which log file to analyze
if (-not $LogFile) {
    $logFileObj = Get-ChildItem -Path ".\Analyze-FolderUsage_CENTRAL_*.log" -ErrorAction SilentlyContinue |
                  Sort-Object LastWriteTime |
                  Select-Object -Last 1

    if (-not $logFileObj) {
        Write-Error "No central debug log files found. Please run Analyze-FolderUsage.ps1 with -Debug parameter first."
        exit 1
    }
    $LogFile = $logFileObj.FullName
} elseif (-not (Test-Path -Path $LogFile)) {
    Write-Error "Specified log file '$LogFile' not found."
    exit 1
}

Write-Output "`n==============================================="
Write-Output "    WORK-STEALING QUEUE ANALYSIS REPORT v2.0.0"
Write-Output "==============================================="
Write-Output "Log File: $(Split-Path -Leaf $LogFile)"
Write-Output "Analysis Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC')"
Write-Output "==============================================="

# Helper function to extract log messages for a specific category
function Get-LogCategory {
    param([string]$Category)
    return Select-String -Path $LogFile -Pattern "\[$Category\]" | ForEach-Object {
        $_.Line -replace '^.*\[.*?\]\s*\[.*?\]\s*\[.*?\]\s*', ''
    }
}

# Helper function to extract timestamp and value patterns
function Get-LogPattern {
    param([string]$Pattern)
    return Select-String -Path $LogFile -Pattern $Pattern | ForEach-Object {
        $_.Line -replace '^.*\[.*?\]\s*\[.*?\]\s*\[.*?\]\s*', ''
    }
}

# 1. WORK QUEUE SYSTEM CONFIGURATION
Write-Output "`n WORK-STEALING QUEUE CONFIGURATION:"
$config = Get-LogCategory -Category "PARALLEL_MONITOR" | Where-Object { $_ -match "configured with MaxThreads|Initial directories|Work queue system" }
if ($config) {
    $config | ForEach-Object { Write-Output "  $_" }
} else {
    Write-Output "  Configuration data not found (may be using older version)"
}

# 2. JOB DISTRIBUTION ANALYSIS
Write-Output "`n INTELLIGENT JOB DISTRIBUTION:"
$jobDist = Get-LogCategory -Category "JOB_DISTRIBUTION"
if ($jobDist) {
    $jobDist | Select-Object -First 10 | ForEach-Object { Write-Output "  $_" }
    if ($jobDist.Count -gt 10) {
        Write-Output "  ... ($(($jobDist.Count - 10)) more distribution entries)"
    }
} else {
    Write-Output "  No job distribution data found"
}

# 3. WORK QUEUE PERFORMANCE
Write-Output "`n WORK QUEUE PERFORMANCE:"
$workQueue = Get-LogCategory -Category "WORK_QUEUE"
if ($workQueue) {
    $queueInitialized = $workQueue | Where-Object { $_ -match "Work queue initialized" }
    $jobsLaunched = $workQueue | Where-Object { $_ -match "Job #.*launched" }
    $newItemsAdded = $workQueue | Where-Object { $_ -match "Added.*new subdirectories" }

    Write-Output "  Queue Initialization: $(if ($queueInitialized) { $queueInitialized[0] } else { 'Not found' })"
    Write-Output "  Jobs Launched: $($jobsLaunched.Count) total"
    Write-Output "  Dynamic Subdirectory Additions: $($newItemsAdded.Count) batches"

    if ($ShowDetailed -and $jobsLaunched.Count -gt 0) {
        Write-Output "`n   DETAILED JOB LAUNCH SEQUENCE:"
        $jobsLaunched | Select-Object -First 15 | ForEach-Object { Write-Output "    $_" }
        if ($jobsLaunched.Count -gt 15) {
            Write-Output "    ... ($(($jobsLaunched.Count - 15)) more jobs launched)"
        }
    }
} else {
    Write-Output "  No work queue performance data found"
}

# 4. THREAD UTILIZATION ANALYSIS
Write-Output "`n THREAD UTILIZATION ANALYSIS:"
$utilization = Get-LogCategory -Category "PARALLEL_MONITOR" | Where-Object { $_ -match "NEW PEAK|peak utilization|Active jobs.*utilization" }
if ($utilization) {
    $peakJobs = $utilization | Where-Object { $_ -match "NEW PEAK" }
    $finalUtilization = $utilization | Where-Object { $_ -match "peak utilization" } | Select-Object -Last 1
    $activeReports = $utilization | Where-Object { $_ -match "Active jobs.*utilization" }

    Write-Output "  Peak Concurrent Jobs: $(if ($peakJobs) { ($peakJobs | Select-Object -Last 1) } else { 'Not recorded' })"
    Write-Output "  Final Utilization: $(if ($finalUtilization) { $finalUtilization } else { 'Not calculated' })"
    Write-Output "  Active Job Reports: $($activeReports.Count) monitoring intervals"

    if ($ShowDetailed -and $activeReports.Count -gt 0) {
        Write-Output "`n   UTILIZATION OVER TIME:"
        $activeReports | Select-Object -First 10 | ForEach-Object { Write-Output "    $_" }
        if ($activeReports.Count -gt 10) {
            Write-Output "    ... ($(($activeReports.Count - 10)) more monitoring reports)"
        }
    }
} else {
    Write-Output "  No thread utilization data found"
}

# 5. PERFORMANCE TIMING ANALYSIS
Write-Output "`n PERFORMANCE TIMING:"
$timing = Get-LogCategory -Category "PARALLEL_MONITOR" | Where-Object { $_ -match "Total execution time|Average job duration|completed.*in.*s$" }
if ($timing) {
    $executionTime = $timing | Where-Object { $_ -match "Total execution time" }
    $avgDuration = $timing | Where-Object { $_ -match "Average job duration" }
    $completedJobs = $timing | Where-Object { $_ -match "completed.*in.*s$" }

    Write-Output "  Total Execution Time: $(if ($executionTime) { $executionTime } else { 'Not recorded' })"
    Write-Output "  Average Job Duration: $(if ($avgDuration) { $avgDuration } else { 'Not calculated' })"
    Write-Output "  Individual Job Completions: $($completedJobs.Count) jobs timed"

    if ($ShowDetailed -and $completedJobs.Count -gt 0) {
        # Show fastest and slowest jobs
        $jobTimes = $completedJobs | ForEach-Object {
            if ($_ -match 'completed.*in (\d+\.?\d*?)s$') {
                [PSCustomObject]@{
                    Job = $_
                    Duration = [double]$matches[1]
                }
            }
        } | Where-Object { $_.Duration -ne $null } | Sort-Object Duration

        if ($jobTimes.Count -gt 0) {
            Write-Output "`n   FASTEST JOBS:"
            $jobTimes | Select-Object -First 5 | ForEach-Object { Write-Output "    $($_.Job)" }

            Write-Output "`n   SLOWEST JOBS:"
            $jobTimes | Select-Object -Last 5 | ForEach-Object { Write-Output "    $($_.Job)" }
        }
    }
} else {
    Write-Output "  No performance timing data found"
}

# 6. EFFICIENCY ANALYSIS
Write-Output "`n EFFICIENCY ANALYSIS:"
$efficiency = Get-LogPattern -Pattern "EFFICIENCY.*:"
if ($efficiency) {
    $efficiency | ForEach-Object { Write-Output "  $_" }
} else {
    Write-Output "  No efficiency analysis found"
}

# 7. ERROR AND ISSUE DETECTION
Write-Output "`n ISSUES AND ERRORS DETECTED:"
$issues = @()
$issues += Get-LogPattern -Pattern "WARNING.*:"
$issues += Get-LogPattern -Pattern "ERROR.*:"
$issues += Get-LogPattern -Pattern "TIMEOUT.*:"
$issues += Get-LogPattern -Pattern "CRITICAL.*:"
$issues += Get-LogCategory -Category "FOLDER_ERROR"

if ($issues.Count -gt 0) {
    $issues | Select-Object -First 20 | ForEach-Object { Write-Output "  $_" }
    if ($issues.Count -gt 20) {
        Write-Output "  ... ($(($issues.Count - 20)) more issues detected)"
    }
} else {
    Write-Output "   No issues detected"
}

# 8. FINAL SUMMARY REPORT
Write-Output "`n FINAL EXECUTION SUMMARY:"
$finalReport = Get-LogCategory -Category "PARALLEL_MONITOR" | Where-Object { $_ -match "FINAL WORK QUEUE EXECUTION REPORT" -or $_ -match "Total jobs launched.*dynamic" -or $_ -match "Work queue final state" }
if ($finalReport) {
    $finalReport | ForEach-Object { Write-Output "  $_" }
} else {
    Write-Output "  Final summary not found"
}

# 9. ACCESSIBILITY AND ERROR SUMMARY
Write-Output "`n ACCESSIBILITY ANALYSIS:"
$accessibility = Get-LogCategory -Category "ACCESSIBILITY"
if ($accessibility) {
    $accessibility | Select-Object -First 10 | ForEach-Object { Write-Output "  $_" }
    if ($accessibility.Count -gt 10) {
        Write-Output "  ... ($(($accessibility.Count - 10)) more accessibility entries)"
    }
} else {
    Write-Output "  No accessibility data found"
}

# 10. EXPORT RESULTS (if requested)
if ($ExportResults) {
    $exportFile = "MonitoringAnalysis_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').csv"
    Write-Output "`n EXPORTING DETAILED RESULTS TO: $exportFile"

    # Create comprehensive export data
    $exportData = @()

    # Add timing data
    $completedJobs = Get-LogCategory -Category "PARALLEL_MONITOR" | Where-Object { $_ -match "Job #.*completed.*in.*s$" }
    foreach ($job in $completedJobs) {
        if ($job -match 'Job #(\d+) completed: (.*?) \(Depth: (\d+)\) in ([\d.]+)s') {
            $exportData += [PSCustomObject]@{
                Type = "JobCompletion"
                JobNumber = $matches[1]
                Path = $matches[2]
                Depth = $matches[3]
                Duration = [double]$matches[4]
                Category = "Performance"
            }
        }
    }

    # Add error data
    foreach ($issue in $issues) {
        $exportData += [PSCustomObject]@{
            Type = "Issue"
            JobNumber = ""
            Path = ""
            Depth = ""
            Duration = ""
            Category = "Error"
            Description = $issue
        }
    }

    if ($exportData.Count -gt 0) {
        $exportData | Export-Csv -Path $exportFile -NoTypeInformation -Encoding UTF8
        Write-Output "   Export completed: $($exportData.Count) records"
    } else {
        Write-Output "   No data available for export"
    }
}

Write-Output "`n==============================================="
Write-Output "Analysis completed successfully!"
Write-Output "==============================================="


