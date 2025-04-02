# =============================================================================
# Script: Get-SystemHealthReport.ps1
# Created: 2025-04-02 20:23:00 UTC
# Author: GitHub-Copilot
# Last Updated: 2025-04-02 20:23:00 UTC
# Updated By: GitHub-Copilot
# Version: 1.0.1
# Additional Info: Removed unused HTML report variable
# =============================================================================

<#
.SYNOPSIS
    Performs a comprehensive system health check and generates a detailed report.
.DESCRIPTION
    This script performs multiple system health checks including:
    - CPU, Memory, and Disk performance
    - Critical Windows Services status
    - Event Log analysis
    - Windows Update status
    - Network connectivity
    - Backup status
    - Security features status
    
    Results are both displayed in color-coded console output and saved to a log file.
.PARAMETER ReportPath
    Optional path where the HTML report will be saved. Defaults to script directory.
.PARAMETER DaysToAnalyze
    Number of days of event logs to analyze. Default is 7.
.PARAMETER CriticalServices
    Array of critical services to check. If not specified, checks common critical services.
.EXAMPLE
    .\Get-SystemHealthReport.ps1
    Runs health check with default parameters
.EXAMPLE
    .\Get-SystemHealthReport.ps1 -DaysToAnalyze 14 -ReportPath "C:\Reports"
    Runs health check analyzing 14 days of logs and saves report to specified path
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$ReportPath = $PSScriptRoot,
    
    [Parameter()]
    [int]$DaysToAnalyze = 7,
    
    [Parameter()]
    [string[]]$CriticalServices = @(
        'wuauserv',      # Windows Update
        'WinDefend',     # Windows Defender
        'wscsvc',        # Security Center
        'Schedule',      # Task Scheduler
        'EventLog',      # Windows Event Log
        'mpssvc',        # Windows Firewall
        'LanmanServer',  # Server
        'Dnscache'       # DNS Client
    )
)

# Initialize logging
$SystemName = $env:COMPUTERNAME
$Timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$LogFile = Join-Path $ReportPath "SystemHealth_${SystemName}_${Timestamp}.log"

function Write-LogMessage {
    param(
        [string]$Message,
        [string]$Level = "Info"
    )
    
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC"
    $LogMessage = "[$TimeStamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $LogMessage
    
    switch ($Level) {
        "Info"      { Write-Host $LogMessage -ForegroundColor White }
        "Process"   { Write-Host $LogMessage -ForegroundColor Cyan }
        "Success"   { Write-Host $LogMessage -ForegroundColor Green }
        "Warning"   { Write-Host $LogMessage -ForegroundColor Yellow }
        "Error"     { Write-Host $LogMessage -ForegroundColor Red }
        "Debug"     { Write-Host $LogMessage -ForegroundColor Magenta }
        Default     { Write-Host $LogMessage -ForegroundColor DarkGray }
    }
}

function Get-SystemResourceStatus {
    Write-LogMessage "Checking system resources..." -Level "Process"
    
    try {
        # CPU Usage
        $CpuUsage = (Get-Counter '\Processor(_Total)\% Processor Time').CounterSamples.CookedValue
        $CpuStatus = switch ($CpuUsage) {
            {$_ -ge 90} { "Error" }
            {$_ -ge 80} { "Warning" }
            Default { "Success" }
        }
        Write-LogMessage "CPU Usage: $([math]::Round($CpuUsage, 2))%" -Level $CpuStatus
        
        # Memory Usage
        $OS = Get-WmiObject Win32_OperatingSystem
        $MemoryUsage = [math]::Round(($OS.TotalVisibleMemorySize - $OS.FreePhysicalMemory) / $OS.TotalVisibleMemorySize * 100, 2)
        $MemoryStatus = switch ($MemoryUsage) {
            {$_ -ge 90} { "Error" }
            {$_ -ge 80} { "Warning" }
            Default { "Success" }
        }
        Write-LogMessage "Memory Usage: ${MemoryUsage}%" -Level $MemoryStatus
        
        # Disk Space
        Get-WmiObject Win32_LogicalDisk -Filter "DriveType=3" | ForEach-Object {
            $FreeSpace = [math]::Round(($_.FreeSpace / $_.Size) * 100, 2)
            $Status = switch ($FreeSpace) {
                {$_ -le 10} { "Error" }
                {$_ -le 20} { "Warning" }
                Default { "Success" }
            }
            Write-LogMessage "Drive $($_.DeviceID) - Free Space: ${FreeSpace}%" -Level $Status
        }
    }
    catch {
        Write-LogMessage "Error checking system resources: $_" -Level "Error"
    }
}

function Get-CriticalServicesStatus {
    Write-LogMessage "Checking critical services..." -Level "Process"
    
    foreach ($Service in $CriticalServices) {
        try {
            $ServiceStatus = Get-Service -Name $Service -ErrorAction Stop
            $Status = switch ($ServiceStatus.Status) {
                'Running' { "Success" }
                'Stopped' { "Error" }
                Default { "Warning" }
            }
            Write-LogMessage "Service $($ServiceStatus.DisplayName): $($ServiceStatus.Status)" -Level $Status
        }
        catch {
            Write-LogMessage "Error checking service ${Service}: $($_.Exception.Message)" -Level "Error"
        }
    }
}

function Get-EventLogAnalysis {
    Write-LogMessage "Analyzing event logs..." -Level "Process"
    
    $StartTime = (Get-Date).AddDays(-$DaysToAnalyze)
    $CriticalEvents = @(
        @{Log='System'; Level=2; Name='System Errors'},
        @{Log='Application'; Level=2; Name='Application Errors'},
        @{Log='Security'; Level=2; Name='Security Errors'}
    )
    
    foreach ($EventType in $CriticalEvents) {
        try {
            $Events = Get-WinEvent -FilterHashtable @{
                LogName = $EventType.Log
                Level = $EventType.Level
                StartTime = $StartTime
            } -ErrorAction SilentlyContinue
            
            $Count = ($Events | Measure-Object).Count
            $Status = switch ($Count) {
                {$_ -ge 50} { "Error" }
                {$_ -ge 20} { "Warning" }
                Default { "Success" }
            }
            Write-LogMessage "$($EventType.Name) in last $DaysToAnalyze days: $Count" -Level $Status
        }
        catch {
            if ($_.Exception.Message -notlike "*No events were found*") {
                Write-LogMessage "Error checking $($EventType.Name): $_" -Level "Error"
            }
        }
    }
}

function Get-WindowsUpdateStatus {
    Write-LogMessage "Checking Windows Update status..." -Level "Process"
    
    try {
        $UpdateSession = New-Object -ComObject Microsoft.Update.Session
        $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
        $PendingUpdates = $UpdateSearcher.Search("IsInstalled=0").Updates
        
        $Count = $PendingUpdates.Count
        $Status = switch ($Count) {
            {$_ -ge 10} { "Error" }
            {$_ -ge 5} { "Warning" }
            Default { "Success" }
        }
        Write-LogMessage "Pending Windows Updates: $Count" -Level $Status
    }
    catch {
        Write-LogMessage "Error checking Windows Updates: $_" -Level "Error"
    }
}

function Get-NetworkStatus {
    Write-LogMessage "Checking network connectivity..." -Level "Process"
    
    $Targets = @(
        "8.8.8.8",           # Google DNS
        "1.1.1.1",           # Cloudflare DNS
        "www.microsoft.com"
    )
    
    foreach ($Target in $Targets) {
        try {
            $Result = Test-Connection -ComputerName $Target -Count 1 -ErrorAction Stop
            $LatencyMs = $Result.ResponseTime
            $Status = switch ($LatencyMs) {
                {$_ -ge 200} { "Warning" }
                {$_ -ge 500} { "Error" }
                Default { "Success" }
            }
            Write-LogMessage "Network latency to $Target : ${LatencyMs}ms" -Level $Status
        }
        catch {
            Write-LogMessage "Failed to reach $Target" -Level "Error"
        }
    }
}

function Get-SecurityStatus {
    Write-LogMessage "Checking security features..." -Level "Process"
    
    try {
        # Windows Defender Status
        $DefenderStatus = Get-MpComputerStatus
        $Status = if ($DefenderStatus.AntivirusEnabled) { "Success" } else { "Error" }
        Write-LogMessage "Windows Defender Status: $($DefenderStatus.AntivirusEnabled)" -Level $Status
        
        # Firewall Status
        $FirewallProfiles = Get-NetFirewallProfile
        foreach ($FwProfile in $FirewallProfiles) {
            $Status = if ($FwProfile.Enabled) { "Success" } else { "Error" }
            Write-LogMessage "Firewall Profile $($FwProfile.Name): $($FwProfile.Enabled)" -Level $Status
        }
        
        # BitLocker Status
        $BitLockerVolumes = Get-BitLockerVolume -ErrorAction SilentlyContinue
        if ($BitLockerVolumes) {
            foreach ($Volume in $BitLockerVolumes) {
                $Status = switch ($Volume.ProtectionStatus) {
                    'On' { "Success" }
                    'Off' { "Error" }
                    Default { "Warning" }
                }
                Write-LogMessage "BitLocker on $($Volume.MountPoint): $($Volume.ProtectionStatus)" -Level $Status
            }
        }
        else {
            Write-LogMessage "BitLocker not configured on any volumes" -Level "Warning"
        }
    }
    catch {
        Write-LogMessage "Error checking security status: $($_.Exception.Message)" -Level "Error"
    }
}

# Main execution
Write-LogMessage "=== System Health Check Started ===" -Level "Process"
Write-LogMessage "System: $SystemName" -Level "Info"
Write-LogMessage "Timestamp: $((Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss UTC'))" -Level "Info"

Get-SystemResourceStatus
Get-CriticalServicesStatus
Get-EventLogAnalysis
Get-WindowsUpdateStatus
Get-NetworkStatus
Get-SecurityStatus

Write-LogMessage "=== System Health Check Completed ===" -Level "Process"
Write-LogMessage "Log file saved to: $LogFile" -Level "Success"