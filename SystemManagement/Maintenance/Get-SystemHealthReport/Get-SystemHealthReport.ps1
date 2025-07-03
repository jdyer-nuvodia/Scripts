# =============================================================================
# Script: Get-SystemHealthReport.ps1
# Created: 2025-04-02 20:23:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-07-03 00:09:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.4.2
# Additional Info: Fixed PSScriptAnalyzer compliance by properly passing parameters to functions instead of using script scope
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
.PARAMETER VM
    Switch parameter to indicate the system is a virtual machine. When specified, skips or adjusts certain checks like BitLocker and reduces thresholds for VM-specific scenarios.
.EXAMPLE
    .\Get-SystemHealthReport.ps1
    Runs health check with default parameters
.EXAMPLE
    .\Get-SystemHealthReport.ps1 -DaysToAnalyze 14 -ReportPath "C:\Reports"
    Runs health check analyzing 14 days of logs and saves report to specified path
.EXAMPLE
    .\Get-SystemHealthReport.ps1 -VM
    Runs health check optimized for virtual machine environments
#>

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateScript({
            if (-not (Test-Path $_)) {
                New-Item -ItemType Directory -Path $_ -Force | Out-Null
            }
            return $true
        })]
    [string]$ReportPath = $PSScriptRoot,

    [Parameter()]
    [ValidateRange(1, 30)]
    [int]$DaysToAnalyze = 7,

    [Parameter()]
    [string[]]$CriticalServices = @(
        # Windows Update
        'wuauserv',
        # Windows Defender
        'WinDefend',
        # Security Center
        'wscsvc',
        # Task Scheduler
        'Schedule',
        # Windows Event Log
        'EventLog',
        # Windows Firewall
        'mpssvc',
        # Server
        'LanmanServer',
        # DNS Client
        'Dnscache',
        # Datto RMM Agent
        'Datto RMM'
    ),

    [Parameter()]
    [switch]$VM
)

# Initialize error handling
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# Initialize logging
$SystemName = $env:COMPUTERNAME
$Timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$LogFile = Join-Path $ReportPath "SystemHealth_${SystemName}_${Timestamp}.log"

function Write-LogMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter()]
        [ValidateSet("Info", "Process", "Success", "Warning", "Error", "Debug")]
        [string]$Level = "Info"
    )

    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC"
    $LogMessage = "[$TimeStamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $LogMessage

    switch ($Level) {
        "Info" { Write-Information $LogMessage -InformationAction Continue }
        "Process" { Write-Information $LogMessage -InformationAction Continue }
        "Success" { Write-Information $LogMessage -InformationAction Continue }
        "Warning" { Write-Warning $LogMessage }
        "Error" { Write-Error $LogMessage -ErrorAction Continue }
        "Debug" { Write-Debug $LogMessage }
        default { Write-Information $LogMessage -InformationAction Continue }
    }
}

function Get-SystemResourceStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [bool]$IsVM
    )

    Write-LogMessage "Checking system resources..." -Level "Process"

    try {
        # CPU Usage - Adjusted thresholds for VMs
        $CpuUsage = (Get-Counter '\Processor(_Total)\% Processor Time' -ErrorAction Stop).CounterSamples.CookedValue
        if ($IsVM) {
            # More lenient thresholds for VMs
            $CpuStatus = if ($CpuUsage -ge 95) { "Error" } elseif ($CpuUsage -ge 90) { "Warning" } else { "Success" }
        } else {
            $CpuStatus = if ($CpuUsage -ge 90) { "Error" } elseif ($CpuUsage -ge 80) { "Warning" } else { "Success" }
        }
        Write-LogMessage "CPU Usage: $([math]::Round($CpuUsage, 2))%" -Level $CpuStatus

        # Memory Usage - Adjusted thresholds for VMs
        $OS = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        if ($OS.TotalVisibleMemorySize -gt 0) {
            $MemoryUsage = [math]::Round(($OS.TotalVisibleMemorySize - $OS.FreePhysicalMemory) / $OS.TotalVisibleMemorySize * 100, 2)
            if ($IsVM) {
                # More lenient thresholds for VMs
                $MemoryStatus = if ($MemoryUsage -ge 95) { "Error" } elseif ($MemoryUsage -ge 85) { "Warning" } else { "Success" }
            } else {
                $MemoryStatus = if ($MemoryUsage -ge 90) { "Error" } elseif ($MemoryUsage -ge 80) { "Warning" } else { "Success" }
            }
            Write-LogMessage "Memory Usage: ${MemoryUsage}%" -Level $MemoryStatus
        } else {
            Write-LogMessage "Invalid memory size reported by system" -Level "Error"
        }

        # Disk Space
        $Disks = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop
        foreach ($Disk in $Disks) {
            if ($Disk.Size -gt 0) {
                $FreeSpace = [math]::Round(($Disk.FreeSpace / $Disk.Size) * 100, 2)
                $Status = if ($FreeSpace -le 10) { "Error" } elseif ($FreeSpace -le 20) { "Warning" } else { "Success" }
                Write-LogMessage "Drive $($Disk.DeviceID) - Free Space: ${FreeSpace}% ($(([math]::Round($Disk.FreeSpace / 1GB, 2)))GB free)" -Level $Status
            } else {
                Write-LogMessage "Invalid disk size reported for drive $($Disk.DeviceID)" -Level "Error"
            }
        }
    } catch {
        Write-LogMessage "Error checking system resources: $($_.Exception.Message)" -Level "Error"
        Write-LogMessage "Stack Trace: $($_.ScriptStackTrace)" -Level "Debug"
    }
}

function Get-CriticalServicesStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Services
    )

    Write-LogMessage "Checking critical services..." -Level "Process"

    foreach ($Service in $Services) {
        try {
            $ServiceStatus = Get-Service -Name $Service -ErrorAction Stop
            $Status = switch ($ServiceStatus.Status) {
                'Running' { "Success" }
                'Stopped' { "Error" }
                default { "Warning" }
            }
            Write-LogMessage "Service $($ServiceStatus.DisplayName): $($ServiceStatus.Status)" -Level $Status
        } catch {
            Write-LogMessage "Error checking service ${Service}: $($_.Exception.Message)" -Level "Error"
        }
    }
}

function Get-EventLogAnalysis {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$Days
    )

    Write-LogMessage "Analyzing event logs..." -Level "Process"

    $StartTime = (Get-Date).AddDays(-$Days)
    $CriticalEvents = @(
        @{ Log = 'System'; Level = 2; Name = 'System Errors' },
        @{ Log = 'Application'; Level = 2; Name = 'Application Errors' },
        @{ Log = 'Security'; Level = 2; Name = 'Security Errors' }
    )

    foreach ($EventType in $CriticalEvents) {
        try {
            $Events = @(Get-WinEvent -FilterHashtable @{
                    LogName   = $EventType.Log
                    Level     = $EventType.Level
                    StartTime = $StartTime
                } -ErrorAction SilentlyContinue)

            $Count = $Events.Count
            $Status = if ($Count -ge 50) { "Error" } elseif ($Count -ge 20) { "Warning" } else { "Success" }
            Write-LogMessage "$($EventType.Name) in last $Days days: $Count" -Level $Status

            if ($Count -gt 0) {
                $TopErrors = $Events | Group-Object -Property Id |
                Sort-Object -Property Count -Descending |
                Select-Object -First 3
                foreach ($ErrorItem in $TopErrors) {
                    $Sample = $Events | Where-Object Id -eq $ErrorItem.Name | Select-Object -First 1
                    Write-LogMessage "  Top Error (ID $($ErrorItem.Name)): $($Sample.Message.Split([Environment]::NewLine)[0]) - Count: $($ErrorItem.Count)" -Level "Info"
                }
            }
        } catch {
            if ($_.Exception.Message -notlike "*No events were found*") {
                Write-LogMessage "Error checking $($EventType.Name): $($_.Exception.Message)" -Level "Error"
            }
        }
    }
}

function Get-WindowsUpdateStatus {
    [CmdletBinding()]
    param()

    Write-LogMessage "Checking Windows Update status..." -Level "Process"

    try {
        $UpdateSession = New-Object -ComObject Microsoft.Update.Session
        $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
        $SearchResult = $UpdateSearcher.Search("IsInstalled=0")

        $Count = $SearchResult.Updates.Count
        $Status = if ($Count -ge 10) { "Error" } elseif ($Count -ge 5) { "Warning" } else { "Success" }
        Write-LogMessage "Pending Windows Updates: $Count" -Level $Status

        if ($Count -gt 0) {
            $CriticalUpdates = @($SearchResult.Updates | Where-Object { $_.MsrcSeverity -eq "Critical" })
            if ($CriticalUpdates.Count -gt 0) {
                Write-LogMessage "Critical updates pending: $($CriticalUpdates.Count)" -Level "Warning"
                $CriticalUpdates | ForEach-Object {
                    Write-LogMessage "  - $($_.Title)" -Level "Info"
                }
            }
        }
    } catch {
        Write-LogMessage "Error checking Windows Updates: $($_.Exception.Message)" -Level "Error"
    }
}

function Get-NetworkStatus {
    [CmdletBinding()]
    param()

    Write-LogMessage "Checking network connectivity..." -Level "Process"

    $Targets = @(
        @{ Host = "8.8.8.8"; Name = "Google DNS" },
        @{ Host = "1.1.1.1"; Name = "Cloudflare DNS" },
        @{ Host = "www.microsoft.com"; Name = "Microsoft" }
    )

    foreach ($Target in $Targets) {
        try {
            $Result = Test-Connection -TargetName $Target.Host -Count 1 -ErrorAction Stop
            $LatencyMs = $Result.Latency
            $Status = switch ($LatencyMs) {
                { $_ -ge 200 } { "Warning" }
                { $_ -ge 500 } { "Error" }
                default { "Success" }
            }
            Write-LogMessage "Network latency to $($Target.Name) ($($Target.Host)): ${LatencyMs}ms" -Level $Status
        } catch {
            Write-LogMessage "Failed to reach $($Target.Name) ($($Target.Host)): $($_.Exception.Message)" -Level "Error"
        }
    }
}

function Get-SecurityStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [bool]$IsVM
    )

    Write-LogMessage "Checking security features..." -Level "Process"

    try {
        # Windows Defender Status
        $DefenderStatus = Get-MpComputerStatus -ErrorAction Stop
        $Status = if ($DefenderStatus.AntivirusEnabled) { "Success" } else { "Error" }
        Write-LogMessage "Windows Defender Status: $($DefenderStatus.AntivirusEnabled)" -Level $Status

        if ($DefenderStatus.AntivirusEnabled) {
            Write-LogMessage "  Last Scan: $($DefenderStatus.LastFullScanTime)" -Level "Info"
            Write-LogMessage "  Definitions: $($DefenderStatus.AntivirusSignatureLastUpdated)" -Level "Info"

            if ($DefenderStatus.LastFullScanTime -lt (Get-Date).AddDays(-7)) {
                Write-LogMessage "  Warning: Last full scan was more than 7 days ago" -Level "Warning"
            }
            if ($DefenderStatus.AntivirusSignatureLastUpdated -lt (Get-Date).AddDays(-3)) {
                Write-LogMessage "  Warning: Virus definitions are more than 3 days old" -Level "Warning"
            }
        }

        # DattoEDR Status
        $DattoEDRService = Get-Service -Name "Datto EDR Agent" -ErrorAction SilentlyContinue
        if ($DattoEDRService) {
            $Status = if ($DattoEDRService.Status -eq 'Running') { "Success" } else { "Error" }
            Write-LogMessage "DattoEDR Status: $($DattoEDRService.Status)" -Level $Status
        } else {
            Write-LogMessage "DattoEDR not installed" -Level "Error"
        }

        # Firewall Status
        $FirewallProfiles = Get-NetFirewallProfile -ErrorAction Stop
        foreach ($FwProfile in $FirewallProfiles) {
            $Status = if ($FwProfile.Enabled) { "Success" } else { "Error" }
            Write-LogMessage "Firewall Profile $($FwProfile.Name): $($FwProfile.Enabled)" -Level $Status
        }

        # BitLocker Status - Skip for VMs
        if (-not $IsVM) {
            $BitLockerVolumes = Get-BitLockerVolume -ErrorAction SilentlyContinue
            if ($BitLockerVolumes) {
                foreach ($Volume in $BitLockerVolumes) {
                    $Status = switch ($Volume.ProtectionStatus) {
                        'On' { "Success" }
                        'Off' { "Error" }
                        default { "Warning" }
                    }
                    Write-LogMessage "BitLocker on $($Volume.MountPoint): $($Volume.ProtectionStatus)" -Level $Status
                    if ($Volume.ProtectionStatus -eq 'On') {
                        Write-LogMessage "  Encryption Method: $($Volume.EncryptionMethod)" -Level "Info"
                    }
                }
            } else {
                Write-LogMessage "BitLocker not configured on any volumes" -Level "Warning"
            }
        } else {
            Write-LogMessage "BitLocker check skipped (VM environment)" -Level "Info"
        }
    } catch {
        Write-LogMessage "Error checking security status: $($_.Exception.Message)" -Level "Error"
        Write-LogMessage "Stack Trace: $($_.ScriptStackTrace)" -Level "Debug"
    }
}

# Main execution
try {
    Write-LogMessage "=== System Health Check Started ===" -Level "Process"
    Write-LogMessage "System: $SystemName" -Level "Info"
    Write-LogMessage "Timestamp: $((Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss UTC'))" -Level "Info"
    if ($VM) {
        Write-LogMessage "VM Mode: Enabled - Adjusted thresholds and skipped VM-incompatible checks" -Level "Info"
    }

    Get-SystemResourceStatus -IsVM $VM
    Get-CriticalServicesStatus -Services $CriticalServices
    Get-EventLogAnalysis -Days $DaysToAnalyze
    Get-WindowsUpdateStatus
    Get-NetworkStatus
    Get-SecurityStatus -IsVM $VM

    Write-LogMessage "=== System Health Check Completed ===" -Level "Process"
    Write-LogMessage "Log file saved to: $LogFile" -Level "Success"
} catch {
    Write-LogMessage "Script execution failed: $($_.Exception.Message)" -Level "Error"
    Write-LogMessage "Stack Trace: $($_.ScriptStackTrace)" -Level "Error"
    throw
}
