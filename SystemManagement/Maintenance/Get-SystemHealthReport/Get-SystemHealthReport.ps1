# =============================================================================
# Script: Get-SystemHealthReport.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.5.2
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
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
            if (-not [string]::IsNullOrWhiteSpace($_) -and -not (Test-Path -Path $_)) {
                New-Item -ItemType Directory -Path $_ -Force | Out-Null
            }
            return $true
        })]
    [string]$ReportPath,

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

# Set default ReportPath if not provided
if ([string]::IsNullOrWhiteSpace($ReportPath)) {
    $ReportPath = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }
}

# Color system for both PowerShell 5.1 and 7+
if ($PSVersionTable.PSVersion.Major -ge 7) {
    $Script:Colors = @{
        Reset = "`e[0m"
        White = "`e[37m"
        Cyan = "`e[36m"
        Green = "`e[32m"
        Yellow = "`e[33m"
        Red = "`e[31m"
        Magenta = "`e[35m"
        DarkGray = "`e[90m"
        Bold = "`e[1m"
    }
    $Script:UseAnsiColors = $true
} else {
    # PowerShell 5.1 - Use console color mapping
    $Script:Colors = @{
        Reset = ""
        White = "White"
        Cyan = "Cyan"
        Green = "Green"
        Yellow = "Yellow"
        Red = "Red"
        Magenta = "Magenta"
        DarkGray = "DarkGray"
        Bold = ""
    }
    $Script:UseAnsiColors = $false
}

# Initialize logging
$SystemName = $env:COMPUTERNAME
$Timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$LogFile = Join-Path $ReportPath "SystemHealth_${ SystemName}_${ Timestamp}.log"

function Write-ColorOutput {
    <#
    .SYNOPSIS
    Writes colored output that works in both PowerShell 5.1 and 7+ without using Write-Host.
    .DESCRIPTION
    Uses ANSI escape codes for PowerShell 7+ and console color changes for PowerShell 5.1.
    This function complies with copilot-instructions.md by using Write-Output instead of Write-Host.
    .PARAMETER Message
    The message to write with color.
    .PARAMETER Color
    The color to use (White, Cyan, Green, Yellow, Red, Magenta, DarkGray).
    .EXAMPLE
    Write-ColorOutput -Message "Success!" -Color "Green"
    Writes "Success!" in green color.
    #>
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
        Write-Output "${colorCode}${ Message}${resetCode}"
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
    Add-Content -Path $script:LogFile -Value $LogMessage

    # Color mapping for different message levels
    $colorMap = @{
        "Info" = "White"
        "Process" = "Cyan"
        "Success" = "Green"
        "Warning" = "Yellow"
        "Error" = "Red"
        "Debug" = "Magenta"
    }

    $color = if ($colorMap.ContainsKey($Level)) { $colorMap[$Level] } else { "White" }
    Write-ColorOutput -Message $LogMessage -Color $color
}

function Write-SystemError {
    <#
    .SYNOPSIS
    Writes system errors in a clearly formatted way that distinguishes them from script errors.
    .PARAMETER EventID
    The Windows Event ID.
    .PARAMETER Message
    The error message.
    .PARAMETER Count
    The number of occurrences.
    .PARAMETER Type
    The type of error (System, Application, Security).
    #>
    param(
        [Parameter(Mandatory = $true)]
        [int]$EventID,

        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $true)]
        [int]$Count,

        [Parameter(Mandatory = $true)]
        [string]$Type
    )

    $errorPrefix = "SYSTEM ERROR DETECTED"
    $formattedMessage = "[$errorPrefix] Event ID $EventID ($Type): $Message - Occurrences: $Count"
    Write-ColorOutput -Message $formattedMessage -Color "Red"

    # Also log to file
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC"
    $LogMessage = "[$TimeStamp] [SYSTEM_ERROR] Event ID $EventID ($Type): $Message - Count: $Count"
    Add-Content -Path $script:LogFile -Value $LogMessage
}

function Get-VMPlatform {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()

    try {
        $ComputerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        $BIOS = Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop

        # Check if this is a virtual machine
        $IsVirtual = $false
        $Platform = "Physical"

        # Check model first
        if ($ComputerSystem.Model -match "VMware|VirtualBox|Virtual Machine|Xen|KVM") {
            $IsVirtual = $true
            $Platform = switch -Regex ($ComputerSystem.Model) {
                "VMware" { "VMware" }
                "VirtualBox" { "VirtualBox" }
                "Virtual Machine" {
                    if ($ComputerSystem.Manufacturer -match "Microsoft") { "Hyper-V" }
                    else { "Generic" }
                }
                "Xen" { "Xen" }
                "KVM" { "KVM" }
                default { "Generic" }
            }
        }

        # Check BIOS/manufacturer for additional VM indicators
        if (-not $IsVirtual) {
            if ($BIOS.SerialNumber -match "VMware|VirtualBox|Xen" -or
                $ComputerSystem.Manufacturer -match "VMware|Oracle Corporation|Microsoft Corporation" -and
                $ComputerSystem.Model -match "Virtual") {
                $IsVirtual = $true
                $Platform = "Detected"
            }
        }

        return @{
            IsVirtual = $IsVirtual
            Platform = $Platform
            Manufacturer = $ComputerSystem.Manufacturer
            Model = $ComputerSystem.Model
        }
    } catch {
        return @{
            IsVirtual = $false
            Platform = "Physical"
            Manufacturer = $null
            Model = $null
        }
    }
}

function Get-VMOptimizedService {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$VMPlatform,

        [Parameter(Mandatory = $true)]
        [string[]]$BaseServices
    )

    $VMServices = $BaseServices.Clone()

    switch ($VMPlatform) {
        "VMware" {
            $VMServices += @('VMTools', 'VGAuthService')
        }
        "Hyper-V" {
            $VMServices += @('vmicheartbeat', 'vmickvpexchange', 'vmictimesync')
        }
        "VirtualBox" {
            $VMServices += @('VBoxService')
        }
    }

    # Remove services that are less critical in VMs
    $VMServices = $VMServices | Where-Object { $_ -notin @('wuauserv') }

    return $VMServices
}

function Get-VMConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$VMPlatform
    )

    Write-LogMessage "Gathering VM configuration details..." -Level "Process"

    try {
        $ComputerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        $Processor = Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop

        Write-LogMessage "VM Platform: $VMPlatform" -Level "Info"
        Write-LogMessage "Virtual CPUs: $($ComputerSystem.NumberOfProcessors)" -Level "Info"
        Write-LogMessage "Virtual CPU Cores: $($Processor.NumberOfCores)" -Level "Info"
        Write-LogMessage "Virtual Memory: $([math]::Round($ComputerSystem.TotalPhysicalMemory / 1GB, 2))GB" -Level "Info"

        # Check for VM-specific features
        if ($ComputerSystem.HypervisorPresent) {
            Write-LogMessage "Nested virtualization detected" -Level "Info"
        }

    } catch {
        Write-LogMessage "Error gathering VM configuration: $($_.Exception.Message)" -Level "Error"
    }
}

function Get-VMPerformanceMetric {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$VMPlatform
    )

    Write-LogMessage "Checking VM-specific performance metrics..." -Level "Process"

    try {
        # Check for VM memory ballooning
        $MemoryCounters = Get-Counter -Counter "\Memory\*" -ErrorAction SilentlyContinue
        if ($MemoryCounters) {
            $BalloonedMemory = $MemoryCounters.CounterSamples |
                Where-Object { $_.Path -like "*Balloon*" }
            if ($BalloonedMemory) {
                Write-LogMessage "Memory ballooning detected: $($BalloonedMemory.CookedValue)" -Level "Warning"
            }
        }

        # Check for VM-specific disk performance
        if ($VMPlatform -eq "VMware") {
            $VMwareTools = Get-Service -Name "VMTools" -ErrorAction SilentlyContinue
            if ($VMwareTools -and $VMwareTools.Status -eq 'Running') {
                Write-LogMessage "VMware Tools running - optimized disk performance available" -Level "Success"
            }
        }

        # Check for CPU steal time (if available)
        $ProcessorCounters = Get-Counter -Counter "\Processor Information(_Total)\% Processor Time" -ErrorAction SilentlyContinue
        if ($ProcessorCounters) {
            Write-LogMessage "CPU performance counters available for monitoring" -Level "Info"
        }

    } catch {
        Write-LogMessage "Error checking VM performance metrics: $($_.Exception.Message)" -Level "Error"
    }
}

function Get-VMStorageStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$VMPlatform,

        [Parameter(Mandatory = $true)]
        [bool]$IsVM
    )

    if (-not $IsVM) { return }

    Write-LogMessage "Checking VM storage configuration..." -Level "Process"

    try {
        # Check for thin provisioning
        $Disks = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType = 3" -ErrorAction Stop
        foreach ($Disk in $Disks) {
            # VM-specific disk checks            $PhysicalDisk = Get-CimInstance -ClassName Win32_DiskDrive |
            Where-Object { $_.DeviceID -eq $Disk.DeviceID }

            if ($PhysicalDisk -and $PhysicalDisk.Model -match "VMware|Virtual") {
                Write-LogMessage "Virtual disk detected: $($Disk.DeviceID) - $($PhysicalDisk.Model)" -Level "Info"

                # Check for potential thin provisioning issues
                if ($Disk.Size -gt 100GB -and ($Disk.FreeSpace / $Disk.Size) -lt 0.15) {
                    Write-LogMessage "Large virtual disk with low free space - monitor for thin provisioning issues" -Level "Warning"
                }
            }
        }

        # Platform-specific storage checks
        switch ($VMPlatform) {
            "VMware" {
                # Check for VMware snapshot warnings in event logs
                $SnapshotEvents = Get-WinEvent -FilterHashtable @{
                    LogName = 'System'
                    ProviderName = 'VMware*'
                    Level = 3
                    StartTime = (Get-Date).AddDays(-7)
                } -ErrorAction SilentlyContinue

                if ($SnapshotEvents) {
                    Write-LogMessage "VMware snapshot-related warnings found in last 7 days: $($SnapshotEvents.Count)" -Level "Warning"
                }
            }
        }
    } catch {
        Write-LogMessage "Error checking VM storage status: $($_.Exception.Message)" -Level "Error"
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
            Write-LogMessage "Memory Usage: ${ MemoryUsage}%" -Level $MemoryStatus
        } else {
            Write-LogMessage "Invalid memory size reported by system" -Level "Error"
        }

        # Disk Space
        $Disks = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType = 3" -ErrorAction Stop
        foreach ($Disk in $Disks) {
            if ($Disk.Size -gt 0) {
                $FreeSpace = [math]::Round(($Disk.FreeSpace / $Disk.Size) * 100, 2)
                $Status = if ($FreeSpace -le 10) { "Error" } elseif ($FreeSpace -le 20) { "Warning" } else { "Success" }
                Write-LogMessage "Drive $($Disk.DeviceID) - Free Space: ${ FreeSpace}% ($(([math]::Round($Disk.FreeSpace / 1GB, 2)))GB free)" -Level $Status
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
            Write-LogMessage "Error checking service ${ Service}: $($_.Exception.Message)" -Level "Error"
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
                    LogName = $EventType.Log
                    Level = $EventType.Level
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
                    $Sample = $Events | Where-Object { $_.Id -eq $ErrorItem.Name } | Select-Object -First 1
                    Write-SystemError -EventID $ErrorItem.Name -Message $Sample.Message.Split([Environment]::NewLine)[0] -Count $ErrorItem.Count -Type $EventType.Name
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
        $SearchResult = $UpdateSearcher.Search("IsInstalled = 0")

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
            $Result = Test-Connection -ComputerName $Target.Host -Count 1 -ErrorAction Stop
            # Handle different PowerShell versions - newer versions use Latency property
            $LatencyMs = if ($Result.Latency) { $Result.Latency } else { $Result.ResponseTime }
            $Status = switch ($LatencyMs) {
                { $_ -ge 200 } { "Warning" }
                { $_ -ge 500 } { "Error" }
                default { "Success" }
            }
            Write-LogMessage "Network latency to $($Target.Name) ($($Target.Host)): ${ LatencyMs}ms" -Level $Status
        } catch {
            Write-LogMessage "Failed to reach $($Target.Name) ($($Target.Host)): $($_.Exception.Message)" -Level "Error"
        }
    }
}

function Get-VMNetworkStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$VMPlatform
    )

    Write-LogMessage "Checking VM network connectivity for $VMPlatform..." -Level "Process"

    # More lenient latency thresholds for VMs
    $Targets = @(
        @{ Host = "8.8.8.8"; Name = "Google DNS"; VMThreshold = 300; PhysicalThreshold = 200 },
        @{ Host = "1.1.1.1"; Name = "Cloudflare DNS"; VMThreshold = 300; PhysicalThreshold = 200 },
        @{ Host = "www.microsoft.com"; Name = "Microsoft"; VMThreshold = 500; PhysicalThreshold = 300 }
    )

    foreach ($Target in $Targets) {
        try {
            $Result = Test-Connection -ComputerName $Target.Host -Count 3 -ErrorAction Stop
            # Handle different PowerShell versions - newer versions use Latency property
            $AvgLatency = if ($Result[0].Latency) {
                ($Result | Measure-Object -Property Latency -Average).Average
            } else {
                ($Result | Measure-Object -Property ResponseTime -Average).Average
            }

            $Threshold = $Target.VMThreshold
            $Status = switch ($AvgLatency) {
                { $_ -ge ($Threshold * 2) } { "Error" }
                { $_ -ge $Threshold } { "Warning" }
                default { "Success" }
            }
            Write-LogMessage "Network latency to $($Target.Name): $([math]::Round($AvgLatency, 2))ms (VM threshold: ${ Threshold}ms)" -Level $Status
        } catch {
            Write-LogMessage "Failed to reach $($Target.Name): $($_.Exception.Message)" -Level "Error"
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
        try {
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
        } catch {
            Write-LogMessage "Windows Defender module not available or accessible: $($_.Exception.Message)" -Level "Warning"

            # Alternative check using service status
            $DefenderService = Get-Service -Name "WinDefend" -ErrorAction SilentlyContinue
            if ($DefenderService) {
                $Status = if ($DefenderService.Status -eq 'Running') { "Success" } else { "Error" }
                Write-LogMessage "Windows Defender Service Status: $($DefenderService.Status)" -Level $Status
            } else {
                Write-LogMessage "Windows Defender service not found" -Level "Error"
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
    Write-ColorOutput -Message "System Health Check v1.5.1" -Color "Green"
    Write-ColorOutput -Message "==============================" -Color "Green"
    Write-LogMessage "=== System Health Check Started ===" -Level "Process"
    Write-LogMessage "System: $SystemName" -Level "Info"
    Write-LogMessage "Timestamp: $((Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss UTC'))" -Level "Info"

    # Enhanced VM detection and configuration
    $VMInfo = Get-VMPlatform
    $IsVirtualMachine = $VMInfo.IsVirtual -or $VM.IsPresent
    $VMPlatform = $VMInfo.Platform

    if ($IsVirtualMachine) {
        Write-LogMessage "Virtual machine detected: $VMPlatform ($($VMInfo.Manufacturer) $($VMInfo.Model))" -Level "Info"
        Write-LogMessage "VM Mode: Enabled - Platform: $VMPlatform" -Level "Info"
        Get-VMConfiguration -VMPlatform $VMPlatform
        $OptimizedServices = Get-VMOptimizedService -VMPlatform $VMPlatform -BaseServices $CriticalServices
    } else {
        Write-LogMessage "Physical machine detected" -Level "Info"
        $OptimizedServices = $CriticalServices
    }

    Get-SystemResourceStatus -IsVM $IsVirtualMachine
    Get-CriticalServicesStatus -Services $OptimizedServices
    Get-EventLogAnalysis -Days $DaysToAnalyze
    Get-WindowsUpdateStatus

    if ($IsVirtualMachine) {
        Get-VMPerformanceMetric -VMPlatform $VMPlatform
        Get-VMNetworkStatus -VMPlatform $VMPlatform
        Get-VMStorageStatus -VMPlatform $VMPlatform -IsVM $IsVirtualMachine
    } else {
        Get-NetworkStatus
    }

    Get-SecurityStatus -IsVM $IsVirtualMachine

    Write-LogMessage "=== System Health Check Completed ===" -Level "Process"
    Write-LogMessage "Log file saved to: $LogFile" -Level "Success"
} catch {
    Write-LogMessage "Script execution failed: $($_.Exception.Message)" -Level "Error"
    Write-LogMessage "Stack Trace: $($_.ScriptStackTrace)" -Level "Error"
    throw
}
