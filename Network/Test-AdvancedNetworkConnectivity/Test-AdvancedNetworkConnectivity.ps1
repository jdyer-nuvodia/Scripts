# =============================================================================
# Script: Test-AdvancedNetworkConnectivity.ps1
# Created: 2025-06-23 21:45:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-06-27 15:09:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.4.0
# Additional Info: Fixed parallel processing serialization issue by replacing custom class with hashtable-based approach
# =============================================================================

<#
.SYNOPSIS
    Advanced network connectivity testing tool with multi-target support and comprehensive diagnostics.

.DESCRIPTION
    This script provides enhanced network connectivity testing capabilities including:
    - Multi-target testing with parallel processing support
    - Advanced diagnostics beyond basic ping (DNS resolution, port connectivity, MTU discovery)
    - Comprehensive logging and reporting
    - Intelligent defaults for quick testing without configuration
    - Support for both continuous and count-based testing

    The script uses intelligent defaults including common DNS servers and connectivity test targets
    to enable immediate testing without requiring target specification.

.PARAMETER Target
    Array of target hosts to test. Can be IP addresses, hostnames, or FQDNs.
    Default: @("8.8.8.8", "1.1.1.1", "microsoft.com")

.PARAMETER TargetFile
    Path to CSV file containing targets to test. CSV format: Target,Description,Priority

.PARAMETER Count
    Number of pings to send to each target. Use 0 for continuous testing.
    Default: 10

.PARAMETER TestType
    Types of network tests to perform. Options: Ping, DNS, Port, MTU, All
    Default: @("All")

.PARAMETER Ports
    Array of ports to test when using Port test type.
    Default: @(80, 443, 53)

.PARAMETER OutputPath
    Directory path where log files will be saved.
    Default: Same directory as script

.PARAMETER Parallel
    Enable parallel processing for multiple targets.
    Default: $true

.PARAMETER MaxMTU
    Maximum MTU size to test when using MTU discovery.
    Default: 1500

.PARAMETER Timeout
    Timeout in milliseconds for network operations.
    Default: 5000

.PARAMETER WhatIf
    Shows what would be performed without executing the operations.

.EXAMPLE
    .\Test-AdvancedNetworkConnectivity.ps1
    Performs comprehensive testing (All test types) on default targets (8.8.8.8, 1.1.1.1, microsoft.com)

.EXAMPLE
    .\Test-AdvancedNetworkConnectivity.ps1 -Target @("google.com", "cloudflare.com") -Count 50 -TestType All
    Performs comprehensive testing on specified targets with 50 iterations each

.EXAMPLE
    .\Test-AdvancedNetworkConnectivity.ps1 -TargetFile "C:\targets.csv" -Parallel -TestType @("Ping", "Port") -Ports @(80, 443, 22)
    Tests targets from CSV file with parallel processing, testing ping and specific ports

.EXAMPLE
    .\Test-AdvancedNetworkConnectivity.ps1 -Target "server01.domain.com" -TestType MTU -MaxMTU 9000
    Performs MTU discovery on specified target up to 9000 bytes

.NOTES
    Validation Requirements: Verify network connectivity, file system access, DNS resolution capabilities
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string[]]$Target = @("8.8.8.8", "1.1.1.1", "microsoft.com"),

    [string]$TargetFile,

    [int]$Count = 10,

    [ValidateSet("Ping", "DNS", "Port", "MTU", "All")]
    [string[]]$TestType = @("All"),

    [int[]]$Ports = @(80, 443, 53),

    [string]$OutputPath = $PSScriptRoot,

    [bool]$Parallel = $true,

    [int]$MaxMTU = 1500,

    [int]$Timeout = 5000
)

# Initialize script variables
$script:logFile = $null
$script:results = @{}
$script:interrupted = $false

# Target test results structure
function Initialize-NetworkTestResult {
    param([string]$Target)

    return @{
        Target = $Target
        Description = ""
        Priority = "Medium"
        PingResults = @{}
        DNSResults = @{}
        PortResults = @{}
        MTUResults = @{}
        TestStartTime = Get-Date
        TestEndTime = $null
        Status = "Running"
        Errors = @()
        LogBuffer = @()
    }
}

function Write-LogMessage {
    param(
        [string]$Message,
        [string]$FilePath,
        [switch]$NoConsole,
        [ref]$LogBuffer
    )

    $timestampedMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC'): $Message"

    if ($LogBuffer) {
        $LogBuffer.Value += $timestampedMessage
    }
    elseif ($FilePath) {
        Add-Content -Path $FilePath -Value $timestampedMessage -ErrorAction SilentlyContinue
    }

    if (-not $NoConsole) {
        Write-Information -MessageData $timestampedMessage -InformationAction Continue
    }
}

function Write-TargetLogSection {
    param(
        [hashtable]$TestResult,
        [string]$FilePath
    )

    # Write target header
    $headerMessages = @(
        "========================================",
        "TARGET: $($TestResult.Target)",
        "Description: $($TestResult.Description)",
        "Priority: $($TestResult.Priority)",
        "Test Started: $($TestResult.TestStartTime.ToString('yyyy-MM-dd HH:mm:ss UTC'))",
        "Test Completed: $($TestResult.TestEndTime.ToString('yyyy-MM-dd HH:mm:ss UTC'))",
        "Status: $($TestResult.Status)",
        "========================================"
    )

    foreach ($message in $headerMessages) {
        $timestampedMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC'): $message"
        Add-Content -Path $FilePath -Value $timestampedMessage -ErrorAction SilentlyContinue
    }

    # Write all buffered log messages for this target
    foreach ($logEntry in $TestResult.LogBuffer) {
        Add-Content -Path $FilePath -Value $logEntry -ErrorAction SilentlyContinue
    }

    # Add section separator
    $separatorMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC'): Tests completed for target: $($TestResult.Target)"
    Add-Content -Path $FilePath -Value $separatorMessage -ErrorAction SilentlyContinue
    Add-Content -Path $FilePath -Value "" -ErrorAction SilentlyContinue
}

function Get-FormattedSize {
    param([int64]$Size)

    if ($Size -gt 1GB) { return "{0:N2} GB" -f ($Size / 1GB) }
    if ($Size -gt 1MB) { return "{0:N2} MB" -f ($Size / 1MB) }
    if ($Size -gt 1KB) { return "{0:N2} KB" -f ($Size / 1KB) }
    return "$Size Bytes"
}

function Import-TargetsFromFile {
    param([string]$FilePath)

    if (-not (Test-Path -Path $FilePath)) {
        Write-Error -Message "Target file not found: $FilePath"
        return @()
    }

    try {
        $targets = Import-Csv -Path $FilePath
        $targetList = @()

        foreach ($target in $targets) {
            $targetObj = [PSCustomObject]@{
                Target = $target.Target
                Description = if ($target.Description) { $target.Description } else { "" }
                Priority = if ($target.Priority) { $target.Priority } else { "Medium" }
            }
            $targetList += $targetObj
        }

        return $targetList
    }
    catch {
        Write-Error -Message "Error reading target file: $_"
        return @()
    }
}

function Test-PingConnectivity {
    param(
        [string]$TargetHost,
        [int]$PingCount,
        [int]$TimeoutMs,
        [ref]$LogBuffer
    )

    $pingResults = @{
        Sent = 0
        Received = 0
        Lost = 0
        MinTime = [int]::MaxValue
        MaxTime = 0
        AvgTime = 0
        TotalTime = 0
        PacketLoss = 0
        Details = @()
    }

    Write-LogMessage -Message "Starting ping test for $TargetHost ($PingCount packets)" -LogBuffer $LogBuffer

    for ($i = 1; $i -le $PingCount; $i++) {
        try {
            $ping = Test-Connection -ComputerName $TargetHost -Count 1 -TimeoutSeconds ($TimeoutMs / 1000) -ErrorAction Stop

            $responseTime = $ping.ResponseTime
            $pingResults.Sent++
            $pingResults.Received++
            $pingResults.TotalTime += $responseTime

            if ($responseTime -lt $pingResults.MinTime) { $pingResults.MinTime = $responseTime }
            if ($responseTime -gt $pingResults.MaxTime) { $pingResults.MaxTime = $responseTime }

            $pingResults.Details += "Reply from $($ping.Address): time=$($responseTime)ms"

            Write-LogMessage -Message "Ping $i/$PingCount to ${TargetHost}: $($responseTime)ms" -LogBuffer $LogBuffer
        }
        catch {
            $pingResults.Sent++
            $pingResults.Lost++
            $pingResults.Details += "Request timeout for ping $i"
            Write-LogMessage -Message "Ping $i/$PingCount to ${TargetHost}: Request timeout" -LogBuffer $LogBuffer
        }

        if ($i -lt $PingCount) {
            Start-Sleep -Milliseconds 1000
        }
    }

    if ($pingResults.Received -gt 0) {
        $pingResults.AvgTime = [math]::Round($pingResults.TotalTime / $pingResults.Received, 2)
    }

    if ($pingResults.MinTime -eq [int]::MaxValue) {
        $pingResults.MinTime = 0
    }

    $pingResults.PacketLoss = [math]::Round(($pingResults.Lost / $pingResults.Sent) * 100, 2)

    return $pingResults
}

function Test-DNSResolution {
    param(
        [string]$TargetHost,
        [ref]$LogBuffer
    )

    $dnsResults = @{
        HostName = $TargetHost
        IPAddresses = @()
        ResolutionTime = 0
        Success = $false
        ErrorMessage = ""
    }

    Write-LogMessage -Message "Starting DNS resolution test for $TargetHost" -LogBuffer $LogBuffer

    try {
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $dnsResult = Resolve-DnsName -Name $TargetHost -ErrorAction Stop
        $stopwatch.Stop()

        $dnsResults.ResolutionTime = $stopwatch.ElapsedMilliseconds
        $dnsResults.Success = $true

        foreach ($record in $dnsResult) {
            if ($record.IPAddress) {
                $dnsResults.IPAddresses += $record.IPAddress
            }
        }

        Write-LogMessage -Message "DNS resolution for $TargetHost successful: $($dnsResults.IPAddresses -join ', ') ($($dnsResults.ResolutionTime)ms)" -LogBuffer $LogBuffer
    }
    catch {
        $dnsResults.ErrorMessage = $_.Exception.Message
        Write-LogMessage -Message "DNS resolution for $TargetHost failed: $($_.Exception.Message)" -LogBuffer $LogBuffer
    }

    return $dnsResults
}

function Test-PortConnectivity {
    param(
        [string]$TargetHost,
        [int[]]$PortList,
        [int]$TimeoutMs,
        [ref]$LogBuffer
    )

    $portResults = @{
        TestedPorts = @()
        OpenPorts = @()
        ClosedPorts = @()
        Results = @{}
    }

    Write-LogMessage -Message "Starting port connectivity test for $TargetHost on ports: $($PortList -join ', ')" -LogBuffer $LogBuffer

    foreach ($port in $PortList) {
        $portResults.TestedPorts += $port

        try {
            $tcpClient = New-Object System.Net.Sockets.TcpClient
            $connectTask = $tcpClient.ConnectAsync($TargetHost, $port)

            if ($connectTask.Wait($TimeoutMs)) {
                if ($tcpClient.Connected) {
                    $portResults.OpenPorts += $port
                    $portResults.Results[$port] = "Open"
                    Write-LogMessage -Message "Port $port on ${TargetHost}: Open" -LogBuffer $LogBuffer
                }
                else {
                    $portResults.ClosedPorts += $port
                    $portResults.Results[$port] = "Closed"
                    Write-LogMessage -Message "Port $port on ${TargetHost}: Closed" -LogBuffer $LogBuffer
                }
            }
            else {
                $portResults.ClosedPorts += $port
                $portResults.Results[$port] = "Timeout"
                Write-LogMessage -Message "Port $port on ${TargetHost}: Timeout" -LogBuffer $LogBuffer
            }

            $tcpClient.Close()
        }
        catch {
            $portResults.ClosedPorts += $port
            $portResults.Results[$port] = "Error: $($_.Exception.Message)"
            Write-LogMessage -Message "Port $port on ${TargetHost}: Error - $($_.Exception.Message)" -LogBuffer $LogBuffer
        }
    }

    return $portResults
}

function Test-MTUDiscovery {
    param(
        [string]$TargetHost,
        [int]$MaxMTUSize,
        [ref]$LogBuffer
    )

    $mtuResults = @{
        MaxMTU = 0
        OptimalMTU = 0
        TestResults = @()
        Success = $false
    }

    Write-LogMessage -Message "Starting MTU discovery for $TargetHost (max size: $MaxMTUSize)" -LogBuffer $LogBuffer

    # Start with common MTU sizes and work up
    $testSizes = @(576, 1024, 1280, 1460, 1500)
    if ($MaxMTUSize -gt 1500) {
        $testSizes += @(4000, 8000, $MaxMTUSize)
    }

    foreach ($size in $testSizes | Sort-Object) {
        if ($size -gt $MaxMTUSize) { continue }

        try {
            $pingSize = $size - 28
            # Subtract IP and ICMP headers
            if ($pingSize -lt 1) { continue }

            $ping = Test-Connection -ComputerName $TargetHost -BufferSize $pingSize -Count 1 -ErrorAction Stop

            if ($ping) {
                $mtuResults.MaxMTU = $size
                $mtuResults.TestResults += "MTU $size bytes: Success"
                Write-LogMessage -Message "MTU test for $TargetHost at $size bytes: Success" -LogBuffer $LogBuffer
            }
        }
        catch {
            $mtuResults.TestResults += "MTU $size bytes: Failed"
            Write-LogMessage -Message "MTU test for $TargetHost at $size bytes: Failed" -LogBuffer $LogBuffer
            break
        }
    }

    if ($mtuResults.MaxMTU -gt 0) {
        $mtuResults.OptimalMTU = $mtuResults.MaxMTU
        $mtuResults.Success = $true
    }

    return $mtuResults
}

function Test-SingleTarget {
    param(
        [string]$TargetHost,
        [string]$Description = "",
        [string]$Priority = "Medium",
        [string[]]$TestTypes = @("Ping", "DNS"),
        [int]$TestCount = 10,
        [int]$TestTimeout = 5000,
        [int[]]$TestPorts = @(80, 443, 53),
        [int]$TestMaxMTU = 1500
    )

    $testResult = Initialize-NetworkTestResult -Target $TargetHost
    $testResult.Description = $Description
    $testResult.Priority = $Priority

    try {
        # Ping Test
        if ($TestTypes -contains "Ping" -or $TestTypes -contains "All") {
            $testResult.PingResults = Test-PingConnectivity -TargetHost $TargetHost -PingCount $TestCount -TimeoutMs $TestTimeout -LogBuffer ([ref]$testResult.LogBuffer)
        }

        # DNS Test
        if ($TestTypes -contains "DNS" -or $TestTypes -contains "All") {
            $testResult.DNSResults = Test-DNSResolution -TargetHost $TargetHost -LogBuffer ([ref]$testResult.LogBuffer)
        }

        # Port Test
        if ($TestTypes -contains "Port" -or $TestTypes -contains "All") {
            $testResult.PortResults = Test-PortConnectivity -TargetHost $TargetHost -PortList $TestPorts -TimeoutMs $TestTimeout -LogBuffer ([ref]$testResult.LogBuffer)
        }

        # MTU Test
        if ($TestTypes -contains "MTU" -or $TestTypes -contains "All") {
            $testResult.MTUResults = Test-MTUDiscovery -TargetHost $TargetHost -MaxMTUSize $TestMaxMTU -LogBuffer ([ref]$testResult.LogBuffer)
        }

        $testResult.Status = "Completed"
        $testResult.TestEndTime = Get-Date
    }
    catch {
        $testResult.Status = "Failed"
        $testResult.Errors += $_.Exception.Message
        $testResult.TestEndTime = Get-Date
        Write-LogMessage -Message "Tests failed for target: $TargetHost - $($_.Exception.Message)" -LogBuffer ([ref]$testResult.LogBuffer)
    }

    return $testResult
}

function Write-TestSummary {
    param([hashtable]$AllResults)

    Write-LogMessage -Message "`n========================================" -FilePath $script:logFile
    Write-LogMessage -Message "TEST SUMMARY" -FilePath $script:logFile
    Write-LogMessage -Message "========================================" -FilePath $script:logFile

    foreach ($targetName in $AllResults.Keys) {
        $result = $AllResults[$targetName]

        Write-LogMessage -Message "`nTarget: $($result.Target)" -FilePath $script:logFile
        Write-LogMessage -Message "Status: $($result.Status)" -FilePath $script:logFile
        Write-LogMessage -Message "Test Duration: $((($result.TestEndTime - $result.TestStartTime).TotalSeconds).ToString('N2')) seconds" -FilePath $script:logFile

        # Ping Summary
        if ($result.PingResults.Count -gt 0) {
            $ping = $result.PingResults
            Write-LogMessage -Message "Ping Results: $($ping.Received)/$($ping.Sent) successful ($($ping.PacketLoss)% loss)" -FilePath $script:logFile
            if ($ping.Received -gt 0) {
                Write-LogMessage -Message "  Latency: Min=$($ping.MinTime)ms, Max=$($ping.MaxTime)ms, Avg=$($ping.AvgTime)ms" -FilePath $script:logFile
            }
        }

        # DNS Summary
        if ($result.DNSResults.Count -gt 0) {
            $dns = $result.DNSResults
            if ($dns.Success) {
                Write-LogMessage -Message "DNS Resolution: Success ($($dns.ResolutionTime)ms) - $($dns.IPAddresses -join ', ')" -FilePath $script:logFile
            }
            else {
                Write-LogMessage -Message "DNS Resolution: Failed - $($dns.ErrorMessage)" -FilePath $script:logFile
            }
        }

        # Port Summary
        if ($result.PortResults.Count -gt 0) {
            $ports = $result.PortResults
            Write-LogMessage -Message "Port Test: $($ports.OpenPorts.Count) open, $($ports.ClosedPorts.Count) closed/filtered" -FilePath $script:logFile
            if ($ports.OpenPorts.Count -gt 0) {
                Write-LogMessage -Message "  Open Ports: $($ports.OpenPorts -join ', ')" -FilePath $script:logFile
            }
        }

        # MTU Summary
        if ($result.MTUResults.Count -gt 0) {
            $mtu = $result.MTUResults
            if ($mtu.Success) {
                Write-LogMessage -Message "MTU Discovery: Maximum MTU = $($mtu.MaxMTU) bytes" -FilePath $script:logFile
            }
            else {
                Write-LogMessage -Message "MTU Discovery: Failed or no response" -FilePath $script:logFile
            }
        }

        if ($result.Errors.Count -gt 0) {
            Write-LogMessage -Message "Errors: $($result.Errors -join '; ')" -FilePath $script:logFile
        }
    }

    Write-LogMessage -Message "`n========================================" -FilePath $script:logFile
    Write-LogMessage -Message "Log file saved: $script:logFile" -FilePath $script:logFile
    if (Test-Path -Path $script:logFile) {
        Write-LogMessage -Message "Log file size: $(Get-FormattedSize (Get-Item $script:logFile).Length)" -FilePath $script:logFile
    }
    else {
        Write-LogMessage -Message "Log file size: 0 Bytes (WhatIf mode)" -FilePath $script:logFile
    }
    Write-LogMessage -Message "========================================" -FilePath $script:logFile
}

# Main execution
try {
    Write-Information -MessageData "Advanced Network Connectivity Test Starting..." -InformationAction Continue

    # Validate output path
    if (-not (Test-Path -Path $OutputPath)) {
        if ($PSCmdlet.ShouldProcess($OutputPath, "Create Directory")) {
            New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
            Write-Information -MessageData "Created output directory: $OutputPath" -InformationAction Continue
        }
    }

    # Create log file
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $computerName = $env:COMPUTERNAME
    $fileName = "AdvancedNetworkTest_${computerName}_${timestamp}.log"
    $script:logFile = Join-Path -Path $OutputPath -ChildPath $fileName

    # Create log header
    $header = @"
========================================
Advanced Network Connectivity Test Results
========================================
Computer Name: $computerName
Test Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC')
Test Types: $($TestType -join ', ')
Count per Target: $(if($Count -eq 0){"Continuous"}else{$Count})
Parallel Processing: $(if($Parallel){"Enabled"}else{"Disabled"})
Timeout: $Timeout ms
========================================

"@

    if ($PSCmdlet.ShouldProcess($script:logFile, "Create Log File")) {
        Set-Content -Path $script:logFile -Value $header
        Write-Information -MessageData "Log file created: $script:logFile" -InformationAction Continue
    }

    # Prepare target list
    $targetList = @()

    if ($TargetFile) {
        Write-Information -MessageData "Loading targets from file: $TargetFile" -InformationAction Continue
        $importedTargets = Import-TargetsFromFile -FilePath $TargetFile
        foreach ($target in $importedTargets) {
            $targetList += @{
                Target = $target.Target
                Description = $target.Description
                Priority = $target.Priority
            }
        }
    }
    else {
        foreach ($target in $Target) {
            $targetList += @{
                Target = $target
                Description = "Default target"
                Priority = "Medium"
            }
        }
    }

    Write-Information -MessageData "Testing $($targetList.Count) target(s)" -InformationAction Continue

    # Execute tests
    if ($Parallel -and $targetList.Count -gt 1) {
        Write-Information -MessageData "Running tests in parallel..." -InformationAction Continue

        $jobs = @()
        foreach ($targetInfo in $targetList) {
            $jobScriptBlock = {
                param($TargetHost, $Description, $Priority, $TestTypes, $TestCount, $TestTimeout, $TestPorts, $TestMaxMTU)

                # Define Initialize-NetworkTestResult function in job scope
                function Initialize-NetworkTestResult {
                    param([string]$Target)

                    return @{
                        Target = $Target
                        Description = ""
                        Priority = "Medium"
                        PingResults = @{}
                        DNSResults = @{}
                        PortResults = @{}
                        MTUResults = @{}
                        TestStartTime = Get-Date
                        TestEndTime = $null
                        Status = "Running"
                        Errors = @()
                        LogBuffer = @()
                    }
                }

                # Define Write-LogMessage function in job scope
                function Write-LogMessage {
                    param(
                        [string]$Message,
                        [string]$FilePath,
                        [switch]$NoConsole,
                        [ref]$LogBuffer
                    )

                    $timestampedMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC'): $Message"

                    if ($LogBuffer) {
                        $LogBuffer.Value += $timestampedMessage
                    }
                    elseif ($FilePath) {
                        Add-Content -Path $FilePath -Value $timestampedMessage -ErrorAction SilentlyContinue
                    }

                    if (-not $NoConsole) {
                        Write-Information -MessageData $timestampedMessage -InformationAction Continue
                    }
                }

                # Define Test-PingConnectivity function in job scope
                function Test-PingConnectivity {
                    param(
                        [string]$TargetHost,
                        [int]$PingCount,
                        [int]$TimeoutMs,
                        [ref]$LogBuffer
                    )

                    $pingResults = @{
                        Sent = 0
                        Received = 0
                        Lost = 0
                        MinTime = [int]::MaxValue
                        MaxTime = 0
                        AvgTime = 0
                        TotalTime = 0
                        PacketLoss = 0
                        Details = @()
                    }

                    Write-LogMessage -Message "Starting ping test for $TargetHost ($PingCount packets)" -LogBuffer $LogBuffer

                    for ($i = 1; $i -le $PingCount; $i++) {
                        try {
                            $ping = Test-Connection -ComputerName $TargetHost -Count 1 -TimeoutSeconds ($TimeoutMs / 1000) -ErrorAction Stop

                            $responseTime = $ping.ResponseTime
                            $pingResults.Sent++
                            $pingResults.Received++
                            $pingResults.TotalTime += $responseTime

                            if ($responseTime -lt $pingResults.MinTime) { $pingResults.MinTime = $responseTime }
                            if ($responseTime -gt $pingResults.MaxTime) { $pingResults.MaxTime = $responseTime }

                            $pingResults.Details += "Reply from $($ping.Address): time=$($responseTime)ms"

                            Write-LogMessage -Message "Ping $i/$PingCount to ${TargetHost}: $($responseTime)ms" -LogBuffer $LogBuffer
                        }
                        catch {
                            $pingResults.Sent++
                            $pingResults.Lost++
                            $pingResults.Details += "Request timeout for ping $i"
                            Write-LogMessage -Message "Ping $i/$PingCount to ${TargetHost}: Request timeout" -LogBuffer $LogBuffer
                        }

                        if ($i -lt $PingCount) {
                            Start-Sleep -Milliseconds 1000
                        }
                    }

                    if ($pingResults.Received -gt 0) {
                        $pingResults.AvgTime = [math]::Round($pingResults.TotalTime / $pingResults.Received, 2)
                    }

                    if ($pingResults.MinTime -eq [int]::MaxValue) {
                        $pingResults.MinTime = 0
                    }

                    $pingResults.PacketLoss = [math]::Round(($pingResults.Lost / $pingResults.Sent) * 100, 2)

                    return $pingResults
                }

                # Define Test-DNSResolution function in job scope
                function Test-DNSResolution {
                    param(
                        [string]$TargetHost,
                        [ref]$LogBuffer
                    )

                    $dnsResults = @{
                        HostName = $TargetHost
                        IPAddresses = @()
                        ResolutionTime = 0
                        Success = $false
                        ErrorMessage = ""
                    }

                    Write-LogMessage -Message "Starting DNS resolution test for $TargetHost" -LogBuffer $LogBuffer

                    try {
                        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                        $dnsResult = Resolve-DnsName -Name $TargetHost -ErrorAction Stop
                        $stopwatch.Stop()

                        $dnsResults.ResolutionTime = $stopwatch.ElapsedMilliseconds
                        $dnsResults.Success = $true

                        foreach ($record in $dnsResult) {
                            if ($record.IPAddress) {
                                $dnsResults.IPAddresses += $record.IPAddress
                            }
                        }

                        Write-LogMessage -Message "DNS resolution for $TargetHost successful: $($dnsResults.IPAddresses -join ', ') ($($dnsResults.ResolutionTime)ms)" -LogBuffer $LogBuffer
                    }
                    catch {
                        $dnsResults.ErrorMessage = $_.Exception.Message
                        Write-LogMessage -Message "DNS resolution for $TargetHost failed: $($_.Exception.Message)" -LogBuffer $LogBuffer
                    }

                    return $dnsResults
                }

                # Define Test-PortConnectivity function in job scope
                function Test-PortConnectivity {
                    param(
                        [string]$TargetHost,
                        [int[]]$PortList,
                        [int]$TimeoutMs,
                        [ref]$LogBuffer
                    )

                    $portResults = @{
                        TestedPorts = @()
                        OpenPorts = @()
                        ClosedPorts = @()
                        Results = @{}
                    }

                    Write-LogMessage -Message "Starting port connectivity test for $TargetHost on ports: $($PortList -join ', ')" -LogBuffer $LogBuffer

                    foreach ($port in $PortList) {
                        $portResults.TestedPorts += $port

                        try {
                            $tcpClient = New-Object System.Net.Sockets.TcpClient
                            $connectTask = $tcpClient.ConnectAsync($TargetHost, $port)

                            if ($connectTask.Wait($TimeoutMs)) {
                                if ($tcpClient.Connected) {
                                    $portResults.OpenPorts += $port
                                    $portResults.Results[$port] = "Open"
                                    Write-LogMessage -Message "Port $port on ${TargetHost}: Open" -LogBuffer $LogBuffer
                                }
                                else {
                                    $portResults.ClosedPorts += $port
                                    $portResults.Results[$port] = "Closed"
                                    Write-LogMessage -Message "Port $port on ${TargetHost}: Closed" -LogBuffer $LogBuffer
                                }
                            }
                            else {
                                $portResults.ClosedPorts += $port
                                $portResults.Results[$port] = "Timeout"
                                Write-LogMessage -Message "Port $port on ${TargetHost}: Timeout" -LogBuffer $LogBuffer
                            }

                            $tcpClient.Close()
                        }
                        catch {
                            $portResults.ClosedPorts += $port
                            $portResults.Results[$port] = "Error: $($_.Exception.Message)"
                            Write-LogMessage -Message "Port $port on ${TargetHost}: Error - $($_.Exception.Message)" -LogBuffer $LogBuffer
                        }
                    }

                    return $portResults
                }

                # Define Test-MTUDiscovery function in job scope
                function Test-MTUDiscovery {
                    param(
                        [string]$TargetHost,
                        [int]$MaxMTUSize,
                        [ref]$LogBuffer
                    )

                    $mtuResults = @{
                        MaxMTU = 0
                        OptimalMTU = 0
                        TestResults = @()
                        Success = $false
                    }

                    Write-LogMessage -Message "Starting MTU discovery for $TargetHost (max size: $MaxMTUSize)" -LogBuffer $LogBuffer

                    # Start with common MTU sizes and work up
                    $testSizes = @(576, 1024, 1280, 1460, 1500)
                    if ($MaxMTUSize -gt 1500) {
                        $testSizes += @(4000, 8000, $MaxMTUSize)
                    }

                    foreach ($size in $testSizes | Sort-Object) {
                        if ($size -gt $MaxMTUSize) { continue }

                        try {
                            $pingSize = $size - 28
                            # Subtract IP and ICMP headers
                            if ($pingSize -lt 1) { continue }

                            $ping = Test-Connection -ComputerName $TargetHost -BufferSize $pingSize -Count 1 -ErrorAction Stop

                            if ($ping) {
                                $mtuResults.MaxMTU = $size
                                $mtuResults.TestResults += "MTU $size bytes: Success"
                                Write-LogMessage -Message "MTU test for $TargetHost at $size bytes: Success" -LogBuffer $LogBuffer
                            }
                        }
                        catch {
                            $mtuResults.TestResults += "MTU $size bytes: Failed"
                            Write-LogMessage -Message "MTU test for $TargetHost at $size bytes: Failed" -LogBuffer $LogBuffer
                            break
                        }
                    }

                    if ($mtuResults.MaxMTU -gt 0) {
                        $mtuResults.OptimalMTU = $mtuResults.MaxMTU
                        $mtuResults.Success = $true
                    }

                    return $mtuResults
                }

                # Define Test-SingleTarget function in job scope
                function Test-SingleTarget {
                    param(
                        [string]$TargetHost,
                        [string]$Description = "",
                        [string]$Priority = "Medium",
                        [string[]]$TestTypes = @("Ping", "DNS"),
                        [int]$TestCount = 10,
                        [int]$TestTimeout = 5000,
                        [int[]]$TestPorts = @(80, 443, 53),
                        [int]$TestMaxMTU = 1500
                    )

                    $testResult = Initialize-NetworkTestResult -Target $TargetHost
                    $testResult.Description = $Description
                    $testResult.Priority = $Priority

                    try {
                        # Ping Test
                        if ($TestTypes -contains "Ping" -or $TestTypes -contains "All") {
                            $testResult.PingResults = Test-PingConnectivity -TargetHost $TargetHost -PingCount $TestCount -TimeoutMs $TestTimeout -LogBuffer ([ref]$testResult.LogBuffer)
                        }

                        # DNS Test
                        if ($TestTypes -contains "DNS" -or $TestTypes -contains "All") {
                            $testResult.DNSResults = Test-DNSResolution -TargetHost $TargetHost -LogBuffer ([ref]$testResult.LogBuffer)
                        }

                        # Port Test
                        if ($TestTypes -contains "Port" -or $TestTypes -contains "All") {
                            $testResult.PortResults = Test-PortConnectivity -TargetHost $TargetHost -PortList $TestPorts -TimeoutMs $TestTimeout -LogBuffer ([ref]$testResult.LogBuffer)
                        }

                        # MTU Test
                        if ($TestTypes -contains "MTU" -or $TestTypes -contains "All") {
                            $testResult.MTUResults = Test-MTUDiscovery -TargetHost $TargetHost -MaxMTUSize $TestMaxMTU -LogBuffer ([ref]$testResult.LogBuffer)
                        }

                        $testResult.Status = "Completed"
                        $testResult.TestEndTime = Get-Date
                    }
                    catch {
                        $testResult.Status = "Failed"
                        $testResult.Errors += $_.Exception.Message
                        $testResult.TestEndTime = Get-Date
                        Write-LogMessage -Message "Tests failed for target: $TargetHost - $($_.Exception.Message)" -LogBuffer ([ref]$testResult.LogBuffer)
                    }

                    return $testResult
                }

                # Execute the test for this target
                return Test-SingleTarget -TargetHost $TargetHost -Description $Description -Priority $Priority -TestTypes $TestTypes -TestCount $TestCount -TestTimeout $TestTimeout -TestPorts $TestPorts -TestMaxMTU $TestMaxMTU
            }

            $job = Start-Job -ScriptBlock $jobScriptBlock -ArgumentList $targetInfo.Target, $targetInfo.Description, $targetInfo.Priority, $TestType, $Count, $Timeout, $Ports, $MaxMTU
            $jobs += $job
        }

        # Wait for all jobs to complete
        Write-Information -MessageData "Waiting for parallel jobs to complete..." -InformationAction Continue
        $jobs | Wait-Job | Out-Null

        # Collect results
        foreach ($job in $jobs) {
            $result = Receive-Job -Job $job -ErrorAction SilentlyContinue
            if ($result) {
                $script:results[$result.Target] = $result
            }
        }

        # Clean up jobs
        $jobs | Remove-Job -Force
    }
    else {
        Write-Information -MessageData "Running tests sequentially..." -InformationAction Continue

        foreach ($targetInfo in $targetList) {
            $result = Test-SingleTarget -TargetHost $targetInfo.Target -Description $targetInfo.Description -Priority $targetInfo.Priority -TestTypes $TestType -TestCount $Count -TestTimeout $Timeout -TestPorts $Ports -TestMaxMTU $MaxMTU
            $script:results[$result.Target] = $result
        }
    }

    # Write organized target sections to log file
    if ($PSCmdlet.ShouldProcess($script:logFile, "Write Target Test Sections")) {
        foreach ($targetName in $script:results.Keys | Sort-Object) {
            Write-TargetLogSection -TestResult $script:results[$targetName] -FilePath $script:logFile
        }
    }

    # Write summary
    Write-TestSummary -AllResults $script:results

    Write-Information -MessageData "`nAdvanced Network Connectivity Test Completed Successfully" -InformationAction Continue
    Write-Information -MessageData "Results saved to: $script:logFile" -InformationAction Continue
}
catch {
    Write-Error -Message "Error during network connectivity test: $($_.Exception.Message)"
    if ($script:logFile) {
        Write-LogMessage -Message "ERROR: $($_.Exception.Message)" -FilePath $script:logFile
    }
    exit 1
}
