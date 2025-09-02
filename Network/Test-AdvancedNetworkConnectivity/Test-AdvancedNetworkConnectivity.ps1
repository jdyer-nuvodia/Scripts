# =============================================================================
# Script: Test-AdvancedNetworkConnectivity.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.7.2
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
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
    - Support for both continuous loop testing and single-run testing
    - Aggregated statistics across multiple test iterations

    The script uses intelligent defaults including common DNS servers and connectivity test targets
    to enable immediate testing without requiring target specification. By default, the script runs
    in continuous loop mode, performing repeated test cycles and aggregating results until manually stopped.

.PARAMETER Target
    Array of target hosts to test. Can be IP addresses, hostnames, or FQDNs.
    Default: @("8.8.8.8", "1.1.1.1", "microsoft.com")

.PARAMETER TargetFile
    Path to CSV file containing targets to test. CSV format: Target, Description, Priority

.PARAMETER Count
    Number of tests to perform per target in each loop iteration.
    Default: 10

.PARAMETER Loop
    Enable continuous loop testing mode. When enabled, the script runs test cycles continuously until manually stopped with Ctrl+C.
    Default: $true

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

.PARAMETER IncludeLocalNetwork
    Include local network testing (default gateway and traceroute).
    Default: $true

.PARAMETER IncludeResultAnalysis
    Include comprehensive result analysis and scoring.
    Default: $true

.PARAMETER WhatIf
    Shows what would be performed without executing the operations.

.EXAMPLE
    .\Test-AdvancedNetworkConnectivity.ps1
    Performs comprehensive testing (All test types) on default targets with continuous loop mode and local network analysis

.EXAMPLE
    .\Test-AdvancedNetworkConnectivity.ps1 -Target @("google.com", "cloudflare.com") -Count 50 -TestType All -Loop $true
    Performs comprehensive testing on specified targets with 50 tests per cycle in continuous loop mode

.EXAMPLE
    .\Test-AdvancedNetworkConnectivity.ps1 -TargetFile "C:\targets.csv" -Parallel -TestType @("Ping", "Port") -Ports @(80, 443, 22) -Loop $false
    Tests targets from CSV file with parallel processing, testing ping and specific ports in single-run mode

.EXAMPLE
    .\Test-AdvancedNetworkConnectivity.ps1 -Target "server01.domain.com" -TestType MTU -MaxMTU 9000 -IncludeLocalNetwork $false -Loop $false
    Performs MTU discovery on specified target up to 9000 bytes without local network testing in single-run mode

.NOTES
    Validation Requirements: Verify network connectivity, file system access, DNS resolution capabilities
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string[]]$Target = @("8.8.8.8", "1.1.1.1", "microsoft.com"),

    [string]$TargetFile,

    [int]$Count = 10,

    [bool]$Loop = $true,

    [ValidateSet("Ping", "DNS", "Port", "MTU", "All")]
    [string[]]$TestType = @("All"),

    [int[]]$Ports = @(80, 443, 53),

    [string]$OutputPath = $PSScriptRoot,

    [bool]$Parallel = $true,

    [int]$MaxMTU = 1500,

    [int]$Timeout = 5000,

    [bool]$IncludeLocalNetwork = $true,

    [bool]$IncludeResultAnalysis = $true
)

# Initialize script variables
$script:logFile = $null
$script:results = @{}
$script:interrupted = $false

# Loop aggregation variables (matching Test-NetworkConnectivity.ps1 exactly)
$script:totalTestRuns = 0
$script:totalTargetTests = 0
$script:totalSuccessfulTests = 0
$script:totalFailedTests = 0
$script:aggregatedResults = @{}
$script:loopStartTime = $null

# Target test results structure
function Initialize-NetworkTestResult {
    param([string]$Target)

    return @{
        Target = $Target
        Description   = ""
        Priority      = "Medium"
        PingResults   = @{}
        DNSResults    = @{}
        PortResults   = @{}
        MTUResults    = @{}
        TestStartTime = Get-Date
        TestEndTime   = $null
        Status        = "Running"
        Errors        = @()
        LogBuffer     = @()
    }
}

function Write-FinalLoopStatistic {
    param([switch]$Interrupted)

    if ($script:logFile) {
        try {
            $testDuration = if ($script:loopStartTime) { (Get-Date) - $script:loopStartTime } else { New-TimeSpan }
            $successRate = if ($script:totalTargetTests -gt 0) { ($script:totalSuccessfulTests / $script:totalTargetTests * 100) } else { 0 }

            $finalStats = @"

========================================
Final Loop Statistics $(if($Interrupted) { "(Script Interrupted)"}):
========================================
Test Duration: $testDuration
Total Test Runs: $script:totalTestRuns
Total Target Tests: $script:totalTargetTests
Successful Tests: $script:totalSuccessfulTests
Failed Tests: $script:totalFailedTests
Success Rate: $($successRate.ToString('N2'))%
========================================
Test completed$(if($Interrupted) { " (Interrupted)"}): $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Log file size: $(Get-FormattedSize (Get-Item $script:logFile).Length)
========================================
"@
            # Force write the final statistics
            $finalStats | Out-File -FilePath $script:logFile -Append -Force

            Write-Information -MessageData $finalStats -InformationAction Continue

            # Add clear message about log file location
            $logInfo = @"

==================================================
Log file has been saved:
Name: $(Split-Path $script:logFile -Leaf)
Location: $(Split-Path $script:logFile)
Full Path: $script:logFile
Size: $(Get-FormattedSize (Get-Item $script:logFile).Length)
==================================================
"@
            Write-Information -MessageData $logInfo -InformationAction Continue

            # Ensure file is flushed
            [System.IO.File]::WriteAllText($script:logFile, (Get-Content $script:logFile -Raw))
        } catch {
            Write-Error "Error writing final statistics: $_"
        }
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
    } elseif ($FilePath) {
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

    if ($Size -gt 1GB) { return " { 0:N2} GB" -f ($Size / 1GB) }
    if ($Size -gt 1MB) { return " { 0:N2} MB" -f ($Size / 1MB) }
    if ($Size -gt 1KB) { return " { 0:N2} KB" -f ($Size / 1KB) }
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
    } catch {
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

            $responseTime = $ping.Latency
            $pingResults.Sent++
            $pingResults.Received++
            $pingResults.TotalTime += $responseTime

            if ($responseTime -lt $pingResults.MinTime) { $pingResults.MinTime = $responseTime }
            if ($responseTime -gt $pingResults.MaxTime) { $pingResults.MaxTime = $responseTime }

            $pingResults.Details += "Reply from $($ping.Address): time = $($responseTime)ms"

            Write-LogMessage -Message "Ping $i/$PingCount to ${ TargetHost}: $($responseTime)ms" -LogBuffer $LogBuffer
        } catch {
            $pingResults.Sent++
            $pingResults.Lost++
            $pingResults.Details += "Request timeout for ping $i"
            Write-LogMessage -Message "Ping $i/$PingCount to ${ TargetHost}: Request timeout" -LogBuffer $LogBuffer
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
    } catch {
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
        OpenPorts   = @()
        ClosedPorts = @()
        Results     = @{}
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
                    Write-LogMessage -Message "Port $port on ${ TargetHost}: Open" -LogBuffer $LogBuffer
                } else {
                    $portResults.ClosedPorts += $port
                    $portResults.Results[$port] = "Closed"
                    Write-LogMessage -Message "Port $port on ${ TargetHost}: Closed" -LogBuffer $LogBuffer
                }
            } else {
                $portResults.ClosedPorts += $port
                $portResults.Results[$port] = "Timeout"
                Write-LogMessage -Message "Port $port on ${ TargetHost}: Timeout" -LogBuffer $LogBuffer
            }

            $tcpClient.Close()
        } catch {
            $portResults.ClosedPorts += $port
            $portResults.Results[$port] = "Error: $($_.Exception.Message)"
            Write-LogMessage -Message "Port $port on ${ TargetHost}: Error - $($_.Exception.Message)" -LogBuffer $LogBuffer
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
        } catch {
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
        [string]$Priority    = "Medium",
        [string[]]$TestTypes = @("Ping", "DNS"),
        [int]$TestCount      = 10,
        [int]$TestTimeout    = 5000,
        [int[]]$TestPorts    = @(80, 443, 53),
        [int]$TestMaxMTU     = 1500
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
    } catch {
        $testResult.Status = "Failed"
        $testResult.Errors += $_.Exception.Message
        $testResult.TestEndTime = Get-Date
        Write-LogMessage -Message "Tests failed for target: $TargetHost - $($_.Exception.Message)" -LogBuffer ([ref]$testResult.LogBuffer)
    }

    return $testResult
}

function Update-AggregatedStatistic {
    [CmdletBinding(SupportsShouldProcess)]
    param([hashtable]$TestResults)

    if ($PSCmdlet.ShouldProcess("Script aggregated statistics", "Update aggregated test statistics")) {
        foreach ($targetName in $TestResults.Keys) {
            $result = $TestResults[$targetName]

            # Initialize aggregated results for this target if not exists
            if (-not $script:aggregatedResults.ContainsKey($targetName)) {
                $script:aggregatedResults[$targetName] = @{
                    Target = $result.Target
                    TotalTests      = 0
                    SuccessfulTests = 0
                    FailedTests     = 0
                    PingStats       = @{
                        TotalSent     = 0
                        TotalReceived = 0
                        TotalLost     = 0
                        TotalTime     = 0
                        MinTime       = [int]::MaxValue
                        MaxTime       = 0
                    }
                    DNSStats = @{
                        TotalAttempts       = 0
                        SuccessfulAttempts  = 0
                        TotalResolutionTime = 0
                    }
                    PortStats = @{
                        TotalPortTests      = 0
                        SuccessfulPortTests = 0
                    }
                    MTUStats = @{
                        TotalMTUTests      = 0
                        SuccessfulMTUTests = 0
                        MaxMTUFound        = 0
                    }
                }
            }

            $agg = $script:aggregatedResults[$targetName]
            $agg.TotalTests++

            if ($result.Status -eq "Completed") {
                $agg.SuccessfulTests++
                $script:totalSuccessfulTests++
            } else {
                $agg.FailedTests++
                $script:totalFailedTests++
            }

            # Aggregate ping statistics
            if ($result.PingResults.Sent -gt 0) {
                $ping = $result.PingResults
                $agg.PingStats.TotalSent += $ping.Sent
                $agg.PingStats.TotalReceived += $ping.Received
                $agg.PingStats.TotalLost += $ping.Lost
                $agg.PingStats.TotalTime += $ping.TotalTime
                if ($ping.MinTime -lt $agg.PingStats.MinTime -and $ping.MinTime -gt 0) {
                    $agg.PingStats.MinTime = $ping.MinTime
                }
                if ($ping.MaxTime -gt $agg.PingStats.MaxTime) {
                    $agg.PingStats.MaxTime = $ping.MaxTime
                }
            }

            # Aggregate DNS statistics
            if ($null -ne $result.DNSResults.Success) {
                $agg.DNSStats.TotalAttempts++
                if ($result.DNSResults.Success) {
                    $agg.DNSStats.SuccessfulAttempts++
                    $agg.DNSStats.TotalResolutionTime += $result.DNSResults.ResolutionTime
                }
            }

            # Aggregate port statistics
            if ($result.PortResults.TestedPorts.Count -gt 0) {
                $agg.PortStats.TotalPortTests += $result.PortResults.TestedPorts.Count
                $agg.PortStats.SuccessfulPortTests += $result.PortResults.OpenPorts.Count
            }

            # Aggregate MTU statistics
            if ($null -ne $result.MTUResults.Success) {
                $agg.MTUStats.TotalMTUTests++
                if ($result.MTUResults.Success) {
                    $agg.MTUStats.SuccessfulMTUTests++
                    if ($result.MTUResults.MaxMTU -gt $agg.MTUStats.MaxMTUFound) {
                        $agg.MTUStats.MaxMTUFound = $result.MTUResults.MaxMTU
                    }
                }
            }

            $script:totalTargetTests++
        }
    }
}

function Write-LoopIterationSummary {
    param(
        [int]$IterationNumber,
        [hashtable]$IterationResults
    )

    Write-LogMessage -Message "`n========================================" -FilePath $script:logFile
    Write-LogMessage -Message "LOOP ITERATION #$IterationNumber SUMMARY" -FilePath $script:logFile
    Write-LogMessage -Message "========================================" -FilePath $script:logFile

    foreach ($targetName in $IterationResults.Keys | Sort-Object) {
        $result = $IterationResults[$targetName]
        Write-LogMessage -Message "Target: $($result.Target) - Status: $($result.Status)" -FilePath $script:logFile

        # Brief summary for each target
        if ($result.PingResults.Sent -gt 0) {
            $ping = $result.PingResults
            Write-LogMessage -Message "  Ping: $($ping.Received)/$($ping.Sent) successful ($($ping.PacketLoss)% loss), Avg: $($ping.AvgTime)ms" -FilePath $script:logFile
        }

        if ($null -ne $result.DNSResults.Success) {
            $dns = $result.DNSResults
            Write-LogMessage -Message "  DNS: $(if($dns.Success) { 'Success'}else { 'Failed'}) $(if($dns.Success) { '(' + $dns.ResolutionTime + 'ms)'}else { ''})" -FilePath $script:logFile
        }

        if ($result.PortResults.TestedPorts.Count -gt 0) {
            $ports = $result.PortResults
            Write-LogMessage -Message "  Ports: $($ports.OpenPorts.Count)/$($ports.TestedPorts.Count) open" -FilePath $script:logFile
        }

        if ($null -ne $result.MTUResults.Success) {
            $mtu = $result.MTUResults
            Write-LogMessage -Message "  MTU: $(if($mtu.Success) { $mtu.MaxMTU.ToString()}else { 'Failed'})" -FilePath $script:logFile
        }
    }

    Write-LogMessage -Message "========================================" -FilePath $script:logFile
}

function Write-AggregatedStatistic {
    param([int]$IterationNumber)

    if ($IterationNumber % 10 -eq 0) {
        Write-LogMessage -Message "`n========================================" -FilePath $script:logFile
        Write-LogMessage -Message "AGGREGATED STATISTICS (After $IterationNumber runs)" -FilePath $script:logFile
        Write-LogMessage -Message "========================================" -FilePath $script:logFile

        $overallSuccessRate = if ($script:totalTargetTests -gt 0) { ($script:totalSuccessfulTests / $script:totalTargetTests * 100) } else { 0 }
        Write-LogMessage -Message "Overall Success Rate: $([Math]::Round($overallSuccessRate, 2))% ($script:totalSuccessfulTests/$script:totalTargetTests)" -FilePath $script:logFile

        foreach ($targetName in $script:aggregatedResults.Keys | Sort-Object) {
            $agg = $script:aggregatedResults[$targetName]
            Write-LogMessage -Message "`nTarget: $($agg.Target)" -FilePath $script:logFile
            Write-LogMessage -Message "  Total Tests: $($agg.TotalTests) (Success: $($agg.SuccessfulTests), Failed: $($agg.FailedTests))" -FilePath $script:logFile

            # Ping aggregation
            if ($agg.PingStats.TotalSent -gt 0) {
                $avgPingTime = if ($agg.PingStats.TotalReceived -gt 0) { $agg.PingStats.TotalTime / $agg.PingStats.TotalReceived } else { 0 }
                $pingLossRate = ($agg.PingStats.TotalLost / $agg.PingStats.TotalSent) * 100
                Write-LogMessage -Message "  Ping Totals: $($agg.PingStats.TotalReceived)/$($agg.PingStats.TotalSent) successful ($([Math]::Round($pingLossRate, 2))% loss)" -FilePath $script:logFile
                if ($agg.PingStats.TotalReceived -gt 0) {
                    $minTime = if ($agg.PingStats.MinTime -eq [int]::MaxValue) { 0 } else { $agg.PingStats.MinTime }
                    Write-LogMessage -Message "  Ping Times: Min = $($minTime)ms, Max = $($agg.PingStats.MaxTime)ms, Avg = $([Math]::Round($avgPingTime, 2))ms" -FilePath $script:logFile
                }
            }

            # DNS aggregation
            if ($agg.DNSStats.TotalAttempts -gt 0) {
                $dnsSuccessRate = ($agg.DNSStats.SuccessfulAttempts / $agg.DNSStats.TotalAttempts) * 100
                $avgDNSTime = if ($agg.DNSStats.SuccessfulAttempts -gt 0) { $agg.DNSStats.TotalResolutionTime / $agg.DNSStats.SuccessfulAttempts } else { 0 }
                Write-LogMessage -Message "  DNS Success Rate: $([Math]::Round($dnsSuccessRate, 2))% ($($agg.DNSStats.SuccessfulAttempts)/$($agg.DNSStats.TotalAttempts))" -FilePath $script:logFile
                if ($agg.DNSStats.SuccessfulAttempts -gt 0) {
                    Write-LogMessage -Message "  DNS Avg Time: $([Math]::Round($avgDNSTime, 2))ms" -FilePath $script:logFile
                }
            }

            # Port aggregation
            if ($agg.PortStats.TotalPortTests -gt 0) {
                $portSuccessRate = ($agg.PortStats.SuccessfulPortTests / $agg.PortStats.TotalPortTests) * 100
                Write-LogMessage -Message "  Port Success Rate: $([Math]::Round($portSuccessRate, 2))% ($($agg.PortStats.SuccessfulPortTests)/$($agg.PortStats.TotalPortTests))" -FilePath $script:logFile
            }

            # MTU aggregation
            if ($agg.MTUStats.TotalMTUTests -gt 0) {
                $mtuSuccessRate = ($agg.MTUStats.SuccessfulMTUTests / $agg.MTUStats.TotalMTUTests) * 100
                Write-LogMessage -Message "  MTU Success Rate: $([Math]::Round($mtuSuccessRate, 2))% ($($agg.MTUStats.SuccessfulMTUTests)/$($agg.MTUStats.TotalMTUTests))" -FilePath $script:logFile
                if ($agg.MTUStats.MaxMTUFound -gt 0) {
                    Write-LogMessage -Message "  Max MTU Found: $($agg.MTUStats.MaxMTUFound)" -FilePath $script:logFile
                }
            }
        }

        Write-LogMessage -Message "========================================" -FilePath $script:logFile
        Write-Information -MessageData "Aggregated statistics updated after $IterationNumber test runs" -InformationAction Continue
    }
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
        if ($result.PingResults.Sent -gt 0) {
            $ping = $result.PingResults
            Write-LogMessage -Message "Ping Results: $($ping.Received)/$($ping.Sent) successful ($($ping.PacketLoss)% loss)" -FilePath $script:logFile
            if ($ping.Received -gt 0) {
                Write-LogMessage -Message "  Latency: Min = $($ping.MinTime)ms, Max = $($ping.MaxTime)ms, Avg = $($ping.AvgTime)ms" -FilePath $script:logFile
            }
        }

        # DNS Summary
        if ($null -ne $result.DNSResults.Success) {
            $dns = $result.DNSResults
            if ($dns.Success) {
                Write-LogMessage -Message "DNS Resolution: Success ($($dns.ResolutionTime)ms) - $($dns.IPAddresses -join ', ')" -FilePath $script:logFile
            } else {
                Write-LogMessage -Message "DNS Resolution: Failed - $($dns.ErrorMessage)" -FilePath $script:logFile
            }
        }

        # Port Summary
        if ($result.PortResults.TestedPorts.Count -gt 0) {
            $ports = $result.PortResults
            Write-LogMessage -Message "Port Test: $($ports.OpenPorts.Count) open, $($ports.ClosedPorts.Count) closed/filtered" -FilePath $script:logFile
            if ($ports.OpenPorts.Count -gt 0) {
                Write-LogMessage -Message "  Open Ports: $($ports.OpenPorts -join ', ')" -FilePath $script:logFile
            }
        }

        # MTU Summary
        if ($null -ne $result.MTUResults.Success) {
            $mtu = $result.MTUResults
            if ($mtu.Success) {
                Write-LogMessage -Message "MTU Discovery: Maximum MTU = $($mtu.MaxMTU) bytes" -FilePath $script:logFile
            } else {
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
    } else {
        Write-LogMessage -Message "Log file size: 0 Bytes (WhatIf mode)" -FilePath $script:logFile
    }
    Write-LogMessage -Message "========================================" -FilePath $script:logFile
}

function Get-DefaultGateway {
    try {
        $defaultRoute = Get-NetRoute -DestinationPrefix "0.0.0.0/0" -ErrorAction Stop | Select-Object -First 1
        return $defaultRoute.NextHop
    } catch {
        Write-Warning -Message "Could not determine default gateway: $($_.Exception.Message)"
        return $null
    }
}

function Test-TracerouteConnectivity {
    param(
        [string]$TargetHost,
        [int]$MaxHops = 30,
        [ref]$LogBuffer
    )

    $traceResults = @{
        Target = $TargetHost
        Hops = @()
        TotalHops      = 0
        Success        = $false
        CompletionTime = 0
        FailedHops     = 0
    }

    Write-LogMessage -Message "Starting traceroute to $TargetHost (max $MaxHops hops)" -LogBuffer $LogBuffer

    try {
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

        # Use Test-NetConnection for traceroute functionality
        $traceRoute = Test-NetConnection -ComputerName $TargetHost -TraceRoute -ErrorAction Stop

        $stopwatch.Stop()
        $traceResults.CompletionTime = $stopwatch.ElapsedMilliseconds

        if ($traceRoute.TraceRoute) {
            $hopNumber = 1
            foreach ($hop in $traceRoute.TraceRoute) {
                $hopInfo  = @{
                    HopNumber    = $hopNumber
                    IPAddress    = $hop
                    HostName     = ""
                    ResponseTime = 0
                }

                # Try to resolve hostname
                try {
                    $hostInfo = [System.Net.Dns]::GetHostEntry($hop)
                    $hopInfo.HostName = $hostInfo.HostName
                } catch {
                    $hopInfo.HostName = "Unknown"
                }

                # Test response time for this hop
                try {
                    $hopPing = Test-Connection -ComputerName $hop -Count 1 -ErrorAction Stop
                    $hopInfo.ResponseTime = $hopPing.Latency
                } catch {
                    $hopInfo.ResponseTime = -1
                    $traceResults.FailedHops++
                }

                $traceResults.Hops += $hopInfo
                Write-LogMessage -Message "Hop $hopNumber`: $hop ($($hopInfo.HostName)) - $($hopInfo.ResponseTime)ms" -LogBuffer $LogBuffer
                $hopNumber++
            }

            $traceResults.TotalHops = $traceResults.Hops.Count
            $traceResults.Success = $traceRoute.PingSucceeded
        }

        Write-LogMessage -Message "Traceroute to $TargetHost completed: $($traceResults.TotalHops) hops, $($traceResults.FailedHops) failed" -LogBuffer $LogBuffer
    } catch {
        Write-LogMessage -Message "Traceroute to $TargetHost failed: $($_.Exception.Message)" -LogBuffer $LogBuffer
    }

    return $traceResults
}

function Test-LocalNetworkConnectivity {
    param(
        [int]$TestTimeout = 5000,
        [ref]$LogBuffer
    )

    $localResults = @{
        DefaultGateway = ""
        GatewayReachable = $false
        GatewayLatency = 0
        NetworkAdapters = @()
        TracerouteResults = @()
        LocalNetworkHealth = "Unknown"
    }

    Write-LogMessage -Message "Starting local network connectivity tests" -LogBuffer $LogBuffer

    # Get default gateway
    $gateway = Get-DefaultGateway
    if ($gateway) {
        $localResults.DefaultGateway = $gateway
        Write-LogMessage -Message "Default gateway detected: $gateway" -LogBuffer $LogBuffer

        # Test gateway connectivity
        try {
            $gatewayPing = Test-Connection -ComputerName $gateway -Count 3 -TimeoutSeconds ($TestTimeout / 1000) -ErrorAction Stop
            $localResults.GatewayReachable = $true
            $localResults.GatewayLatency = ($gatewayPing | Measure-Object -Property Latency -Average).Average
            Write-LogMessage -Message "Gateway ping successful: Average latency $($localResults.GatewayLatency)ms" -LogBuffer $LogBuffer
        } catch {
            Write-LogMessage -Message "Gateway ping failed: $($_.Exception.Message)" -LogBuffer $LogBuffer
        }

        # Test traceroute to a common external target through gateway
        $localResults.TracerouteResults = Test-TracerouteConnectivity -TargetHost "8.8.8.8" -LogBuffer $LogBuffer
    } else {
        Write-LogMessage -Message "Could not detect default gateway" -LogBuffer $LogBuffer
    }

    # Get network adapter information
    try {
        $adapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" -and $_.Virtual -eq $false }
        foreach ($adapter in $adapters) {
            $adapterInfo = @{
                Name = $adapter.Name
                InterfaceDescription = $adapter.InterfaceDescription
                LinkSpeed = $adapter.LinkSpeed
                Status = $adapter.Status
                IPAddresses = @()
            }

            # Get IP addresses for this adapter
            $ipConfig = Get-NetIPAddress -InterfaceIndex $adapter.InterfaceIndex -ErrorAction SilentlyContinue
            foreach ($ip in $ipConfig) {
                if ($ip.AddressFamily -eq "IPv4" -and $ip.IPAddress -ne "127.0.0.1") {
                    $adapterInfo.IPAddresses += $ip.IPAddress
                }
            }

            $localResults.NetworkAdapters += $adapterInfo
            Write-LogMessage -Message "Network adapter: $($adapter.Name) - $($adapterInfo.IPAddresses -join ', ')" -LogBuffer $LogBuffer
        }
    } catch {
        Write-LogMessage -Message "Error getting network adapter information: $($_.Exception.Message)" -LogBuffer $LogBuffer
    }

    # Determine local network health
    if ($localResults.GatewayReachable -and $localResults.GatewayLatency -lt 50 -and $localResults.NetworkAdapters.Count -gt 0) {
        $localResults.LocalNetworkHealth = "Excellent"
    } elseif ($localResults.GatewayReachable -and $localResults.GatewayLatency -lt 100) {
        $localResults.LocalNetworkHealth = "Good"
    } elseif ($localResults.GatewayReachable) {
        $localResults.LocalNetworkHealth = "Concerning"
    } else {
        $localResults.LocalNetworkHealth = "Poor"
    }

    return $localResults
}

function Get-NetworkHealthScore {
    param([hashtable]$TestResult)

    $scores = @{
        PingScore    = 0
        DNSScore     = 0
        PortScore    = 0
        MTUScore     = 0
        OverallScore = 0
        HealthStatus = "Unknown"
    }

    # Ping scoring (40% weight)
    if ($TestResult.PingResults.Sent -gt 0) {
        $pingSuccessRate = ($TestResult.PingResults.Received / $TestResult.PingResults.Sent) * 100
        $scores.PingScore = $pingSuccessRate

        # Bonus/penalty for latency
        if ($TestResult.PingResults.Received -gt 0) {
            $avgLatency = $TestResult.PingResults.AvgTime
            if ($avgLatency -le 50) {
                $scores.PingScore = [Math]::Min($scores.PingScore + 5, 100)
            } elseif ($avgLatency -gt 200) {
                $scores.PingScore = [Math]::Max($scores.PingScore - 10, 0)
            }
        }
    }

    # DNS scoring (25% weight)
    if ($null -ne $TestResult.DNSResults.Success) {
        if ($TestResult.DNSResults.Success) {
            # Score based on resolution time
            if ($TestResult.DNSResults.ResolutionTime -le 100) {
                $scores.DNSScore = 100
            } elseif ($TestResult.DNSResults.ResolutionTime -le 500) {
                $scores.DNSScore = 95
            } else {
                $scores.DNSScore = 85
            }
        } else {
            $scores.DNSScore = 0
        }
    }

    # Port scoring (20% weight) - using adjusted scoring that excludes expected failures
    if ($TestResult.PortResults.TestedPorts.Count -gt 0) {
        $scores.PortScore = Get-AdjustedPortScore -PortResults $TestResult.PortResults -Target $TestResult.Target
    }

    # MTU scoring (15% weight)
    if ($null -ne $TestResult.MTUResults.Success) {
        if ($TestResult.MTUResults.Success -and $TestResult.MTUResults.MaxMTU -ge 1500) {
            $scores.MTUScore = 100
        } elseif ($TestResult.MTUResults.Success -and $TestResult.MTUResults.MaxMTU -ge 1200) {
            $scores.MTUScore = 80
        } elseif ($TestResult.MTUResults.Success) {
            $scores.MTUScore = 60
        } else {
            $scores.MTUScore = 0
        }
    }

    # Calculate weighted overall score
    $scores.OverallScore = [Math]::Round(
        ($scores.PingScore * 0.40) +
        ($scores.DNSScore * 0.25) +
        ($scores.PortScore * 0.20) +
        ($scores.MTUScore * 0.15), 2
    )

    # Determine health status
    if ($scores.OverallScore -ge 98) {
        $scores.HealthStatus = "Excellent"
    } elseif ($scores.OverallScore -ge 90) {
        $scores.HealthStatus = "Concerning"
    } else {
        $scores.HealthStatus = "Poor"
    }

    return $scores
}

function Test-ExpectedPortFailure {
    param(
        [string]$Target,
        [int]$Port
    )

    # Convert target to lowercase for comparison
    $targetLower = $Target.ToLower()

    # Expected failures for DNS servers
    if ($targetLower -match '^(8\.8\.8\.8|1\.1\.1\.1|208\.67\.222\.222|208\.67\.220\.220)$') {
        # DNS servers typically don't serve HTTP
        if ($Port -eq 80) {
            return $true
        }
    }

    # Expected failures for web servers
    if ($targetLower -match '\.(com|org|net|edu|gov)$' -or $targetLower -match '^(www\.|web\.|mail\.)') {
        # Most web servers don't run DNS services
        if ($Port -eq 53) {
            return $true
        }
    }

    # Expected failures for specific well-known services
    switch ($targetLower) {
        'microsoft.com' {
            # Microsoft.com doesn't run DNS
            if ($Port -eq 53) { return $true }
        }
        'google.com' {
            # Google.com doesn't run DNS (different from 8.8.8.8)
            if ($Port -eq 53) { return $true }
        }
        'cloudflare.com' {
            # Cloudflare.com website doesn't run DNS
            if ($Port -eq 53) { return $true }
        }
    }

    return $false
}

function Get-AdjustedPortScore {
    param(
        [hashtable]$PortResults,
        [string]$Target
    )

    if ($PortResults.TestedPorts.Count -eq 0) {
        return 0
    }

    $expectedFailures = 0
    $unexpectedFailures = 0
    $successfulPorts = $PortResults.OpenPorts.Count
    $bonusPoints = 0

    # Analyze each tested port
    foreach ($port in $PortResults.TestedPorts) {
        $isOpen = $port -in $PortResults.OpenPorts
        $isExpectedFailure = Test-ExpectedPortFailure -Target $Target -Port $port

        if (-not $isOpen) {
            if ($isExpectedFailure) {
                $expectedFailures++
            } else {
                $unexpectedFailures++
            }
        } elseif ($isExpectedFailure) {
            # Bonus for unexpected successful connections
            $bonusPoints += 5
        }
    }

    # Calculate effective ports tested (excluding expected failures)
    $effectivePortsTested = $PortResults.TestedPorts.Count - $expectedFailures

    if ($effectivePortsTested -eq 0) {
        # All failures were expected, give full score
        return [Math]::Min(100 + $bonusPoints, 100)
    }

    # Calculate success rate based on ports that should work
    $baseScore = ($successfulPorts / $effectivePortsTested) * 100
    $finalScore = [Math]::Min($baseScore + $bonusPoints, 100)

    return [Math]::Round($finalScore, 2)
}

function Write-NetworkAnalysis {
    param(
        [hashtable]$AllResults,
        [hashtable]$LocalNetworkResults = $null,
        [string]$FilePath
    )

    Write-LogMessage -Message "`n========================================" -FilePath $FilePath
    Write-LogMessage -Message "NETWORK ANALYSIS AND SCORING" -FilePath $FilePath
    Write-LogMessage -Message "========================================" -FilePath $FilePath

    $overallScores = @()
    $excellentCount = 0
    $concerningCount = 0
    $poorCount = 0

    # Analyze each target
    foreach ($targetName in $AllResults.Keys) {
        $result = $AllResults[$targetName]
        $scores = Get-NetworkHealthScore -TestResult $result

        Write-LogMessage -Message "`nTarget: $($result.Target)" -FilePath $FilePath
        Write-LogMessage -Message "Health Status: $($scores.HealthStatus)" -FilePath $FilePath
        Write-LogMessage -Message "Overall Score: $($scores.OverallScore)%" -FilePath $FilePath
        Write-LogMessage -Message "  Ping Score: $($scores.PingScore)% (Weight: 40%)" -FilePath $FilePath
        Write-LogMessage -Message "  DNS Score: $($scores.DNSScore)% (Weight: 25%)" -FilePath $FilePath
        Write-LogMessage -Message "  Port Score: $($scores.PortScore)% (Weight: 20%)" -FilePath $FilePath
        Write-LogMessage -Message "  MTU Score: $($scores.MTUScore)% (Weight: 15%)" -FilePath $FilePath

        $overallScores += $scores.OverallScore

        switch ($scores.HealthStatus) {
            "Excellent" { $excellentCount++ }
            "Concerning" { $concerningCount++ }
            "Poor" { $poorCount++ }
        }
    }

    # Overall network assessment
    Write-LogMessage -Message "`n========================================" -FilePath $FilePath
    Write-LogMessage -Message "OVERALL NETWORK ASSESSMENT" -FilePath $FilePath
    Write-LogMessage -Message "========================================" -FilePath $FilePath

    $totalTargets = $AllResults.Keys.Count
    $averageScore = if ($overallScores.Count -gt 0) { ($overallScores | Measure-Object -Average).Average } else { 0 }

    Write-LogMessage -Message "Total Targets Tested: $totalTargets" -FilePath $FilePath
    Write-LogMessage -Message "Average Network Score: $([Math]::Round($averageScore, 2))%" -FilePath $FilePath
    Write-LogMessage -Message "Excellent Targets: $excellentCount ($([Math]::Round(($excellentCount/$totalTargets)*100, 1))%)" -FilePath $FilePath
    Write-LogMessage -Message "Concerning Targets: $concerningCount ($([Math]::Round(($concerningCount/$totalTargets)*100, 1))%)" -FilePath $FilePath
    Write-LogMessage -Message "Poor Targets: $poorCount ($([Math]::Round(($poorCount/$totalTargets)*100, 1))%)" -FilePath $FilePath

    # Overall recommendation
    $overallStatus = if ($averageScore -ge 98) { "Excellent" } elseif ($averageScore -ge 90) { "Concerning" } else { "Poor" }
    Write-LogMessage -Message "`nOverall Network Status: $overallStatus" -FilePath $FilePath

    switch ($overallStatus) {
        "Excellent" {
            Write-LogMessage -Message "Recommendation: Network performance is excellent. No immediate action required." -FilePath $FilePath
        }
        "Concerning" {
            Write-LogMessage -Message "Recommendation: Network performance has some issues. Review concerning targets and consider network optimization." -FilePath $FilePath
        }
        "Poor" {
            Write-LogMessage -Message "Recommendation: Network performance is poor. Immediate investigation and remediation required." -FilePath $FilePath
        }
    }

    # Local network analysis
    if ($LocalNetworkResults) {
        Write-LogMessage -Message "`n========================================" -FilePath $FilePath
        Write-LogMessage -Message "LOCAL NETWORK ANALYSIS" -FilePath $FilePath
        Write-LogMessage -Message "========================================" -FilePath $FilePath

        Write-LogMessage -Message "Default Gateway: $($LocalNetworkResults.DefaultGateway)" -FilePath $FilePath
        Write-LogMessage -Message "Gateway Reachable: $($LocalNetworkResults.GatewayReachable)" -FilePath $FilePath
        if ($LocalNetworkResults.GatewayReachable) {
            Write-LogMessage -Message "Gateway Latency: $([Math]::Round($LocalNetworkResults.GatewayLatency, 2))ms" -FilePath $FilePath
        }
        Write-LogMessage -Message "Local Network Health: $($LocalNetworkResults.LocalNetworkHealth)" -FilePath $FilePath
        Write-LogMessage -Message "Active Network Adapters: $($LocalNetworkResults.NetworkAdapters.Count)" -FilePath $FilePath

        if ($LocalNetworkResults.TracerouteResults.Success) {
            Write-LogMessage -Message "Traceroute to 8.8.8.8: $($LocalNetworkResults.TracerouteResults.TotalHops) hops, $($LocalNetworkResults.TracerouteResults.FailedHops) failed" -FilePath $FilePath
        }
    }
}

# Handle Ctrl+C
$null = [Console]::TreatControlCAsInput = $false

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
Count per Target: $Count
Loop Mode: $(if($Loop) { "Enabled (Continuous)"}else { "Disabled (Single Run)"})
Parallel Processing: $(if($Parallel) { "Enabled"}else { "Disabled"})
Timeout: $Timeout ms
========================================

"@

    if ($PSCmdlet.ShouldProcess($script:logFile, "Create Log File")) {
        Set-Content -Path $script:logFile -Value $header
        Write-Information -MessageData "Log file: $script:logFile" -InformationAction Continue
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
    } else {
        foreach ($target in $Target) {
            $targetList += @{
                Target = $target
                Description = "Default target"
                Priority = "Medium"
            }
        }
    }

    Write-Information -MessageData "Testing $($targetList.Count) target(s)$(if($Loop) { " in continuous loop mode"}else { " in single-run mode"})" -InformationAction Continue
    if ($Loop) {
        Write-Information -MessageData "Press Ctrl+C to stop continuous loop mode" -InformationAction Continue
    }

    # Initialize loop tracking
    $script:loopStartTime = Get-Date
    $script:totalTestRuns = 0

    # Main execution loop (matches Test-NetworkConnectivity.ps1 pattern exactly)
    while (!$script:interrupted) {
        try {
            $script:totalTestRuns++
            $script:results = @{}

            Write-Information -MessageData "`nStarting test run #$script:totalTestRuns..." -InformationAction Continue
            Write-LogMessage -Message "`n======================================== TEST RUN #$script:totalTestRuns ========================================" -FilePath $script:logFile

            # Execute tests for this iteration
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
                                Target        = $Target
                                Description   = ""
                                Priority      = "Medium"
                                PingResults   = @{}
                                DNSResults    = @{}
                                PortResults   = @{}
                                MTUResults    = @{}
                                TestStartTime = Get-Date
                                TestEndTime   = $null
                                Status        = "Running"
                                Errors        = @()
                                LogBuffer     = @()
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
                            } elseif ($FilePath) {
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
                                Sent       = 0
                                Received   = 0
                                Lost       = 0
                                MinTime    = [int]::MaxValue
                                MaxTime    = 0
                                AvgTime    = 0
                                TotalTime  = 0
                                PacketLoss = 0
                                Details    = @()
                            }

                            Write-LogMessage -Message "Starting ping test for $TargetHost ($PingCount packets)" -LogBuffer $LogBuffer

                            for ($i = 1; $i -le $PingCount; $i++) {
                                try {
                                    $ping = Test-Connection -ComputerName $TargetHost -Count 1 -TimeoutSeconds ($TimeoutMs / 1000) -ErrorAction Stop

                                    $responseTime = $ping.Latency
                                    $pingResults.Sent++
                                    $pingResults.Received++
                                    $pingResults.TotalTime += $responseTime

                                    if ($responseTime -lt $pingResults.MinTime) { $pingResults.MinTime = $responseTime }
                                    if ($responseTime -gt $pingResults.MaxTime) { $pingResults.MaxTime = $responseTime }

                                    $pingResults.Details += "Reply from $($ping.Address): time = $($responseTime)ms"

                                    Write-LogMessage -Message "Ping $i/$PingCount to ${ TargetHost}: $($responseTime)ms" -LogBuffer $LogBuffer
                                } catch {
                                    $pingResults.Sent++
                                    $pingResults.Lost++
                                    $pingResults.Details += "Request timeout for ping $i"
                                    Write-LogMessage -Message "Ping $i/$PingCount to ${ TargetHost}: Request timeout" -LogBuffer $LogBuffer
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
                                HostName       = $TargetHost
                                IPAddresses    = @()
                                ResolutionTime = 0
                                Success        = $false
                                ErrorMessage   = ""
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
                            } catch {
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
                                TestedPorts  = @()
                                OpenPorts    = @()
                                ClosedPorts  = @()
                                Results      = @{}
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
                                            Write-LogMessage -Message "Port $port on ${ TargetHost}: Open" -LogBuffer $LogBuffer
                                        } else {
                                            $portResults.ClosedPorts += $port
                                            $portResults.Results[$port] = "Closed"
                                            Write-LogMessage -Message "Port $port on ${ TargetHost}: Closed" -LogBuffer $LogBuffer
                                        }
                                    } else {
                                        $portResults.ClosedPorts += $port
                                        $portResults.Results[$port] = "Timeout"
                                        Write-LogMessage -Message "Port $port on ${ TargetHost}: Timeout" -LogBuffer $LogBuffer
                                    }

                                    $tcpClient.Close()
                                } catch {
                                    $portResults.ClosedPorts += $port
                                    $portResults.Results[$port] = "Error: $($_.Exception.Message)"
                                    Write-LogMessage -Message "Port $port on ${ TargetHost}: Error - $($_.Exception.Message)" -LogBuffer $LogBuffer
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
                                MaxMTU       = 0
                                OptimalMTU   = 0
                                TestResults  = @()
                                Success      = $false
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
                                } catch {
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
                                [string]$Description  = "",
                                [string]$Priority     = "Medium",
                                [string[]]$TestTypes  = @("Ping", "DNS"),
                                [int]$TestCount       = 10,
                                [int]$TestTimeout     = 5000,
                                [int[]]$TestPorts     = @(80, 443, 53),
                                [int]$TestMaxMTU      = 1500
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
                            } catch {
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
            } else {
                Write-Information -MessageData "Running tests sequentially..." -InformationAction Continue

                foreach ($targetInfo in $targetList) {
                    $result = Test-SingleTarget -TargetHost $targetInfo.Target -Description $targetInfo.Description -Priority $targetInfo.Priority -TestTypes $TestType -TestCount $Count -TestTimeout $Timeout -TestPorts $Ports -TestMaxMTU $MaxMTU
                    $script:results[$result.Target] = $result
                }
            }

            # Update aggregated statistics
            Update-AggregatedStatistic -TestResults $script:results

            # Write iteration summary to log
            Write-LoopIterationSummary -IterationNumber $script:totalTestRuns -IterationResults $script:results

            # Write organized target sections to log file for this iteration
            if ($PSCmdlet.ShouldProcess($script:logFile, "Write Target Test Sections")) {
                foreach ($targetName in $script:results.Keys | Sort-Object) {
                    Write-TargetLogSection -TestResult $script:results[$targetName] -FilePath $script:logFile
                }
            }

            # Perform local network testing if enabled (only on first run or every 10th run to avoid log bloat)
            $localNetworkResults = $null
            if ($IncludeLocalNetwork -and ($script:totalTestRuns -eq 1 -or $script:totalTestRuns % 10 -eq 0)) {
                Write-Information -MessageData "Running local network connectivity tests..." -InformationAction Continue
                $localNetworkResults = Test-LocalNetworkConnectivity -TestTimeout $Timeout -LogBuffer ([ref]@())

                # Write local network results to log
                Write-LogMessage -Message "`n========================================" -FilePath $script:logFile
                Write-LogMessage -Message "LOCAL NETWORK TEST RESULTS (Run #$script:totalTestRuns)" -FilePath $script:logFile
                Write-LogMessage -Message "========================================" -FilePath $script:logFile
                Write-LogMessage -Message "Default Gateway: $($localNetworkResults.DefaultGateway)" -FilePath $script:logFile
                Write-LogMessage -Message "Gateway Reachable: $($localNetworkResults.GatewayReachable)" -FilePath $script:logFile
                if ($localNetworkResults.GatewayReachable) {
                    Write-LogMessage -Message "Gateway Latency: $([Math]::Round($localNetworkResults.GatewayLatency, 2))ms" -FilePath $script:logFile
                }
                Write-LogMessage -Message "Local Network Health: $($localNetworkResults.LocalNetworkHealth)" -FilePath $script:logFile
                Write-LogMessage -Message "Active Network Adapters: $($localNetworkResults.NetworkAdapters.Count)" -FilePath $script:logFile

                if ($localNetworkResults.TracerouteResults.Success) {
                    Write-LogMessage -Message "Traceroute Results: $($localNetworkResults.TracerouteResults.TotalHops) hops, $($localNetworkResults.TracerouteResults.FailedHops) failed" -FilePath $script:logFile
                }
            }

            # Write aggregated statistics every 10 iterations
            Write-AggregatedStatistic -IterationNumber $script:totalTestRuns

            # Write comprehensive network analysis if enabled (only on first run or every 10th run)
            if ($IncludeResultAnalysis -and ($script:totalTestRuns -eq 1 -or $script:totalTestRuns % 10 -eq 0)) {
                Write-NetworkAnalysis -AllResults $script:results -LocalNetworkResults $localNetworkResults -FilePath $script:logFile
            }

            Write-Information -MessageData "Test run #$script:totalTestRuns completed" -InformationAction Continue

            # Check if we should stop (single run mode)
            if (-not $Loop) {
                Write-Information -MessageData "`nSingle-run mode: Test completed" -InformationAction Continue
                break
            }

            # Small delay between test runs in loop mode
            if ($Loop) {
                Start-Sleep -Seconds 5
            }

        } catch {
            if ($_.Exception.Message -match "cancelled by the user") {
                $script:interrupted = $true
                break
            }
            throw
        }
    }

    # Write final summary and statistics
    Write-TestSummary -AllResults $script:results

    Write-Information -MessageData "`nAdvanced Network Connectivity Test Completed$(if($script:interrupted) { " (Interrupted)"})" -InformationAction Continue
    Write-Information -MessageData "Results saved to: $script:logFile" -InformationAction Continue
} catch {
    Write-Error -Message "Error during network connectivity test: $($_.Exception.Message)"
    Write-Information -MessageData "Stack Trace: $($_.ScriptStackTrace)" -InformationAction Continue
    if ($script:logFile) {
        Write-LogMessage -Message "ERROR: $($_.Exception.Message)" -FilePath $script:logFile
        Write-LogMessage -Message "Stack Trace: $($_.ScriptStackTrace)" -FilePath $script:logFile
    }
} finally {
    if ($script:interrupted) {
        Write-FinalLoopStatistic -Interrupted
    } else {
        Write-FinalLoopStatistic
    }
}
