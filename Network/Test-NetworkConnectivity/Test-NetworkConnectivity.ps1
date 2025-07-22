# =============================================================================
# Script: Test-NetworkConnectivity.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-22 00:25:07 UTC
# Updated By: jdyer-nuvodia
# Version: 2.17.6
# Additional Info: Updated header with accurate UTC timestamp
# =============================================================================

<#
.SYNOPSIS
    Extended network connectivity test with detailed logging capabilities.
.DESCRIPTION
    Performs comprehensive network connectivity testing with the following features:
     - Continuous or count-based ping testing
     - Detailed network configuration logging
     - Statistical analysis of connectivity
     - Formatted output with color-coded status
     - Automatic log file generation

    Dependencies:
     - Windows PowerShell 5.1 or higher
     - Administrator rights for network configuration

    Security considerations:
     - Requires network access to target
     - Creates log files in specified directory

    Performance impact:
     - Minimal CPU usage
     - Network bandwidth based on ICMP packet size
.PARAMETER Target
    The IP address or hostname to test connectivity against. Default is 8.8.8.8
.PARAMETER Count
    Number of pings to send. 0 means continuous mode. Default is 0
.PARAMETER OutputPath
    Directory path for log files. Default is the script's directory. When run from a module, defaults to the location where the script is stored
.PARAMETER DetailedOutput
    Enables detailed verbose output showing full ping response information
.EXAMPLE
    .\Test-NetworkConnectivity.ps1
    Tests connectivity to 8.8.8.8 continuously
.EXAMPLE
    .\Test-NetworkConnectivity.ps1 -Target "microsoft.com" -Count 100
    Sends 100 pings to microsoft.com
.EXAMPLE
    .\Test-NetworkConnectivity.ps1 -Target "8.8.8.8" -DetailedOutput
    Tests connectivity with detailed response information
.NOTES
    Security Level: Low
    Required Permissions: Network access
    Validation Requirements: Verify network connectivity, file system access
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Target = "8.8.8.8",

    [Parameter(Position = 1)]
    [int]$Count = 0,
    # 0 means continuous
    [Parameter()]
    [string]$OutputPath,

    [Parameter()]
    [switch]$DetailedOutput
)

# Initialize script variables and ensure OutputPath is set correctly
if (-not $OutputPath) {
    # If script is run as a module or from another location, $PSScriptRoot might be empty
    if (-not $PSScriptRoot) {
        $scriptPath = $MyInvocation.MyCommand.Path
        if ($scriptPath) {
            $OutputPath = Split-Path -Parent $scriptPath
        } else {
            # Fallback if we cannot determine script path
            $OutputPath = $PWD.Path
        }
    } else {
        $OutputPath = $PSScriptRoot
    }
}
$script:logFile = $null
$script:sent = 0
$script:received = 0
$script:totalTime = 0
$script:minTime = [int]::MaxValue
$script:maxTime = 0
$script:interrupted = $false

function Write-FinalStatistic {
    param([switch]$Interrupted)

    if ($script:logFile) {
        try {
            $packetLoss     = if ($script:sent -gt 0) { 100 - ($script:received / $script:sent * 100) } else { 0 }
            $avgTime        = if ($script:received -gt 0) { $script:totalTime / $script:received } else { 0 }

            $finalStats = @"

========================================
Final Statistics $(if($Interrupted) { "(Script Interrupted)"}):
========================================
Test Duration: $((Get-Date) - (Get-Item $script:logFile).CreationTime)
Packets: Sent = $script:sent, Received = $script:received, Lost = $($script:sent - $script:received) ($($packetLoss.ToString('N2'))% loss)
Round Trip Times: Min = $(if($script:minTime -eq [int]::MaxValue) { "0"}else { $script:minTime})ms, Max = $script:maxTime`ms, Avg = $($avgTime.ToString('N2'))ms
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
        [switch]$NoConsole
    )

    # Add timestamp to message
    $timestampedMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message"

    # Write to file
    Add-Content -Path $FilePath -Value $timestampedMessage

    # Write to console if not suppressed
    if (!$NoConsole) {
        Write-Information -MessageData $timestampedMessage -InformationAction Continue
    }
}

function Get-FormattedSize {
    param([int64]$Size)

    if ($Size -gt 1GB) { return "{0:N2} GB" -f ($Size / 1GB) }
    if ($Size -gt 1MB) { return "{0:N2} MB" -f ($Size / 1MB) }
    if ($Size -gt 1KB) { return "{0:N2} KB" -f ($Size / 1KB) }
    return "$Size Bytes"
}

function Get-DetailedPingResult {
    param(
        [string]$Target,
        [switch]$DetailedOutput
    )

    $result = @{
        Success = $false
        ResponseTime = 0
        IPAddress = $null
        ResponseMessage = $null
        DetailedOutput = $null
    }

    try {
        # Use native ping command for detailed information
        $pingOutput = ping $Target -n 1 -w 5000 | Out-String

        if ($DetailedOutput) {
            $result.DetailedOutput = $pingOutput
        }
        # Parse ping output for detailed information
        $lines = $pingOutput -split "`r?`n"

        foreach ($line in $lines) {
            $line = $line.Trim()

            # Check for ANY reply format from ping
            if ($line -match "Reply from ([0-9\.]+):") {
                $result.Success = $true
                $result.IPAddress = $matches[1]
                $result.ResponseMessage = $line

                # Try to extract the time
                if ($line -match "time=(\d+)ms" -or $line -match "time=(\d+) ms" -or $line -match "time = (\d+)ms" -or $line -match "time = (\d+) ms") {
                    $result.ResponseTime = [int]$matches[1]
                } elseif ($line -match "time<(\d+)ms" -or $line -match "time<(\d+) ms" -or $line -match "time < (\d+)ms" -or $line -match "time < (\d+) ms") {
                    $result.ResponseTime = 1
                } else {
                    # Default if we can't extract time
                    $result.ResponseTime = 1
                }
                break
            }
            # Check for destination host unreachable
            elseif ($line -match "Reply from ([0-9\.]+): Destination host unreachable") {
                $result.Success = $false
                $result.IPAddress = $matches[1]
                $result.ResponseMessage = $line
                break
            }
            # Check for request timed out
            elseif ($line -match "Request timed out") {
                $result.Success = $false
                $result.ResponseMessage = $line
                break
            }
            # Check for other error messages
            elseif ($line -match "Destination net unreachable|General failure|Unable to contact IP driver") {
                $result.Success = $false
                $result.ResponseMessage = $line
                break
            }
        }

        # If no specific message found but we have output, use a generic timeout
        if (-not $result.ResponseMessage -and $pingOutput) {
            $result.ResponseMessage = "Request timed out."
        }
    } catch {
        $result.Success = $false
        $result.ResponseMessage = "Error executing ping: $_"
    }

    return $result
}

# Handle Ctrl+C
$null = [Console]::TreatControlCAsInput = $false

try {
    # Create output directory if it doesn't exist
    if (!(Test-Path -Path $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        Write-Information -MessageData "Created output directory: $OutputPath" -InformationAction Continue
    }

    # Create timestamp and filename
    $timestamp      = Get-Date -Format "yyyyMMdd_HHmmss"
    $computerName   = $env:COMPUTERNAME
    $fileName       = "PingTest_${computerName}_${timestamp}.log"
    $script:logFile = Join-Path $OutputPath $fileName

    # Create log file with header
    $header = @"
========================================
Extended Ping Test Results
========================================
Computer Name: $computerName
Test Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Target: $Target
Mode: $(if($Count -eq 0) { "Continuous"}else { "Count: $Count"})
========================================

"@
    Set-Content -Path $script:logFile -Value $header

    Write-Information -MessageData "Starting network test - Results will be saved to: $script:logFile" -InformationAction Continue
    Write-Information -MessageData "Press Ctrl+C to stop continuous mode" -InformationAction Continue

    # Get and log network configuration
    Write-LogMessage -Message "Getting network configuration..." -FilePath $script:logFile
    Write-LogMessage -Message "`nNETWORK CONFIGURATION:" -FilePath $script:logFile
    Write-LogMessage -Message "----------------------------------------" -FilePath $script:logFile

    try {
        $ipConfig = & ipconfig /all
        if ($ipConfig) {
            $filteredConfig     = @()
            $currentAdapter     = @()
            $skipAdapter        = $false

            foreach ($line in $ipConfig) {
                # Check if we're starting a new adapter section
                if ($line -match "^[A-Za-z].*adapter.*:$") {
                    # Process previous adapter if we have one
                    if ($currentAdapter.Count -gt 0) {
                        if (-not $skipAdapter) {
                            $filteredConfig += $currentAdapter
                        }
                    }
                    # Start new adapter
                    $currentAdapter     = @($line)
                    $skipAdapter        = $false
                } elseif ($line -match "Media State.*Media disconnected") {
                    # Mark this adapter to be skipped
                    $skipAdapter = $true
                    $currentAdapter += $line
                } else {
                    # Add line to current adapter
                    $currentAdapter += $line
                }
            }

            # Process the last adapter
            if ($currentAdapter.Count -gt 0 -and -not $skipAdapter) {
                $filteredConfig += $currentAdapter
            }

            # Output filtered configuration
            foreach ($line in $filteredConfig) {
                Add-Content -Path $script:logFile -Value $line
                Write-Information -MessageData $line -InformationAction Continue
            }

            $skippedCount = ($ipConfig.Count - $filteredConfig.Count)
            Write-Information -MessageData "Network configuration captured successfully ($($filteredConfig.Count) lines, $skippedCount disconnected adapter lines filtered)" -InformationAction Continue
        } else {
            Write-LogMessage -Message "No network configuration data returned" -FilePath $script:logFile
            Write-Information -MessageData "Warning: No network configuration data returned" -InformationAction Continue
        }
    } catch {
        Write-LogMessage -Message "Error getting network configuration: $_" -FilePath $script:logFile
        Write-Information -MessageData "Error getting network configuration: $_" -InformationAction Continue
    }
    Write-LogMessage -Message "----------------------------------------`n" -FilePath $script:logFile

    # Start ping test
    Write-LogMessage -Message "Starting ping test to $Target..." -FilePath $script:logFile
    if ($DetailedOutput) {
        Write-LogMessage -Message "Detailed output mode enabled - full ping information will be shown" -FilePath $script:logFile
    }

    while (!$script:interrupted) {
        try {
            $pingResult = Get-DetailedPingResult -Target $Target -DetailedOutput:$DetailedOutput

            $script:sent++
            $currentTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

            if ($pingResult.Success) {
                $script:received++
                $responseTime = $pingResult.ResponseTime
                $script:totalTime += $responseTime
                $script:minTime = [Math]::Min($script:minTime, $responseTime)
                $script:maxTime = [Math]::Max($script:maxTime, $responseTime)

                $result = $pingResult.ResponseMessage
                Write-LogMessage -Message $result -FilePath $script:logFile -NoConsole
                Write-Information -MessageData "[$currentTime] $result" -InformationAction Continue

                if ($DetailedOutput -and $pingResult.DetailedOutput) {
                    Write-LogMessage -Message "Detailed ping output:" -FilePath $script:logFile
                    Write-LogMessage -Message $pingResult.DetailedOutput -FilePath $script:logFile
                    Write-Information -MessageData "Detailed ping output:" -InformationAction Continue
                    Write-Information -MessageData $pingResult.DetailedOutput -InformationAction Continue
                }
            } else {
                $result = $pingResult.ResponseMessage
                Write-LogMessage -Message $result -FilePath $script:logFile -NoConsole
                Write-Information -MessageData "[$currentTime] $result" -InformationAction Continue

                if ($DetailedOutput -and $pingResult.DetailedOutput) {
                    Write-LogMessage -Message "Detailed ping output:" -FilePath $script:logFile
                    Write-LogMessage -Message $pingResult.DetailedOutput -FilePath $script:logFile
                    Write-Information -MessageData "Detailed ping output:" -InformationAction Continue
                    Write-Information -MessageData $pingResult.DetailedOutput -InformationAction Continue
                }
            }

            # Update statistics every 10 pings
            if ($script:sent % 10 -eq 0) {
                $packetLoss     = 100 - ($script:received / $script:sent * 100)
                $avgTime        = if ($script:received -gt 0) { $script:totalTime / $script:received } else { 0 }

                $stats = @"

Current Statistics:
----------------
Packets: Sent = $script:sent, Received = $script:received, Lost = $($script:sent - $script:received) ($($packetLoss.ToString('N2'))% loss)
Round Trip Times: Min = $(if($script:minTime -eq [int]::MaxValue) { "0"}else { $script:minTime})ms, Max = $script:maxTime`ms, Avg = $($avgTime.ToString('N2'))ms

"@
                Write-LogMessage -Message $stats -FilePath $script:logFile
                Write-Information -MessageData $stats -InformationAction Continue
            }

            # Check if we should stop
            if ($Count -gt 0 -and $script:sent -ge $Count) {
                break
            }

            # Small delay between pings
            Start-Sleep -Milliseconds 1000
        } catch {
            if ($_.Exception.Message -match "cancelled by the user") {
                $script:interrupted = $true
                break
            }
            throw
        }
    }
} catch {
    Write-Error "Error during ping test: $_"
    Write-Information -MessageData "Stack Trace: $($_.ScriptStackTrace)" -InformationAction Continue
    if ($script:logFile) {
        Write-LogMessage -Message "ERROR: $_" -FilePath $script:logFile
        Write-LogMessage -Message "Stack Trace: $($_.ScriptStackTrace)" -FilePath $script:logFile
    }
} finally {
    if ($script:interrupted) {
        Write-FinalStatistic -Interrupted
    } else {
        Write-FinalStatistic
    }
}
