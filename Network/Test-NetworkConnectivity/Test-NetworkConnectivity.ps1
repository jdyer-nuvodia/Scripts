# Script: Test-NetworkConnectivity.ps1
# Version: 2.6
# Description: Extended ping test with network configuration logging and continuous mode
# Author: jdyer-nuvodia
# Created: 2025-02-05 23:55:27

# Use script block to contain all code
$scriptBlock = {
    [CmdletBinding()]
    param(
        [Parameter(Position=0)]
        [string]$Target = "8.8.8.8",
        
        [Parameter(Position=1)]
        [int]$Count = 0,  # 0 means continuous
        
        [Parameter()]
        [string]$OutputPath = "C:\PingLogs"  # Changed default path
    )

    # Initialize global variables for trap access
    $global:logFile = $null
    $global:sent = 0
    $global:received = 0
    $global:totalTime = 0
    $global:minTime = [int]::MaxValue
    $global:maxTime = 0
    $global:interrupted = $false

    # Trap Ctrl+C and ensure graceful exit
    trap [System.Management.Automation.PipelineStoppedException] {
        $global:interrupted = $true
        Write-Host "`nScript interrupted by user. Writing final statistics..." -ForegroundColor Yellow
        
        if ($global:logFile) {
            try {
                # Calculate final statistics
                $packetLoss = if ($global:sent -gt 0) { 100 - ($global:received / $global:sent * 100) } else { 0 }
                $avgTime = if ($global:received -gt 0) { $global:totalTime / $global:received } else { 0 }
                
                $finalStats = @"

========================================
Final Statistics (Script Interrupted):
========================================
Test Duration: $((Get-Date) - (Get-Item $global:logFile).CreationTime)
Packets: Sent = $global:sent, Received = $global:received, Lost = $($global:sent - $global:received) ($($packetLoss.ToString('N2'))% loss)
Round Trip Times: Min = $($global:minTime)ms, Max = $($global:maxTime)ms, Avg = $($avgTime.ToString('N2'))ms
========================================
Test completed (Interrupted): $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Log file size: $(Get-FormattedSize (Get-Item $global:logFile).Length)
========================================
"@
                # Force the statistics to be written to the log file
                $finalStats | Out-File -FilePath $global:logFile -Append -Force
                [System.IO.File]::WriteAllLines($global:logFile, (Get-Content $global:logFile))

                # Add clear message about log file location
                Write-Host "`n==================================================" -ForegroundColor Yellow
                Write-Host "Log file has been saved:" -ForegroundColor Yellow
                Write-Host "Name: $(Split-Path $global:logFile -Leaf)" -ForegroundColor Yellow
                Write-Host "Location: $(Split-Path $global:logFile)" -ForegroundColor Yellow
                Write-Host "Full Path: $global:logFile" -ForegroundColor Yellow
                Write-Host "Size: $(Get-FormattedSize (Get-Item $global:logFile).Length)" -ForegroundColor Yellow
                Write-Host "==================================================" -ForegroundColor Yellow
            }
            catch {
                Write-Host "Error writing final statistics: $_" -ForegroundColor Red
            }
            finally {
                # Ensure we flush any remaining content and close file handles
                if ($global:logFile) {
                    $null = [System.IO.File]::OpenWrite($global:logFile).Close()
                }
            }
        }
        # Exit the script cleanly
        exit
    }

    function Write-LogMessage {
        param(
            [string]$Message,
            [string]$FilePath,
            [switch]$NoConsole,
            [string]$ForegroundColor = 'White'
        )
        
        # Add timestamp to message
        $timestampedMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message"
        
        # Write to file
        Add-Content -Path $FilePath -Value $timestampedMessage
        
        # Write to console if not suppressed
        if (!$NoConsole) {
            Write-Host $timestampedMessage -ForegroundColor $ForegroundColor
        }
    }

    function Get-FormattedSize {
        param([int64]$Size)
        
        if ($Size -gt 1GB) { return "{0:N2} GB" -f ($Size / 1GB) }
        if ($Size -gt 1MB) { return "{0:N2} MB" -f ($Size / 1MB) }
        if ($Size -gt 1KB) { return "{0:N2} KB" -f ($Size / 1KB) }
        return "$Size Bytes"
    }

    try {
        # Create output directory if it doesn't exist
        if (!(Test-Path -Path $OutputPath)) {
            try {
                New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
                Write-Host "Created output directory: $OutputPath" -ForegroundColor Yellow
            }
            catch {
                Write-Error "Failed to create output directory: $_"
                exit 1
            }
        }

        # Create timestamp and filename
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $computerName = $env:COMPUTERNAME
        $fileName = "PingTest_${computerName}_${timestamp}.log"
        $global:logFile = Join-Path $OutputPath $fileName
        
        # Create log file with header
        $header = @"
========================================
Extended Ping Test Results
========================================
Computer Name: $computerName
Test Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Target: $Target
Mode: $(if($Count -eq 0){"Continuous"}else{"Count: $Count"})
========================================

"@
        Set-Content -Path $global:logFile -Value $header
        
        Write-Host "Starting network test - Results will be saved to: $global:logFile" -ForegroundColor Cyan
        Write-Host "Press Ctrl+C to stop continuous mode" -ForegroundColor Yellow
        
        # Get and log network configuration
        Write-LogMessage -Message "Getting network configuration..." -FilePath $global:logFile
        Write-LogMessage -Message "`nNETWORK CONFIGURATION:" -FilePath $global:logFile
        Write-LogMessage -Message "----------------------------------------" -FilePath $global:logFile
        
        $ipConfig = ipconfig /all
        Add-Content -Path $global:logFile -Value $ipConfig
        Write-LogMessage -Message "----------------------------------------`n" -FilePath $global:logFile
        
        # Start ping test
        Write-LogMessage -Message "Starting ping test to $Target..." -FilePath $global:logFile -ForegroundColor Cyan
        
        while (!$global:interrupted) {
            $pingResult = Test-Connection -ComputerName $Target -Count 1 -ErrorAction SilentlyContinue
            $global:sent++
            
            $currentTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            
            if ($pingResult) {
                $global:received++
                $responseTime = $pingResult.ResponseTime
                $global:totalTime += $responseTime
                $global:minTime = [Math]::Min($global:minTime, $responseTime)
                $global:maxTime = [Math]::Max($global:maxTime, $responseTime)
                
                $result = "Reply from $($pingResult.Address): time=${responseTime}ms size=$($pingResult.ReplySize)bytes"
                Write-LogMessage -Message $result -FilePath $global:logFile -NoConsole
                Write-Host "[$currentTime] $result" -ForegroundColor Green
            }
            else {
                $result = "Request timed out."
                Write-LogMessage -Message $result -FilePath $global:logFile -NoConsole
                Write-Host "[$currentTime] $result" -ForegroundColor Red
            }
            
            # Update statistics every 10 pings
            if ($global:sent % 10 -eq 0) {
                $packetLoss = 100 - ($global:received / $global:sent * 100)
                $avgTime = if ($global:received -gt 0) { $global:totalTime / $global:received } else { 0 }
                
                $stats = @"

Current Statistics:
----------------
Packets: Sent = $global:sent, Received = $global:received, Lost = $($global:sent - $global:received) ($($packetLoss.ToString('N2'))% loss)
Round Trip Times: Min = $($global:minTime)ms, Max = $($global:maxTime)ms, Avg = $($avgTime.ToString('N2'))ms

"@
                Write-LogMessage -Message $stats -FilePath $global:logFile
                Write-Host $stats -ForegroundColor Cyan
            }
            
            # Check if we should stop
            if ($Count -gt 0 -and $global:sent -ge $Count) {
                break
            }
            
            # Small delay between pings
            Start-Sleep -Milliseconds 1000
        }

        # Log final statistics
        $packetLoss = if ($global:sent -gt 0) { 100 - ($global:received / $global:sent * 100) } else { 0 }
        $avgTime = if ($global:received -gt 0) { $global:totalTime / $global:received } else { 0 }
        
        $finalStats = @"

========================================
Final Statistics:
========================================
Test Duration: $((Get-Date) - (Get-Item $global:logFile).CreationTime)
Packets: Sent = $global:sent, Received = $global:received, Lost = $($global:sent - $global:received) ($($packetLoss.ToString('N2'))% loss)
Round Trip Times: Min = $($global:minTime)ms, Max = $($global:maxTime)ms, Avg = $($avgTime.ToString('N2'))ms
========================================
Test completed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Log file size: $(Get-FormattedSize (Get-Item $global:logFile).Length)
========================================
"@
        Add-Content -Path $global:logFile -Value $finalStats -Force
        Write-Host $finalStats -ForegroundColor Cyan

        # Add clear message about log file location
        Write-Host "`n==================================================" -ForegroundColor Green
        Write-Host "Log file has been created:" -ForegroundColor Green
        Write-Host "Name: $(Split-Path $global:logFile -Leaf)" -ForegroundColor Yellow
        Write-Host "Location: $(Split-Path $global:logFile)" -ForegroundColor Yellow
        Write-Host "Full Path: $global:logFile" -ForegroundColor Yellow
        Write-Host "Size: $(Get-FormattedSize (Get-Item $global:logFile).Length)" -ForegroundColor Yellow
        Write-Host "==================================================" -ForegroundColor Green
    }
    catch {
        Write-Error "Error during ping test: $_"
        Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
        if ($global:logFile) {
            Write-LogMessage -Message "ERROR: $_" -FilePath $global:logFile
            Write-LogMessage -Message "Stack Trace: $($_.ScriptStackTrace)" -FilePath $global:logFile
        }
    }
}

# Execute the script block with parameters
& $scriptBlock @args