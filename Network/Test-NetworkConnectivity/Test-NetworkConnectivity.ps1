# Script: Test-NetworkConnectivity.ps1
# Version: 2.11
# Description: Extended ping test with network configuration logging and continuous mode
# Author: jdyer-nuvodia
# Created: 2025-02-06 17:28:50

# Enable strict mode for better error handling
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

[CmdletBinding()]
param(
    [Parameter(Position=0)]
    [string]$Target = "8.8.8.8",
    
    [Parameter(Position=1)]
    [int]$Count = 0,  # 0 means continuous
    
    [Parameter()]
    [string]$OutputPath = "C:\PingLogs"  # Changed default path
)

# Initialize global variables
$script:logFile = $null
$script:sent = 0
$script:received = 0
$script:totalTime = 0
$script:minTime = [int]::MaxValue
$script:maxTime = 0
$script:interrupted = $false

# Function to handle cleanup and final statistics
function Write-FinalStatistics {
    param([switch]$Interrupted)
    
    if ($script:logFile) {
        try {
            $packetLoss = if ($script:sent -gt 0) { 100 - ($script:received / $script:sent * 100) } else { 0 }
            $avgTime = if ($script:received -gt 0) { $script:totalTime / $script:received } else { 0 }
            
            $finalStats = @"

========================================
Final Statistics $(if($Interrupted){"(Script Interrupted)"}):
========================================
Test Duration: $((Get-Date) - (Get-Item $script:logFile).CreationTime)
Packets: Sent = $script:sent, Received = $script:received, Lost = $($script:sent - $script:received) ($($packetLoss.ToString('N2'))% loss)
Round Trip Times: Min = $($script:minTime)ms, Max = $($script:maxTime)ms, Avg = $($avgTime.ToString('N2'))ms
========================================
Test completed$(if($Interrupted){" (Interrupted)"}): $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Log file size: $(Get-FormattedSize (Get-Item $script:logFile).Length)
========================================
"@
            # Force write the final statistics
            $finalStats | Out-File -FilePath $script:logFile -Append -Force
            
            Write-Host $finalStats -ForegroundColor $(if($Interrupted){"Yellow"}else{"Cyan"})
            
            # Add clear message about log file location
            Write-Host "`n==================================================" -ForegroundColor $(if($Interrupted){"Yellow"}else{"Green"})
            Write-Host "Log file has been saved:" -ForegroundColor $(if($Interrupted){"Yellow"}else{"Green"})
            Write-Host "Name: $(Split-Path $script:logFile -Leaf)" -ForegroundColor Yellow
            Write-Host "Location: $(Split-Path $script:logFile)" -ForegroundColor Yellow
            Write-Host "Full Path: $script:logFile" -ForegroundColor Yellow
            Write-Host "Size: $(Get-FormattedSize (Get-Item $script:logFile).Length)" -ForegroundColor Yellow
            Write-Host "==================================================" -ForegroundColor $(if($Interrupted){"Yellow"}else{"Green"})

            # Ensure file is flushed and closed properly
            $streamWriter = [System.IO.StreamWriter]::new($script:logFile, $true)
            $streamWriter.WriteLine($finalStats)
            $streamWriter.Close()
        }
        catch {
            Write-Host "Error writing final statistics: $_" -ForegroundColor Red
        }
    }
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
    
    # Write to file using StreamWriter for better control
    try {
        $streamWriter = [System.IO.StreamWriter]::new($FilePath, $true)
        $streamWriter.WriteLine($timestampedMessage)
        $streamWriter.Close()
    }
    catch {
        Write-Warning "Failed to write to log file: $_"
    }
    
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

# Create an event job to handle Ctrl+C
$job = Start-Job -ScriptBlock {
    $host.UI.RawUI.FlushInputBuffer()
    while ($true) {
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            if ($key.Key -eq [ConsoleKey]::C -and $key.Modifiers -eq [ConsoleModifiers]::Control) {
                return $true
            }
        }
        Start-Sleep -Milliseconds 100
    }
}

try {
    # Create output directory if it doesn't exist
    if (!(Test-Path -Path $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        Write-Host "Created output directory: $OutputPath" -ForegroundColor Yellow
    }

    # Create timestamp and filename
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $computerName = $env:COMPUTERNAME
    $fileName = "PingTest_${computerName}_${timestamp}.log"
    $script:logFile = Join-Path $OutputPath $fileName
    
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
    Set-Content -Path $script:logFile -Value $header -Force
    
    Write-Host "Starting network test - Results will be saved to: $script:logFile" -ForegroundColor Cyan
    Write-Host "Press Ctrl+C to stop continuous mode" -ForegroundColor Yellow
    
    # Get and log network configuration
    Write-LogMessage -Message "Getting network configuration..." -FilePath $script:logFile
    Write-LogMessage -Message "`nNETWORK CONFIGURATION:" -FilePath $script:logFile
    Write-LogMessage -Message "----------------------------------------" -FilePath $script:logFile
    
    $ipConfig = ipconfig /all
    $streamWriter = [System.IO.StreamWriter]::new($script:logFile, $true)
    $streamWriter.WriteLine($ipConfig)
    $streamWriter.Close()
    
    Write-LogMessage -Message "----------------------------------------`n" -FilePath $script:logFile
    
    # Start ping test
    Write-LogMessage -Message "Starting ping test to $Target..." -FilePath $script:logFile -ForegroundColor Cyan
    
    while (!$script:interrupted) {
        # Check for Ctrl+C
        if ($job.HasMoreData) {
            $script:interrupted = $true
            break
        }

        $pingResult = Test-Connection -ComputerName $Target -Count 1 -ErrorAction SilentlyContinue
        $script:sent++
        
        $currentTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        
        if ($pingResult) {
            $script:received++
            $responseTime = $pingResult.ResponseTime
            $script:totalTime += $responseTime
            $script:minTime = [Math]::Min($script:minTime, $responseTime)
            $script:maxTime = [Math]::Max($script:maxTime, $responseTime)
            
            $result = "Reply from $($pingResult.Address): time=${responseTime}ms size=$($pingResult.ReplySize)bytes"
            Write-LogMessage -Message $result -FilePath $script:logFile -NoConsole
            Write-Host "[$currentTime] $result" -ForegroundColor Green
        }
        else {
            $result = "Request timed out."
            Write-LogMessage -Message $result -FilePath $script:logFile -NoConsole
            Write-Host "[$currentTime] $result" -ForegroundColor Red
        }
        
        # Update statistics every 10 pings
        if ($script:sent % 10 -eq 0) {
            $packetLoss = 100 - ($script:received / $script:sent * 100)
            $avgTime = if ($script:received -gt 0) { $script:totalTime / $script:received } else { 0 }
            
            $stats = @"

Current Statistics:
----------------
Packets: Sent = $script:sent, Received = $script:received, Lost = $($script:sent - $script:received) ($($packetLoss.ToString('N2'))% loss)
Round Trip Times: Min = $($script:minTime)ms, Max = $($script:maxTime)ms, Avg = $($avgTime.ToString('N2'))ms

"@
            Write-LogMessage -Message $stats -FilePath $script:logFile
            Write-Host $stats -ForegroundColor Cyan
        }
        
        # Check if we should stop
        if ($Count -gt 0 -and $script:sent -ge $Count) {
            break
        }
        
        # Small delay between pings
        Start-Sleep -Milliseconds 1000
    }
}
catch {
    Write-Error "Error during ping test: $_"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    if ($script:logFile) {
        Write-LogMessage -Message "ERROR: $_" -FilePath $script:logFile
        Write-LogMessage -Message "Stack Trace: $($_.ScriptStackTrace)" -FilePath $script:logFile
    }
}
finally {
    # Clean up the background job
    Stop-Job -Job $job
    Remove-Job -Job $job
    
    # Write final statistics
    if ($script:interrupted) {
        Write-FinalStatistics -Interrupted
    }
    else {
        Write-FinalStatistics
    }
}