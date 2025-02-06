# Script: Test-NetworkConnectivity.ps1
# Version: 2.7
# Description: Extended ping test with network configuration logging and continuous mode
# Author: jdyer-nuvodia
# Created: 2025-02-05 23:59:45

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

    # Initialize global variables
    $global:logFile = $null
    $global:sent = 0
    $global:received = 0
    $global:totalTime = 0
    $global:minTime = [int]::MaxValue
    $global:maxTime = 0
    $global:interrupted = $false

    # Register event handler for Ctrl+C
    $null = Register-ObjectEvent -InputObject ([Console]) -EventName CancelKeyPress -Action {
        $global:interrupted = $true
        $event.MessageData = $true
        $event.Cancel = $true
    }

    # Function to handle cleanup and final statistics
    function Write-FinalStatistics {
        param([switch]$Interrupted)
        
        if ($global:logFile) {
            try {
                $packetLoss = if ($global:sent -gt 0) { 100 - ($global:received / $global:sent * 100) } else { 0 }
                $avgTime = if ($global:received -gt 0) { $global:totalTime / $global:received } else { 0 }
                
                $finalStats = @"

========================================
Final Statistics $(if($Interrupted){"(Script Interrupted)"}):
========================================
Test Duration: $((Get-Date) - (Get-Item $global:logFile).CreationTime)
Packets: Sent = $global:sent, Received = $global:received, Lost = $($global:sent - $global:received) ($($packetLoss.ToString('N2'))% loss)
Round Trip Times: Min = $($global:minTime)ms, Max = $($global:maxTime)ms, Avg = $($avgTime.ToString('N2'))ms
========================================
Test completed$(if($Interrupted){" (Interrupted)"}): $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Log file size: $(Get-FormattedSize (Get-Item $global:logFile).Length)
========================================
"@
                # Force write the final statistics
                $finalStats | Out-File -FilePath $global:logFile -Append -Force
                
                # Ensure file is flushed
                [System.IO.File]::WriteAllLines($global:logFile, (Get-Content $global:logFile))
                
                Write-Host $finalStats -ForegroundColor $(if($Interrupted){"Yellow"}else{"Cyan"})
                
                # Add clear message about log file location
                Write-Host "`n==================================================" -ForegroundColor $(if($Interrupted){"Yellow"}else{"Green"})
                Write-Host "Log file has been saved:" -ForegroundColor $(if($Interrupted){"Yellow"}else{"Green"})
                Write-Host "Name: $(Split-Path $global:logFile -Leaf)" -ForegroundColor Yellow
                Write-Host "Location: $(Split-Path $global:logFile)" -ForegroundColor Yellow
                Write-Host "Full Path: $global:logFile" -ForegroundColor Yellow
                Write-Host "Size: $(Get-FormattedSize (Get-Item $global:logFile).Length)" -ForegroundColor Yellow
                Write-Host "==================================================" -ForegroundColor $(if($Interrupted){"Yellow"}else{"Green"})
            }
            catch {
                Write-Host "Error writing final statistics: $_" -ForegroundColor Red
            }
            finally {
                # Ensure we flush any remaining content and close file handles
                try {
                    $fileStream = [System.IO.File]::Open($global:logFile, 'Open', 'Write', 'Read')
                    $fileStream.Close()
                }
                catch { }
            }
        }
    }

    # Rest of the script remains the same until the while loop...
    
    try {
        # [Previous try block content remains the same until the while loop]
        
        while (!$global:interrupted) {
            if ($global:interrupted) { break }
            
            # [Rest of the while loop content remains the same]
            
            # Check if we should stop
            if ($Count -gt 0 -and $global:sent -ge $Count) {
                break
            }
            
            # Small delay between pings
            Start-Sleep -Milliseconds 1000
        }

        # Write final statistics
        if ($global:interrupted) {
            Write-FinalStatistics -Interrupted
        }
        else {
            Write-FinalStatistics
        }
    }
    catch {
        Write-Error "Error during ping test: $_"
        Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
        if ($global:logFile) {
            Write-LogMessage -Message "ERROR: $_" -FilePath $global:logFile
            Write-LogMessage -Message "Stack Trace: $($_.ScriptStackTrace)" -FilePath $global:logFile
        }
    }
    finally {
        # Unregister the event handler
        Get-EventSubscriber | Where-Object {$_.SourceObject.ToString() -eq 'System.Console'} | Unregister-Event
    }
}

# Execute the script block with parameters
& $scriptBlock @args