# EXACT mimic of main script job execution to find the root cause
param([string]$TestPath = "C:\Temp")

$Script:CentralLogPath = ".\Test-ExactMainScript_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"

function Write-CentralLog {
    param([string]$Message, [string]$Category = "TEST", [string]$Source = "MAIN")
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $logEntry = "[$timestamp] [$Source] [$Category] $Message"
    Add-Content -Path $Script:CentralLogPath -Value $logEntry -Encoding UTF8
    Write-Output $logEntry
}

Write-CentralLog -Message "=== EXACT MAIN SCRIPT MIMIC TEST ==="

# Use EXACT same scriptblock structure as main script
$scriptBlock = {
    param($FolderPath, $MaxDepth, $EnableDebug, $CentralLogPath)
    
    # Strict output control - capture all unwanted output
    $ErrorActionPreference = 'Continue'
    $DebugPreference = 'SilentlyContinue'
    $VerbosePreference = 'SilentlyContinue'
    $InformationPreference = 'SilentlyContinue'
    $WarningPreference = 'SilentlyContinue'
    $ProgressPreference = 'SilentlyContinue'

    # Debug flag passed from main script
    $DebugEnabled = $EnableDebug
    # Make CentralLogPath available to nested functions
    $script:CentralLogPath = $CentralLogPath
    
    # Debug output function for runspace
    function Write-RunspaceDebug {
        param([string]$Message, [string]$Category = "RUNSPACE")
        if ($DebugEnabled -and -not [string]::IsNullOrEmpty($script:CentralLogPath)) {
            try {
                $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
                $logEntry = "[$timestamp] [RUNSPACE:$Category] $Message"
                Add-Content -Path $script:CentralLogPath -Value $logEntry -Encoding UTF8 -ErrorAction SilentlyContinue
            } catch {
                Write-Error -Message "Failed to write to central log: $($_.Exception.Message)" -ErrorAction SilentlyContinue
            }
        }
    }
    
    # Simplified test - just return a basic PSCustomObject
    try {
        Write-RunspaceDebug -Message "Starting analysis of root directory: $FolderPath" -Category "ROOT_START"
        
        $result = [PSCustomObject]@{
            Path = $FolderPath
            SizeBytes = 1000
            FileCount = 5
            SubfolderCount = 2
            LargestFile = "test.txt"
            LargestFileSize = 500
            IsAccessible = $true
            HasCloudFiles = $false
            Error = $null
            MaxDepthReached = $false
        }
        
        Write-RunspaceDebug -Message "Completed analysis of $FolderPath - returning result" -Category "ROOT_COMPLETE"
        return $result
    } catch {
        Write-RunspaceDebug -Message "CRITICAL ERROR in root analysis of $FolderPath`: $($_.Exception.Message)" -Category "ROOT_ERROR"
        return [PSCustomObject]@{
            Path = $FolderPath
            SizeBytes = 0
            FileCount = 0
            SubfolderCount = 0
            LargestFile = $null
            LargestFileSize = 0
            IsAccessible = $false
            HasCloudFiles = $false
            Error = $_.Exception.Message
            MaxDepthReached = $false
        }
    }
}

# EXACT same runspace setup as main script
$maxThreads = 4
$runspacePool = [runspacefactory]::CreateRunspacePool(1, $maxThreads)
$runspacePool.Open()
$jobs = @()
$results = @()

Write-CentralLog -Message "Creating runspace pool with $maxThreads threads"

# Test with actual directories from C:\Temp
$testFolders = Get-ChildItem -Path $TestPath -Directory | Select-Object -First 4 | ForEach-Object { $_.FullName }
Write-CentralLog -Message "Test folders: $($testFolders -join ', ')"

try {
    # Launch parallel jobs - EXACT same as main script
    Write-CentralLog -Message "Starting parallel job execution for $($testFolders.Count) directories"
    foreach ($folder in $testFolders) {
        $powershell = [powershell]::Create()
        $powershell.RunspacePool = $runspacePool
        $powershell.AddScript($scriptBlock).AddParameter("FolderPath", $folder).AddParameter("MaxDepth", 2).AddParameter("EnableDebug", $true).AddParameter("CentralLogPath", $Script:CentralLogPath) | Out-Null
        
        $jobs += [PSCustomObject]@{
            PowerShell = $powershell
            Handle = $powershell.BeginInvoke()
            Path = $folder
            StartTime = Get-Date
            Processed = $false
        }
        Write-CentralLog -Message "Launched job for directory: $folder"
    }

    # EXACT same job collection loop as main script
    $completed = 0
    $timeout = (Get-Date).AddMinutes(2)  # Shorter for test
    $lastProgress = Get-Date
    $progressCheckInterval = 500
    $individualJobTimeout = 60

    Write-CentralLog -Message "Starting job collection loop..."

    while ($jobs.Count -gt $completed -and (Get-Date) -lt $timeout) {
        Write-CentralLog -Message "Loop iteration: completed=$completed, total=$($jobs.Count), time remaining=$((($timeout - (Get-Date)).TotalSeconds)) seconds"
        
        foreach ($job in $jobs) {
            Write-CentralLog -Message "Checking job $($job.Path): IsCompleted=$($job.Handle.IsCompleted), Processed=$($job.Processed)"
            
            if ($job.Handle.IsCompleted -and -not $job.Processed) {
                Write-CentralLog -Message "Processing completed job: $($job.Path)"
                try {
                    $jobResults = $job.PowerShell.EndInvoke($job.Handle)
                    Write-CentralLog -Message "Raw jobResults type: $($jobResults.GetType().FullName)"
                    Write-CentralLog -Message "Raw jobResults count: $($jobResults.Count)"
                    
                    # Process results with better validation - handle PSDataCollection properly
                    $resultsArray = @($jobResults)  # Convert PSDataCollection to array
                    Write-CentralLog -Message "Converted resultsArray type: $($resultsArray.GetType().FullName)"
                    Write-CentralLog -Message "Converted resultsArray count: $($resultsArray.Count)"
                    
                    if ($resultsArray.Count -gt 0) {
                        Write-CentralLog -Message "Job $($job.Path) returned $($resultsArray.Count) result(s)"
                        foreach ($result in $resultsArray) {
                            Write-CentralLog -Message "Processing result: Type=$($result.GetType().FullName), IsPSCustomObject=$($result -is [PSCustomObject])"
                            if ($result -is [PSCustomObject] -and $result.PSObject.Properties['Path']) {
                                $results += $result
                                Write-CentralLog -Message "✓ Collected valid result for: $($result.Path)"
                            } elseif ($result -is [String]) {
                                Write-CentralLog -Message "⚠ WARNING: String result from job $($job.Path): '$result'"
                            } else {
                                Write-CentralLog -Message "⚠ WARNING: Invalid result type received from job $($job.Path): $($result.GetType().Name)"
                            }
                        }
                    } else {
                        Write-CentralLog -Message "Job $($job.Path) returned no results"
                    }
                    $job.Processed = $true
                    $completed++
                    Write-CentralLog -Message "✓ Job completed for: $($job.Path) (Total completed: $completed/$($jobs.Count))"
                } catch {
                    Write-CentralLog -Message "❌ ERROR: Failed to collect results from job $($job.Path): $($_.Exception.Message)"
                    Write-CentralLog -Message "Stack trace: $($_.ScriptStackTrace)"
                    $job.Processed = $true
                    $completed++
                } finally {
                    $job.PowerShell.Dispose()
                }
            }
            
            # Check for individual job timeout
            if (-not $job.Handle.IsCompleted -and -not $job.Processed) {
                $elapsed = (Get-Date).Subtract($job.StartTime).TotalSeconds
                if ($elapsed -gt $individualJobTimeout) {
                    Write-CentralLog -Message "⚠ WARNING: Job timeout for $($job.Path) after $elapsed seconds"
                    try {
                        $job.PowerShell.Stop()
                        $job.PowerShell.Dispose()
                    } catch {
                        Write-CentralLog -Message "Error stopping timed-out job: $($_.Exception.Message)"
                    }
                    $job.Processed = $true
                    $completed++
                }
            }
        }
        
        Start-Sleep -Milliseconds 100
    }

    # Handle any remaining incomplete jobs
    $incompleteJobs = $jobs | Where-Object { -not $_.Processed }
    if ($incompleteJobs.Count -gt 0) {
        Write-CentralLog -Message "⚠ WARNING: $($incompleteJobs.Count) jobs did not complete within timeout"
        foreach ($incompleteJob in $incompleteJobs) {
            Write-CentralLog -Message "Cleaning up incomplete job: $($incompleteJob.Path)"
            try {
                $incompleteJob.PowerShell.Stop()
                $incompleteJob.PowerShell.Dispose()
            } catch {
                Write-CentralLog -Message "Error cleaning up incomplete job: $($_.Exception.Message)"
            }
        }
    }

    Write-CentralLog -Message "✓ Parallel execution completed. Collected $($results.Count) total results"
    
    # Show results summary
    Write-CentralLog -Message "=== RESULTS SUMMARY ==="
    Write-CentralLog -Message "Total results collected: $($results.Count)"
    foreach ($result in $results) {
        Write-CentralLog -Message "  - $($result.Path): Size=$($result.SizeBytes), Files=$($result.FileCount), Accessible=$($result.IsAccessible)"
    }
    
} catch {
    Write-CentralLog -Message "❌ CRITICAL ERROR in test: $($_.Exception.Message)"
    Write-CentralLog -Message "Stack trace: $($_.ScriptStackTrace)"
} finally {
    if ($runspacePool) {
        $runspacePool.Close()
        $runspacePool.Dispose()
    }
}

Write-Output "Test completed. Check full log: $Script:CentralLogPath"
