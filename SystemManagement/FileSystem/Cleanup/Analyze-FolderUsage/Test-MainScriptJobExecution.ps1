# Test to investigate why the main script parallel jobs return 0 results
param([string]$TestPath = "C:\Temp")

$Script:CentralLogPath = ".\Test-MainScriptJobExecution_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"

function Write-CentralLog {
    param([string]$Message, [string]$Category = "TEST", [string]$Source = "MAIN")
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $logEntry = "[$timestamp] [$Source] [$Category] $Message"
    Add-Content -Path $Script:CentralLogPath -Value $logEntry -Encoding UTF8
    Write-Output $logEntry
}

Write-CentralLog -Message "=== MAIN SCRIPT JOB EXECUTION TEST ==="

# Use the EXACT scriptblock from the main script - just the Get-FolderStatistic part
$scriptBlock = {
    param($FolderPath, $MaxDepth, $EnableDebug, $CentralLogPath)
    
    # Exact same preferences as main script
    $ErrorActionPreference = 'Continue'
    $DebugPreference = 'SilentlyContinue'
    $VerbosePreference = 'SilentlyContinue'
    $InformationPreference = 'SilentlyContinue'
    $WarningPreference = 'SilentlyContinue'
    $ProgressPreference = 'SilentlyContinue'

    $DebugEnabled = $EnableDebug
    $script:CentralLogPath = $CentralLogPath
    
    function Write-RunspaceDebug {
        param([string]$Message, [string]$Category = "RUNSPACE")
        if ($DebugEnabled -and -not [string]::IsNullOrEmpty($script:CentralLogPath)) {
            try {
                $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
                $logEntry = "[$timestamp] [RUNSPACE:$Category] $Message"
                Add-Content -Path $script:CentralLogPath -Value $logEntry -Encoding UTF8 -ErrorAction SilentlyContinue
            } catch {}
        }
    }
    
    # Test - just return a simple PSCustomObject (same structure as main script)
    try {
        Write-RunspaceDebug -Message "Starting analysis of directory: $FolderPath" -Category "START"
        
        $result = [PSCustomObject]@{
            Path = $FolderPath
            SizeBytes = 12345
            FileCount = 10
            SubfolderCount = 3
            LargestFile = "testfile.txt"
            LargestFileSize = 6789
            IsAccessible = $true
            HasCloudFiles = $false
            Error = $null
            MaxDepthReached = $false
        }
        
        Write-RunspaceDebug -Message "Completed analysis of $FolderPath - returning PSCustomObject" -Category "COMPLETE"
        
        # Return the result - this should be a PSCustomObject
        return $result
        
    } catch {
        Write-RunspaceDebug -Message "ERROR in analysis of $FolderPath`: $($_.Exception.Message)" -Category "ERROR"
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

Write-CentralLog -Message "Testing with simplified scriptblock identical to main script structure"

# Test folders
$testFolders = @()
if (Test-Path $TestPath) {
    $testFolders = Get-ChildItem -Path $TestPath -Directory | Select-Object -First 3 | ForEach-Object { $_.FullName }
} else {
    $testFolders = @("C:\Windows\Temp", "C:\Users\Public")
}

Write-CentralLog -Message "Test folders: $($testFolders -join ', ')"

# Setup runspace pool exactly like main script
$maxThreads = 4
$runspacePool = [runspacefactory]::CreateRunspacePool(1, $maxThreads)
$runspacePool.Open()
$jobs = @()
$results = @()

try {
    # Launch jobs exactly like main script
    Write-CentralLog -Message "Launching parallel jobs..."
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
        Write-CentralLog -Message "Launched job for: $folder"
    }

    # Wait for jobs to complete - exactly like main script
    $completed = 0
    $timeout = (Get-Date).AddMinutes(1)  # Shorter timeout for test

    Write-CentralLog -Message "Waiting for jobs to complete..."

    while ($jobs.Count -gt $completed -and (Get-Date) -lt $timeout) {
        foreach ($job in $jobs) {
            if ($job.Handle.IsCompleted -and -not $job.Processed) {
                Write-CentralLog -Message "Processing completed job: $($job.Path)"
                try {
                    $jobResults = $job.PowerShell.EndInvoke($job.Handle)
                    Write-CentralLog -Message "Job $($job.Path) - Raw result type: $($jobResults.GetType().FullName)"
                    Write-CentralLog -Message "Job $($job.Path) - Raw result count: $($jobResults.Count)"
                    
                    # Convert to array like main script
                    $resultsArray = @($jobResults)
                    Write-CentralLog -Message "Job $($job.Path) - Array type: $($resultsArray.GetType().FullName)"
                    Write-CentralLog -Message "Job $($job.Path) - Array count: $($resultsArray.Count)"
                    
                    if ($resultsArray.Count -gt 0) {
                        foreach ($result in $resultsArray) {
                            Write-CentralLog -Message "Job $($job.Path) - Result type: $($result.GetType().FullName)"
                            if ($result -is [PSCustomObject]) {
                                $results += $result
                                Write-CentralLog -Message "✓ Collected PSCustomObject from $($job.Path): $($result.Path)"
                            } elseif ($result -is [String]) {
                                Write-CentralLog -Message "⚠ String result from $($job.Path): '$result'"
                            } else {
                                Write-CentralLog -Message "⚠ Other result type from $($job.Path): $($result.GetType().Name)"
                            }
                        }
                    } else {
                        Write-CentralLog -Message "❌ No results from job: $($job.Path)"
                    }
                    
                    $job.Processed = $true
                    $completed++
                    Write-CentralLog -Message "✓ Job $($job.Path) completed ($completed/$($jobs.Count))"
                      } catch {
                    Write-CentralLog -Message "❌ Error processing job $($job.Path): $($_.Exception.Message)"
                    $job.Processed = $true
                    $completed++
                } finally {
                    $job.PowerShell.Dispose()
                }
            }
        }
        Start-Sleep -Milliseconds 50
    }

    Write-CentralLog -Message "=== FINAL RESULTS ==="
    Write-CentralLog -Message "Total jobs launched: $($jobs.Count)"
    Write-CentralLog -Message "Jobs completed: $completed"
    Write-CentralLog -Message "Results collected: $($results.Count)"
    
    foreach ($result in $results) {
        Write-CentralLog -Message "  Result: $($result.Path) - Size: $($result.SizeBytes), Files: $($result.FileCount)"
    }
    
} catch {
    Write-CentralLog -Message "❌ CRITICAL ERROR: $($_.Exception.Message)"
} finally {
    if ($runspacePool) {
        $runspacePool.Close()
        $runspacePool.Dispose()
    }
}

Write-Output "Test completed. Log: $Script:CentralLogPath"
