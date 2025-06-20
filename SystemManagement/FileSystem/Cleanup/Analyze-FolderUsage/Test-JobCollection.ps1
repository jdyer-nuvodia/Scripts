# Test the exact job collection logic from the main script
param([string]$TestPath = "C:\Temp")

$Script:CentralLogPath = ".\Test-JobCollection_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"

function Write-CentralLog {
    param([string]$Message, [string]$Category = "TEST", [string]$Source = "MAIN")
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $logEntry = "[$timestamp] [$Source] [$Category] $Message"
    Add-Content -Path $Script:CentralLogPath -Value $logEntry -Encoding UTF8
    Write-Output $logEntry
}

Write-CentralLog -Message "Starting job collection test"

$testScriptBlock = {
    param($FolderPath)
    # Simple test that should return a PSCustomObject
    Start-Sleep -Milliseconds 100  # Simulate some work
    return [PSCustomObject]@{
        Path = $FolderPath
        TestResult = "Success"
        IsAccessible = $true
    }
}

$runspacePool = [runspacefactory]::CreateRunspacePool(1, 2)
$runspacePool.Open()
$jobs = @()
$results = @()

# Launch jobs
$testFolders = @("$TestPath\DattoRMM", "$TestPath\RAMMap")
foreach ($folder in $testFolders) {
    $powershell = [powershell]::Create()
    $powershell.RunspacePool = $runspacePool
    $powershell.AddScript($testScriptBlock).AddParameter("FolderPath", $folder) | Out-Null
    $jobs += [PSCustomObject]@{
        PowerShell = $powershell
        Handle = $powershell.BeginInvoke()
        Path = $folder
        StartTime = Get-Date
        Processed = $false
    }
    Write-CentralLog -Message "Launched job for: $folder"
}

Write-CentralLog -Message "Starting job collection loop"

# EXACT job collection logic from main script
$completed = 0
$timeout = (Get-Date).AddMinutes(1)  # Shorter timeout for test
$lastProgress = Get-Date
$progressCheckInterval = 500
$individualJobTimeout = 30

while ($jobs.Count -gt $completed -and (Get-Date) -lt $timeout) {
    Write-CentralLog -Message "Loop iteration: completed=$completed, total=$($jobs.Count)"
    
    foreach ($job in $jobs) {
        Write-CentralLog -Message "Checking job $($job.Path): IsCompleted=$($job.Handle.IsCompleted), Processed=$($job.Processed)"
        
        if ($job.Handle.IsCompleted -and -not $job.Processed) {
            Write-CentralLog -Message "Processing completed job: $($job.Path)"
            try {
                $jobResults = $job.PowerShell.EndInvoke($job.Handle)
                Write-CentralLog -Message "Raw result type: $($jobResults.GetType().FullName)"
                
                # Process results with better validation - handle PSDataCollection properly
                $resultsArray = @($jobResults)  # Convert PSDataCollection to array
                Write-CentralLog -Message "Converted array count: $($resultsArray.Count)"
                
                if ($resultsArray.Count -gt 0) {
                    Write-CentralLog -Message "Job $($job.Path) returned $($resultsArray.Count) result(s)"
                    foreach ($result in $resultsArray) {
                        Write-CentralLog -Message "Result type: $($result.GetType().FullName), IsPSCustomObject: $($result -is [PSCustomObject])"
                        if ($result -is [PSCustomObject] -and $result.PSObject.Properties['Path']) {
                            $results += $result
                            Write-CentralLog -Message "Collected valid result for: $($result.Path)"
                        } else {
                            Write-CentralLog -Message "WARNING: Invalid result: $result"
                        }
                    }
                } else {
                    Write-CentralLog -Message "Job $($job.Path) returned no results"
                }
                $job.Processed = $true
                $completed++
                Write-CentralLog -Message "Job completed: $($job.Path) (Total: $completed/$($jobs.Count))"
            } catch {
                Write-CentralLog -Message "ERROR processing job $($job.Path): $($_.Exception.Message)"
                $job.Processed = $true
                $completed++
            } finally {
                $job.PowerShell.Dispose()
            }
        }
    }
    
    Start-Sleep -Milliseconds 100
}

Write-CentralLog -Message "Job collection completed. Collected $($results.Count) results"
Write-CentralLog -Message "Results summary:"
foreach ($result in $results) {
    Write-CentralLog -Message "  - $($result.Path): $($result.TestResult)"
}

$runspacePool.Close()
$runspacePool.Dispose()

Write-Output "Test completed. Check log: $Script:CentralLogPath"
