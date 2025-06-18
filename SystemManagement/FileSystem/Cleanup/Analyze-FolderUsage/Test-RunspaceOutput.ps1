# Simple test script to isolate the runspace output issue
param(
    [string]$TestPath = "C:\Windows\System32\drivers"
)

$Script:Colors = @{
    Reset = "`e[0m"
    Green = "`e[32m"
    Red = "`e[31m"
    Yellow = "`e[33m"
}

Write-Output "$($Script:Colors.Green)Testing runspace output isolation...$($Script:Colors.Reset)"

$runspacePool = [runspacefactory]::CreateRunspacePool(1, 2)
$runspacePool.Open()

$scriptBlock = {
    param($FolderPath)
    
    # Test what gets output
    $testObject = [PSCustomObject]@{
        Path = $FolderPath
        TestProperty = "ValidObject"
        IsAccessible = $true
    }
    
    # CRITICAL: Only use Write-Output for the return value
    Write-Output $testObject
    return
}

$jobs = @()
$testFolders = @("C:\Windows\System32\drivers", "C:\Windows\System32\config")

foreach ($folder in $testFolders) {
    $powershell = [powershell]::Create()
    $powershell.RunspacePool = $runspacePool
    $powershell.AddScript($scriptBlock).AddParameter("FolderPath", $folder) | Out-Null
    $jobs += [PSCustomObject]@{
        PowerShell = $powershell
        Handle = $powershell.BeginInvoke()
        Path = $folder
    }
}

# Wait for completion
while ($jobs | Where-Object { -not $_.Handle.IsCompleted }) {
    Start-Sleep -Milliseconds 100
}

Write-Output "$($Script:Colors.Yellow)Collecting results...$($Script:Colors.Reset)"

foreach ($job in $jobs) {
    if ($job.Handle.IsCompleted) {
        # BREAKPOINT: Set breakpoint here to see what EndInvoke returns
        $rawResult = $job.PowerShell.EndInvoke($job.Handle)
        
        Write-Output "Job for '$($job.Path)':"
        Write-Output "  Raw result type: $($rawResult.GetType().Name)"
        Write-Output "  Raw result is array: $($rawResult -is [array])"
        
        # CRITICAL FIX: Extract actual results from PSDataCollection
        $result = @($rawResult)  # Convert PSDataCollection to array
        
        Write-Output "  Converted result type: $($result.GetType().Name)"
        Write-Output "  Converted result is array: $($result -is [array])"
        
        if ($result -is [array]) {
            Write-Output "  Array count: $($result.Count)"
            for ($i = 0; $i -lt $result.Count; $i++) {
                $item = $result[$i]
                Write-Output "    Element ${i}: Type=$($item.GetType().Name), IsPSCustomObject=$($item -is [PSCustomObject])"
                if ($item -is [string]) {
                    Write-Output "      String content: $item"
                } elseif ($item -is [PSCustomObject]) {
                    Write-Output "      PSCustomObject properties: $($item.PSObject.Properties.Name -join ', ')"
                }
            }
        } else {
            Write-Output "  Single result: Type=$($result.GetType().Name), IsPSCustomObject=$($result -is [PSCustomObject])"
            if ($result -is [string]) {
                Write-Output "    String content: $result"
            } elseif ($result -is [PSCustomObject]) {
                Write-Output "    PSCustomObject properties: $($result.PSObject.Properties.Name -join ', ')"
            }
        }
        Write-Output ""
    }
    
    $job.PowerShell.Dispose()
}

$runspacePool.Close()
$runspacePool.Dispose()

Write-Output "$($Script:Colors.Green)Test completed.$($Script:Colors.Reset)"
