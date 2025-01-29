# Set the target IP address or hostname
$target = "8.8.8.8"

# Set the output file name
$outputFile = "ping_results.txt"

# Get the current date and time
$startTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Write the header to the output file
"Ping results for $target starting at $startTimestamp" | Out-File -FilePath $outputFile
"" | Out-File -FilePath $outputFile -Append

Write-Host "Starting continuous ping to $target"
Write-Host "Results are being saved to $outputFile"
Write-Host "Press Ctrl+C to stop the script"

$pingCount = 0

while ($true) {
    $pingCount++
    $currentTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $pingOutput = ping -n 1 $target

    if ($pingOutput -match "Reply from") {
        $pingResult = ($pingOutput -match "Reply from.*" | Out-String).Trim()
        $line = "[$currentTimestamp] Ping ${pingCount}: $pingResult"
        Write-Host $line -ForegroundColor Green
        $line | Out-File -FilePath $outputFile -Append
    } else {
        $line = "[$currentTimestamp] Ping ${pingCount}: Request timed out."
        Write-Host $line -ForegroundColor Red
        $line | Out-File -FilePath $outputFile -Append
    }

    "" | Out-File -FilePath $outputFile -Append
    Start-Sleep -Seconds 1
}
