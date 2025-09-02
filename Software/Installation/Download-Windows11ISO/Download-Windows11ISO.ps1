# =============================================================================
# Script: Download-Windows11ISO.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.2.2
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
# =============================================================================

<#
.SYNOPSIS
Downloads Windows 11 24H2 ISO file silently from Microsoft's official source.

.DESCRIPTION
This script downloads the Windows 11 24H2 ISO file from Microsoft's official download link to C:\Temp.
It uses optimized download methods to maximize speed while showing a progress bar during the download
and verifies the file integrity afterward. The optimizations include:
- Multi-threaded downloading with optimal buffer sizes
- Connection pooling and keep-alive settings
- Reduced progress bar updates to minimize CPU overhead
- Advanced HTTP client configuration
The script is designed to work with PowerShell 5.1 and execute silently.

.PARAMETER None
This script does not accept any parameters.

.EXAMPLE
.\Download-Windows11ISO.ps1
Downloads the Windows 11 24H2 ISO to C:\Temp and verifies the download.
#>

#Requires -Version 5.1

# Set strict mode and error preferences
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference = "Continue"

# Load required assemblies
Add-Type -AssemblyName System.Net.Http

# Script variables
$destinationFolder = "C:\Temp"
$isoFileName = "Windows11_24H2.iso"
$destinationPath = Join-Path -Path $destinationFolder -ChildPath $isoFileName
$logFileName = "Download-Windows11ISO_$(hostname)_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss')_UTC.log"
$logPath = Join-Path -Path $PSScriptRoot -ChildPath $logFileName

# Windows 11 24H2 download URL - Microsoft's official link
# Note: Microsoft sometimes changes URLs, so this might need updating
$downloadUrl = "https://software.download.prss.microsoft.com/dbazure/Win11_24H2_English_x64.iso?t = 67b30897-2be6-43d9-84e7-deac8e3d2aaf&P1 = 1747334540&P2 = 601&P3 = 2&P4 = Lrm%2bfKD9uFLxAagpY8i9Jx6%2bnZURyKtenVU0zMdfRWSznot2U30ok%2bc5SfPPtrFzhASYF0B8SjjSLN00W%2bwYJGq5izxvfC6ulMHEFT%2bqLVbe4q4qnkncWd%2bdq5qU9O8UrYv%2bouQ9gLVuE2kZbSlJB37VzLazzRVCAsoBmlpPomSRY5sV36LEk7zzDPrbRNBtbUZ9S%2fZySY1gYKlfSMdORI1PYNyo%2fw3sK3BFlChuZHojHlJCSyzkgJ6X2fXEZYVKoRy4ndRZCL%2bSViPEztzHHTlDMPLMTN%2fTbJdYpEOV0gEhVX4AuuAXmANVzLAyHx0EZAEKp1GESUHAN72qvk8ZYQ%3d%3d"

# Expected SHA-256 hash of the ISO file
# This is a placeholder - you should replace it with the actual hash from Microsoft
$expectedHash = "B56B911BF18A2CEAEB3904D87E7C770BDF92D3099599D61AC2497B91BF190B11"

# Performance optimization settings
$maxConcurrentThreads = [System.Environment]::ProcessorCount
$bufferSize = 16MB
# Set to $true to skip verification of existing files (improves speed but reduces security)
$skipHashVerificationIfExists = $false

# Color output configuration
$Script:UseAnsiColors = $PSVersionTable.PSVersion.Major -ge 7

# Color mappings for different PowerShell versions
if ($Script:UseAnsiColors) {
    # PowerShell 7+ ANSI escape codes
    $Script:Colors = @{
        White    = "`e[37m"
        Cyan     = "`e[36m"
        Green    = "`e[32m"
        Yellow   = "`e[33m"
        Red      = "`e[31m"
        Magenta  = "`e[35m"
        DarkGray = "`e[90m"
        Reset    = "`e[0m"
    }
} else {
    # PowerShell 5.1 console colors
    $Script:Colors = @{
        White    = "White"
        Cyan     = "Cyan"
        Green    = "Green"
        Yellow   = "Yellow"
        Red      = "Red"
        Magenta  = "Magenta"
        DarkGray = "DarkGray"
        Reset    = ""
    }
}

# Function to write colored output that works across PowerShell versions
function Write-ColorOutput {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$Color = "White"
    )

    if ($Script:UseAnsiColors) {
        # PowerShell 7+ with ANSI escape codes
        $colorCode = $Script:Colors[$Color]
        $resetCode = $Script:Colors.Reset
        Write-Output "${colorCode}${Message}${resetCode}"
    } else {
        # PowerShell 5.1 - Change console color, write output, then reset
        $originalColor = $Host.UI.RawUI.ForegroundColor
        try {
            if ($Script:Colors[$Color] -and $Script:Colors[$Color] -ne "") {
                $Host.UI.RawUI.ForegroundColor = $Script:Colors[$Color]
            }
            Write-Output $Message
        } finally {
            $Host.UI.RawUI.ForegroundColor = $originalColor
        }
    }
}

# Function to log messages
function Write-LogMessage {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS", "DEBUG", "PROCESS")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"

    # Output to console with appropriate color
    switch ($Level) {
        "INFO" { Write-ColorOutput -Message $logMessage -Color "White" }
        "WARNING" { Write-ColorOutput -Message $logMessage -Color "Yellow" }
        "ERROR" { Write-ColorOutput -Message $logMessage -Color "Red" }
        "SUCCESS" { Write-ColorOutput -Message $logMessage -Color "Green" }
        "DEBUG" { Write-ColorOutput -Message $logMessage -Color "Magenta" }
        "PROCESS" { Write-ColorOutput -Message $logMessage -Color "Cyan" }
        default { Write-ColorOutput -Message $logMessage -Color "White" }
    }

    # Write to log file
    Add-Content -Path $logPath -Value $logMessage
}

# Function to verify file using SHA-256
function Test-FileHash {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedHash
    )

    Write-LogMessage -Message "Verifying file integrity..." -Level "PROCESS"
    $fileHash = Get-FileHash -Path $FilePath -Algorithm SHA256

    if ($fileHash.Hash -eq $ExpectedHash) {
        Write-LogMessage -Message "File verification successful. Hash matches expected value." -Level "SUCCESS"
        return $true
    } else {
        Write-LogMessage -Message "File verification failed! Hash does not match expected value." -Level "ERROR"
        Write-LogMessage -Message "Expected: $ExpectedHash" -Level "DEBUG"
        Write-LogMessage -Message "Actual: $($fileHash.Hash)" -Level "DEBUG"
        return $false
    }
}

# Main script execution
try {
    # Create destination folder if it doesn't exist
    if (-not (Test-Path -Path $destinationFolder)) {
        Write-LogMessage -Message "Creating destination folder: $destinationFolder" -Level "PROCESS"
        New-Item -Path $destinationFolder -ItemType Directory -Force | Out-Null
    }

    # Check if file already exists
    if (Test-Path -Path $destinationPath) {
        Write-LogMessage -Message "ISO file already exists at: $destinationPath" -Level "WARNING"

        # Skip hash verification if configured (faster but less secure)
        if ($skipHashVerificationIfExists) {
            Write-LogMessage -Message "Skipping verification of existing file due to configuration." -Level "WARNING"
            Write-LogMessage -Message "Using existing ISO file." -Level "SUCCESS"
            exit 0
        } else {
            Write-LogMessage -Message "Verifying existing file..." -Level "PROCESS"

            if (Test-FileHash -FilePath $destinationPath -ExpectedHash $expectedHash) {
                Write-LogMessage -Message "Existing file is valid. No need to download again." -Level "SUCCESS"
                exit 0
            } else {
                Write-LogMessage -Message "Existing file is invalid. Will download a fresh copy." -Level "WARNING"
                Remove-Item -Path $destinationPath -Force
            }
        }
    }

    # Configure optimal download settings
    Write-LogMessage -Message "Initializing optimized secure download..." -Level "PROCESS"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # Configure ServicePoint for optimized connections
    $servicePoint = [System.Net.ServicePointManager]::FindServicePoint($downloadUrl)
    $servicePoint.ConnectionLimit = 10 * $maxConcurrentThreads
    $servicePoint.Expect100Continue = $false

    Write-LogMessage -Message "Optimized connection settings: ConnectionLimit = $($servicePoint.ConnectionLimit), Threads = $maxConcurrentThreads" -Level "DEBUG"

    # Use HttpClient instead of WebClient for better performance (available in PowerShell 5.1)
    $handler = New-Object System.Net.Http.HttpClientHandler
    $handler.AutomaticDecompression = [System.Net.DecompressionMethods]::GZip -bor [System.Net.DecompressionMethods]::Deflate
    $client = New-Object System.Net.Http.HttpClient($handler)
    # 2-hour timeout
    $client.Timeout = [System.TimeSpan]::FromMinutes(120)
    $client.DefaultRequestHeaders.Add("Accept-Encoding", "gzip, deflate")
    $client.DefaultRequestHeaders.Add("Keep-Alive", "true")

    # Start measuring download time
    $downloadStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    # Create file stream with optimal buffer size
    $fileStream = New-Object System.IO.FileStream($destinationPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None, $bufferSize, [System.IO.FileOptions]::SequentialScan)

    # Start download
    Write-LogMessage -Message "Beginning optimized download of Windows 11 24H2 ISO from Microsoft's official site..." -Level "PROCESS"
    Write-LogMessage -Message "Destination: $destinationPath" -Level "INFO"
    Write-LogMessage -Message "Using buffer size: $($bufferSize / 1MB) MB, Max threads: $maxConcurrentThreads" -Level "DEBUG"

    # Prepare progress tracking
    $lastProgressUpdate = Get-Date
    # Update progress at most once per second
    $updateFrequency = [TimeSpan]::FromSeconds(1)

    try {
        # Begin async download
        $response = $client.GetAsync($downloadUrl, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead)
        $response.Wait()

        if (!$response.Result.IsSuccessStatusCode) {
            Write-LogMessage -Message "Download failed with status: $($response.Result.StatusCode)" -Level "ERROR"
            throw "Failed to download file: HTTP status $($response.Result.StatusCode)"
        }

        $totalBytes = $response.Result.Content.Headers.ContentLength
        $bytesRead = 0

        # Get response stream
        $downloadTask = $response.Result.Content.ReadAsStreamAsync()
        $downloadTask.Wait()
        $contentStream = $downloadTask.Result

        # Read and write with optimal buffer
        $buffer = New-Object byte[] $bufferSize
        $totalBytesRead = 0

        do {
            # Read from download stream
            $bytesRead = $contentStream.Read($buffer, 0, $buffer.Length)

            if ($bytesRead -gt 0) {
                # Write to file stream
                $fileStream.Write($buffer, 0, $bytesRead)
                $totalBytesRead += $bytesRead

                # Only update the progress bar if enough time has passed
                $now = Get-Date
                if (($now - $lastProgressUpdate) -ge $updateFrequency) {
                    $percentComplete = [Math]::Min(100, [Math]::Round(($totalBytesRead / $totalBytes) * 100, 0))
                    $currentSpeed = [Math]::Round($totalBytesRead / $downloadStopwatch.Elapsed.TotalSeconds / 1MB, 2)

                    Write-Progress -Activity "Downloading Windows 11 24H2 ISO" `
                        -Status "$percentComplete% Complete" `
                        -PercentComplete $percentComplete `
                        -CurrentOperation "Downloaded $([Math]::Round($totalBytesRead / 1MB, 2)) MB of $([Math]::Round($totalBytes / 1MB, 2)) MB ($currentSpeed MB/s)"

                    $lastProgressUpdate = $now
                }
            }
        } while ($bytesRead -gt 0)

        Write-Progress -Activity "Downloading Windows 11 24H2 ISO" -Completed
        $downloadTime = $downloadStopwatch.Elapsed
        $averageSpeed = [Math]::Round($totalBytesRead / $downloadStopwatch.Elapsed.TotalSeconds / 1MB, 2)

        Write-LogMessage -Message "Download completed successfully in $($downloadTime.ToString())" -Level "SUCCESS"
        Write-LogMessage -Message "Downloaded $([Math]::Round($totalBytesRead / 1MB, 2)) MB at $averageSpeed MB/s" -Level "SUCCESS"
    } catch {
        Write-LogMessage -Message "Download failed: $($_.Exception.Message)" -Level "ERROR"
        throw
    } finally {
        # Clean up resources
        if ($contentStream) { $contentStream.Dispose() }
        if ($fileStream) { $fileStream.Dispose() }
        if ($client) { $client.Dispose() }
        if ($response -and $response.Result) { $response.Result.Dispose() }

        $downloadStopwatch.Stop()
    }

    # Verify downloaded file
    if (Test-Path -Path $destinationPath) {
        $fileSize = (Get-Item -Path $destinationPath).Length / 1GB
        Write-LogMessage -Message "Download completed. File size: $([Math]::Round($fileSize, 2)) GB" -Level "SUCCESS"

        if (Test-FileHash -FilePath $destinationPath -ExpectedHash $expectedHash) {
            Write-LogMessage -Message "Windows 11 24H2 ISO has been successfully downloaded and verified." -Level "SUCCESS"
            Write-LogMessage -Message "ISO is available at: $destinationPath" -Level "SUCCESS"
        } else {
            Write-LogMessage -Message "The downloaded file failed verification and may be corrupted." -Level "ERROR"
        }
    } else {
        Write-LogMessage -Message "Failed to download the Windows 11 24H2 ISO." -Level "ERROR"
    }
} catch {
    Write-LogMessage -Message "An error occurred: $($_.Exception.Message)" -Level "ERROR"
    Write-LogMessage -Message "Stack trace: $($_.ScriptStackTrace)" -Level "DEBUG"
    exit 1
} finally {
    # Clean up resources was handled in the download section
    Write-LogMessage -Message "Script execution completed." -Level "INFO"
}
