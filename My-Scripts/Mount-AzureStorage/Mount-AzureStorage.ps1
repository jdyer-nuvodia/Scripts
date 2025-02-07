$scriptContent = @'
# Enable verbose output
$VerbosePreference = "Continue"

# Start logging
$logFile = "C:\Users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowerShell\Scripts\1My Scripts\Mount-AzureStorage\AzureFileShareMount_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Start-Transcript -Path $logFile -Append

try {
    Write-Verbose "Script started at $(Get-Date)"

    # Specify the network interface to use
    $interfaceToUse = "Ethernet 4"
    Write-Verbose "Specified interface: $interfaceToUse"

    # Get all network adapters
    $allAdapters = Get-NetAdapter
    Write-Verbose "All network adapters:"
    $allAdapters | ForEach-Object { Write-Verbose "  $($_.Name) - $($_.Status)" }

    # Get the specified network interface
    $nic = Get-NetIPConfiguration | Where-Object { 
        $_.InterfaceAlias -eq $interfaceToUse -and 
        $_.IPv4DefaultGateway -ne $null -and 
        $_.NetAdapter.Status -eq "Up" 
    }

    if (-not $nic) {
        throw "Specified interface '$interfaceToUse' not found or not in a valid state."
    }

    Write-Verbose "Selected interface details:"
    Write-Verbose "  Name: $($nic.InterfaceAlias)"
    Write-Verbose "  IP Address: $($nic.IPv4Address.IPAddress)"
    Write-Verbose "  Default Gateway: $($nic.IPv4DefaultGateway.NextHop)"

    Write-Host "Testing connection on interface $($nic.InterfaceAlias)..."

    # Test the connection through the specified network interface
    $connectTestResult = Test-NetConnection -ComputerName jbteststorage0.file.core.windows.net -Port 445 -InterfaceAlias $nic.InterfaceAlias -WarningAction SilentlyContinue
    
    Write-Verbose "Connection test result:"
    Write-Verbose "  TCP test succeeded: $($connectTestResult.TcpTestSucceeded)"
    Write-Verbose "  Ping succeeded: $($connectTestResult.PingSucceeded)"
    Write-Verbose "  Name resolution succeeded: $($connectTestResult.NameResolutionSucceeded)"

    if ($connectTestResult.TcpTestSucceeded) {
        Write-Host "Connection to jbteststorage0.file.core.windows.net on port 445 succeeded through interface $($nic.InterfaceAlias)."

        # Save the password so the drive will persist on reboot
        $securePassword = ConvertTo-SecureString "sxmBaMKjIQJ9NKzWvnaf9fxXhScxN+9btj5xEAmEr8xKEmRBHwwhzzFgBRrLwrTckVp1H7c+IHOP+AStp1FnWg==" -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential ("AZURE\jbteststorage0", $securePassword)

        # Mount the drive
        $driveLetter = "Z"
        $rootPath = "\\jbteststorage0.file.core.windows.net\jbteststorage0"
        
        Write-Verbose "Attempting to mount drive $driveLetter"
        
        # Check if the drive is already mounted
        if (!(Test-Path "${driveLetter}:")) {
            New-PSDrive -Name $driveLetter -PSProvider FileSystem -Root $rootPath -Credential $credential -Persist -ErrorAction Stop
            Write-Host "Drive ${driveLetter}: has been successfully mounted."
        } else {
            Write-Host "Drive ${driveLetter}: is already mounted."
        }

        # Verify the mounted drive
        $mountedDrive = Get-PSDrive -Name $driveLetter -ErrorAction SilentlyContinue
        if ($mountedDrive) {
            Write-Verbose "Mounted drive details:"
            Write-Verbose "  Name: $($mountedDrive.Name)"
            Write-Verbose "  Root: $($mountedDrive.Root)"
            Write-Verbose "  Used space: $($mountedDrive.Used)"
            Write-Verbose "  Free space: $($mountedDrive.Free)"
        } else {
            Write-Warning "Drive $driveLetter is not visible as a PSDrive. It may not have mounted correctly."
        }
    } else {
        throw "Unable to reach the Azure storage account via port 445 on interface $($nic.InterfaceAlias). Check to make sure your organization or ISP is not blocking port 445, or use Azure P2S VPN, Azure S2S VPN, or Express Route to tunnel SMB traffic over a different port."
    }
}
catch {
    Write-Error "An error occurred: $_"
    Write-Verbose "Error details: $($_.Exception.Message)"
    Write-Verbose "Stack trace: $($_.ScriptStackTrace)"
}
finally {
    Write-Verbose "Script completed at $(Get-Date)"
    Stop-Transcript

    Write-Host "`nScript execution completed. Log file saved to: $logFile"
    Write-Host "Press Enter to exit..."
    Read-Host
}
'@

$encodedCommand = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($scriptContent))

Start-Process powershell.exe -ArgumentList "-NoExit", "-EncodedCommand", $encodedCommand
