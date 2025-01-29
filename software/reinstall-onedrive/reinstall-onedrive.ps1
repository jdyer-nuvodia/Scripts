# Requires administrator privileges

# Check if running as administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Output "Script is not running as Administrator. Restarting with elevated privileges..."
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-File `"$($MyInvocation.MyCommand.Path)`""
    Exit
}

# Main script execution with timeout
$job = Start-Job -ScriptBlock {
    # Function to uninstall OneDrive
    function Uninstall-OneDrive {
        Write-Output "Stopping OneDrive processes..."
        Stop-Process -Name OneDrive* -Force -ErrorAction SilentlyContinue

        $oneDrivePaths = @(
            "$env:SystemRoot\System32\OneDriveSetup.exe",
            "$env:SystemRoot\SysWOW64\OneDriveSetup.exe",
            "${env:ProgramFiles}\Microsoft OneDrive\OneDriveSetup.exe",
            "${env:ProgramFiles(x86)}\Microsoft OneDrive\OneDriveSetup.exe"
        )

        foreach ($path in $oneDrivePaths) {
            if (Test-Path $path) {
                Write-Output "Uninstalling OneDrive from $path"
                Start-Process $path -ArgumentList "/uninstall" -Wait
            }
        }

        Write-Output "Uninstalling OneDrive using WinGet..."
        winget uninstall Microsoft.OneDrive
    }

    # Function to download and install the latest OneDrive
    function Install-LatestOneDrive {
        $url = "https://go.microsoft.com/fwlink/p/?LinkID=2182910"
        $outPath = "$env:TEMP\OneDriveSetup.exe"

        Write-Output "Downloading the latest OneDrive installer..."
        Invoke-WebRequest -Uri $url -OutFile $outPath

        Write-Output "Installing the latest version of OneDrive..."
        Start-Process $outPath -ArgumentList "/allusers" -Wait
    }

    # Execute the functions
    Write-Output "Uninstalling all versions of OneDrive..."
    Uninstall-OneDrive

    Write-Output "Installing the latest version of OneDrive..."
    Install-LatestOneDrive

    Write-Output "OneDrive update process completed."
}

$timeout = 300 # 5 minutes in seconds
$completed = Wait-Job $job -Timeout $timeout

if ($completed -eq $null) {
    Write-Output "The script did not complete within 10 minutes. Stopping the process..."
    Stop-Job $job
    Remove-Job $job
} else {
    Receive-Job $job
    Remove-Job $job
}
