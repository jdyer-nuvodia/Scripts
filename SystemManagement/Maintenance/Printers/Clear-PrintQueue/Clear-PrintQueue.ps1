# =============================================================================
# Script: Clear-PrintQueue.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.2
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
# =============================================================================

<#
.SYNOPSIS
Clears all jobs from the Windows print queue and restarts the Print Spooler service.

.DESCRIPTION
This script retrieves all current print jobs using the Win32_PrintJob WMI/CIM class and removes them.
It then stops and starts the Print Spooler (spooler) service.
Requires administrative privileges to run.
Includes -WhatIf support to show what actions would be taken without actually performing them.

.EXAMPLE
PS C:\> .\Clear-PrintQueue.ps1
Attempting to clear print queue and restart Print Spooler service...
Print queue is already clear.
Stopping Print Spooler service (spooler)...
Print Spooler service stopped.
Starting Print Spooler service (spooler)...
Print Spooler service started successfully.
Print queue cleared and Print Spooler service restarted successfully.
[Description: Clears the print queue and restarts the spooler service.]

.EXAMPLE
PS C:\> .\Clear-PrintQueue.ps1 -WhatIf
What if: Performing the operation "Remove" on target "Print Job ID: 1, Document: Microsoft Word - Document1".
What if: Performing the operation "Stop" on target "Service: spooler".
What if: Performing the operation "Start" on target "Service: spooler".
[Description: Shows which print jobs would be removed and indicates that the spooler service would be stopped and started, but does not perform these actions.]

.NOTES
Requires running as Administrator.
Uses Get-CimInstance and standard service cmdlets.
Ensure you have the necessary permissions to manage print jobs and services.
#>

#Requires -RunAsAdministrator

[CmdletBinding(SupportsShouldProcess = $true)]
param()

begin {
    # Color support variables and Write-ColorOutput function
    $Script:UseAnsiColors = $PSVersionTable.PSVersion.Major -ge 7
    $Script:Colors = if ($Script:UseAnsiColors) {
        @{
            'White' = "`e[37m"
            'Cyan' = "`e[36m"
            'Green' = "`e[32m"
            'Yellow' = "`e[33m"
            'Red' = "`e[31m"
            'Magenta' = "`e[35m"
            'DarkGray' = "`e[90m"
            'Reset' = "`e[0m"
        }
    } else {
        @{
            'White' = [ConsoleColor]::White
            'Cyan' = [ConsoleColor]::Cyan
            'Green' = [ConsoleColor]::Green
            'Yellow' = [ConsoleColor]::Yellow
            'Red' = [ConsoleColor]::Red
            'Magenta' = [ConsoleColor]::Magenta
            'DarkGray' = [ConsoleColor]::DarkGray
            'Reset' = ''
        }
    }

    function Write-ColorOutput {
        <#
        .SYNOPSIS
        Outputs colored text in a way that's compatible with PSScriptAnalyzer requirements.

        .DESCRIPTION
        This function provides colored output while maintaining compatibility with PSScriptAnalyzer
        by using only Write-Output and standard PowerShell cmdlets. Supports both PowerShell 5.1
        and PowerShell 7+ with appropriate color handling for each version.
        #>
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
}

process {
    try {
        Write-ColorOutput -Message "Attempting to clear print queue and restart Print Spooler service..." -Color 'Cyan'

        # Get print jobs
        $printJobs = Get-CimInstance -ClassName Win32_PrintJob -ErrorAction SilentlyContinue
        if ($null -ne $printJobs) {
            Write-ColorOutput -Message "Found $($printJobs.Count) print job(s) in the queue." -Color 'White'
            foreach ($job in $printJobs) {
                $jobId = $job.JobId
                $documentName = $job.Document
                $target = "Print Job ID: $jobId, Document: '$($documentName)'"
                if ($PSCmdlet.ShouldProcess($target, "Remove")) {
                    Write-ColorOutput -Message "Removing $target" -Color 'Cyan'
                    Remove-CimInstance -InputObject $job -ErrorAction Stop
                    Write-ColorOutput -Message "Successfully removed print job ID: $jobId." -Color 'Green'
                }
                # -WhatIf is handled implicitly by ShouldProcess
            }
        } else {
            Write-ColorOutput -Message "Print queue is already clear." -Color 'Green'
        }

        # Restart Print Spooler service
        $serviceName = "spooler"
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

        if ($null -ne $service) {
            # Stop the service if it is running
            if ($service.Status -eq [System.ServiceProcess.ServiceControllerStatus]::Running) {
                if ($PSCmdlet.ShouldProcess("Service: $serviceName", "Stop")) {
                    Write-ColorOutput -Message "Stopping Print Spooler service ($serviceName)..." -Color 'Cyan'
                    Stop-Service -Name $serviceName -Force -ErrorAction Stop
                    # Wait for the service to actually stop
                    $service.WaitForStatus([System.ServiceProcess.ServiceControllerStatus]::Stopped, [timespan]::FromSeconds(30))
                    Write-ColorOutput -Message "Print Spooler service stopped." -Color 'Green'
                }
                # -WhatIf is handled implicitly by ShouldProcess
            } else {
                Write-ColorOutput -Message "Print Spooler service ($serviceName) is not running." -Color 'DarkGray'
            }

            # Start the service if it is stopped
            # Ensure we have the latest status after potential stop
            $service.Refresh()
            if ($service.Status -eq [System.ServiceProcess.ServiceControllerStatus]::Stopped) {
                if ($PSCmdlet.ShouldProcess("Service: $serviceName", "Start")) {
                    Write-ColorOutput -Message "Starting Print Spooler service ($serviceName)..." -Color 'Cyan'
                    Start-Service -Name $serviceName -ErrorAction Stop
                    # Wait for the service to actually start
                    $service.WaitForStatus([System.ServiceProcess.ServiceControllerStatus]::Running, [timespan]::FromSeconds(30))
                    Write-ColorOutput -Message "Print Spooler service started successfully." -Color 'Green'
                }
                # -WhatIf is handled implicitly by ShouldProcess
            } else {
                Write-ColorOutput -Message "Print Spooler service ($serviceName) is already running or in a pending state." -Color 'DarkGray'
            }
        } else {
            Write-ColorOutput -Message "Print Spooler service ($serviceName) not found. Cannot restart." -Color 'Yellow'
        }

        Write-ColorOutput -Message "Operation completed." -Color 'Green'

    } catch {
        # Specific error for access denied often seen without elevation
        if ($_.Exception.InnerException -is [System.ComponentModel.Win32Exception] -and $_.Exception.InnerException.NativeErrorCode -eq 5) {
            Write-Error "Access Denied. This script requires administrative privileges. Please run PowerShell as Administrator."
        } else {
            Write-Error "An error occurred: $($_.Exception.Message)"
        }
        # Use Write-Host for red color as Write-Error doesn't directly support it without more complex formatting
        Write-ColorOutput -Message "Script execution failed." -Color 'Red'
        # Exit with a non-zero code to indicate failure
        exit 1
    }
}
