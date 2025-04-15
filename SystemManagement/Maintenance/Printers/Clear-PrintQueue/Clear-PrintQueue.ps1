# =============================================================================
# Script: Clear-PrintQueue.ps1
# Created: 2025-04-15 22:23:15 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-15 22:23:15 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.0
# Additional Info: Clears Windows print queue and restarts the spooler service.
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

process {
    try {
        Write-Host "Attempting to clear print queue and restart Print Spooler service..." -ForegroundColor Cyan

        # Get print jobs
        $printJobs = Get-CimInstance -ClassName Win32_PrintJob -ErrorAction SilentlyContinue
        if ($null -ne $printJobs) {
            Write-Host "Found $($printJobs.Count) print job(s) in the queue." -ForegroundColor White
            foreach ($job in $printJobs) {
                $jobId = $job.JobId
                $documentName = $job.Document
                $target = "Print Job ID: $jobId, Document: '$($documentName)'"
                if ($PSCmdlet.ShouldProcess($target, "Remove")) {
                    Write-Host "Removing $target" -ForegroundColor Cyan
                    Remove-CimInstance -InputObject $job -ErrorAction Stop
                    Write-Host "Successfully removed print job ID: $jobId." -ForegroundColor Green
                }
                # -WhatIf is handled implicitly by ShouldProcess
            }
        } else {
            Write-Host "Print queue is already clear." -ForegroundColor Green
        }

        # Restart Print Spooler service
        $serviceName = "spooler"
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

        if ($null -ne $service) {
            # Stop the service if it is running
            if ($service.Status -eq [System.ServiceProcess.ServiceControllerStatus]::Running) {
                 if ($PSCmdlet.ShouldProcess("Service: $serviceName", "Stop")) {
                    Write-Host "Stopping Print Spooler service ($serviceName)..." -ForegroundColor Cyan
                    Stop-Service -Name $serviceName -Force -ErrorAction Stop
                    # Wait for the service to actually stop
                    $service.WaitForStatus([System.ServiceProcess.ServiceControllerStatus]::Stopped, [timespan]::FromSeconds(30))
                    Write-Host "Print Spooler service stopped." -ForegroundColor Green
                 }
                 # -WhatIf is handled implicitly by ShouldProcess
            } else {
                 Write-Host "Print Spooler service ($serviceName) is not running." -ForegroundColor DarkGray
            }

            # Start the service if it is stopped
            $service.Refresh() # Ensure we have the latest status after potential stop
            if ($service.Status -eq [System.ServiceProcess.ServiceControllerStatus]::Stopped) {
                if ($PSCmdlet.ShouldProcess("Service: $serviceName", "Start")) {
                    Write-Host "Starting Print Spooler service ($serviceName)..." -ForegroundColor Cyan
                    Start-Service -Name $serviceName -ErrorAction Stop
                    # Wait for the service to actually start
                    $service.WaitForStatus([System.ServiceProcess.ServiceControllerStatus]::Running, [timespan]::FromSeconds(30))
                    Write-Host "Print Spooler service started successfully." -ForegroundColor Green
                }
                 # -WhatIf is handled implicitly by ShouldProcess
            } else {
                 Write-Host "Print Spooler service ($serviceName) is already running or in a pending state." -ForegroundColor DarkGray
            }
        } else {
             Write-Host "Print Spooler service ($serviceName) not found. Cannot restart." -ForegroundColor Yellow
        }

        Write-Host "Operation completed." -ForegroundColor Green

    } catch {
        # Specific error for access denied often seen without elevation
        if ($_.Exception.InnerException -is [System.ComponentModel.Win32Exception] -and $_.Exception.InnerException.NativeErrorCode -eq 5) {
             Write-Error "Access Denied. This script requires administrative privileges. Please run PowerShell as Administrator."
        } else {
            Write-Error "An error occurred: $($_.Exception.Message)"
        }
        # Use Write-Host for red color as Write-Error doesn't directly support it without more complex formatting
        Write-Host "Script execution failed." -ForegroundColor Red
        # Exit with a non-zero code to indicate failure
        exit 1
    }
}
