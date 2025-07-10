# =============================================================================
# Script: Get-RecentAccountLockouts.ps1
# Created: 2025-04-15 22:28:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-04-15 23:34:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.3.1
# Additional Info: Retrieves recent account lockout events (Event ID 4740) from Domain Controllers. Includes transcript logging directly in script directory. Displays 'Local System' for S-1-5-18 caller SID. Added end block and improved file handle management to properly release log files. Enhanced transcript handling with proper resource cleanup via garbage collection.
# =============================================================================

#Requires -Modules ActiveDirectory

<#
.SYNOPSIS
Retrieves recent account lockout events (Event ID 4740) from Domain Controllers.

.DESCRIPTION
This script queries the Security event log on all accessible Domain Controllers for account lockout events (Event ID 4740) within a specified time frame.
It extracts relevant details such as the time of the lockout, the user account involved, the caller computer name, and the Domain Controller that logged the event.
Requires appropriate permissions to read event logs on Domain Controllers.

.PARAMETER HoursAgo
Specifies the number of hours back from the current time to search for lockout events. Defaults to 24 hours.

.PARAMETER UserName
Filters the lockout events for a specific user account. If not specified, lockouts for all users are retrieved.

.EXAMPLE
PS C:\> .\Get-RecentAccountLockouts.ps1
[Description: Retrieves account lockout events from the last 24 hours for all users from all accessible Domain Controllers.]

.EXAMPLE
PS C:\> .\Get-RecentAccountLockouts.ps1 -HoursAgo 4
[Description: Retrieves account lockout events from the last 4 hours for all users.]

.EXAMPLE
PS C:\> .\Get-RecentAccountLockouts.ps1 -UserName 'jdoe' -HoursAgo 48
[Description: Retrieves account lockout events for the user 'jdoe' from the last 48 hours.]

.NOTES
Requires membership in the 'Event Log Readers' group or equivalent permissions on the Domain Controllers.
The script attempts to query all DCs found via Get-ADDomainController. Ensure network connectivity and necessary permissions.
Performance may vary depending on the number of DCs and the volume of event logs.
Uses Get-WinEvent for event log retrieval.
Creates a transcript log file in the same directory as the script.
#>

# CmdletBinding and param must be after comment-based help
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, HelpMessage = "Specify the number of hours back to search for lockout events. Default is 24.")]
    # Limit search range for performance
        [ValidateRange(1, 720)]
    [int]$HoursAgo = 24,

    [Parameter(Mandatory = $false, HelpMessage = "Filter lockouts for a specific username.")]
    [string]$UserName
)

begin {
    # Script scope variable to track transcript status
    $script:transcriptActive = $false
    $script:logFile = $null

    # Add script termination handler to ensure transcript is always stopped
    $null = Register-EngineEvent -SourceIdentifier ([System.Management.Automation.PsEngineEvent]::Exiting) -Action {
        if ($script:transcriptActive) {
            try {
                Stop-Transcript -ErrorAction SilentlyContinue
                # Explicit attempt to release file handle even after Stop-Transcript
                if ($null -ne $script:logFile -and (Test-Path -Path $script:logFile)) {
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                }
            } catch {
                # Swallow errors on exit - we're just trying to clean up
            }
        }
    }

    # Function to safely stop transcript, improved based on Get-SetInactivityTimers.ps1
    function Stop-TranscriptSafely {
        [CmdletBinding()]
        param()

        Write-Debug "Entering Stop-TranscriptSafely function."
        # Check the script-scoped flag to see if transcript was started by this script
        if ($script:transcriptActive) {
            Write-Debug "Transcript was active, attempting to stop."
            try {
                # First try - standard Stop-Transcript
                Stop-Transcript -ErrorAction Stop
                Write-Debug "Stop-Transcript command executed."

                # Give the system a moment to release the file handle
                Start-Sleep -Milliseconds 500
                Write-Debug "Slept for 500ms after Stop-Transcript."

                # Force garbage collection to release file handles
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
                Write-Debug "Garbage collection triggered after Stop-Transcript."

                # Try a second round of garbage collection for stubborn handles
                Start-Sleep -Milliseconds 200
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()

                # Set the flag to inactive *after* successful stop
                $script:transcriptActive = $false
                Write-Host "Transcript stopped successfully." -ForegroundColor DarkGray
            } catch {
                Write-Warning "Error stopping transcript: $_"
                try {
                    # Second try - just in case the first attempt failed but didn't throw properly
                    Stop-Transcript -ErrorAction SilentlyContinue
                    Start-Sleep -Milliseconds 500
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                } catch {
                    # Ignore errors on the second attempt
                }
                # Even if stopping failed, mark as inactive to prevent retry loops
                $script:transcriptActive = $false
                Write-Debug "Transcript marked as inactive despite error during stop."
            }
        } else {
            Write-Debug "Transcript was not marked as active by this script, skipping Stop-Transcript."
        }
        Write-Debug "Exiting Stop-TranscriptSafely function."
    }
}

process {
    # Define Log Path and Start Transcript
    $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
    # Log file will be in the same directory as the script
    $script:logFile = Join-Path -Path $scriptPath -ChildPath "Get-RecentAccountLockouts_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    try {
        Start-Transcript -Path $script:logFile -Append -ErrorAction Stop
        # Mark transcript as active
                $script:transcriptActive = $true
        Write-Debug "Transcript started: $($script:logFile)"
    } catch {
        Write-Warning "Failed to start transcript logging to '$($script:logFile)'. Error: $($_.Exception.Message)"
        # Continue execution without transcript logging
    }

    try {
        Write-Host "Starting search for account lockout events (ID 4740)..." -ForegroundColor Cyan
        # Calculate start time based on current UTC time
        $startTime = (Get-Date).ToUniversalTime().AddHours(-$HoursAgo)
        $dcs = Get-ADDomainController -Filter * | Select-Object -ExpandProperty HostName

        if ($null -eq $dcs) {
            Write-Error "Could not retrieve list of Domain Controllers. Ensure the Active Directory module is available and you have permissions."
            # Stop transcript before exiting (using the safe function)
            Stop-TranscriptSafely
            # Exit if DCs cannot be retrieved
                        exit 1
        }

        Write-Host "Searching on Domain Controllers: $($dcs -join ', ')" -ForegroundColor DarkGray
        # Display the calculated UTC start time for clarity
        Write-Host "Searching for events since: $($startTime.ToString('yyyy-MM-dd HH:mm:ss')) UTC" -ForegroundColor DarkGray

        $allLockoutEvents = @()

        foreach ($dc in $dcs) {
            Write-Host "Querying Domain Controller: $dc" -ForegroundColor Cyan
            try {
                $filterHashTable = @{
                    LogName = 'Security'
                    ID = 4740
                    StartTime = $startTime
                }

                # Removed -ErrorAction Stop to prevent terminating error when no events are found
                # Continue if specific DC fails, but log warning below
                                $events = Get-WinEvent -ComputerName $dc -FilterHashtable $filterHashTable -ErrorAction SilentlyContinue

                # Check if the command succeeded before processing results
                if ($LASTEXITCODE -ne 0 -or $Error.Count -gt 0) {
                    # Log a warning if Get-WinEvent failed for reasons other than 'no events found'
                    Write-Warning "Failed to query $dc. Error details might be available above or in transcript."
                    # Clear error record after handling
                                        $Error.Clear()
                }

                # Check if $events is null or empty *after* the call
                if ($null -ne $events -and $events.Count -gt 0) {
                    Write-Host "Found $($events.Count) potential lockout events on $dc since $startTime." -ForegroundColor White

                    # Renamed $event to $lockoutEvent
                                        foreach ($lockoutEvent in $events) {
                        # Extract details from the event message or properties
                        # Property indices based on typical Event ID 4740 structure:
                        # Index 0: Target User Name
                        # Index 1: Target Domain Name (often part of user name)
                        # Index 2: Target SID (not always needed directly)
                        # Index 3: Caller Computer Name
                        # Index 4: Caller User Name (often N/A or SYSTEM)
                        # Index 5: Caller Domain Name
                        # Index 6: Caller Logon ID

                        $eventTime = $lockoutEvent.TimeCreated.ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss UTC')
                        $lockedOutUser = $lockoutEvent.Properties[0].Value
                        # Refined check for LOCAL SYSTEM SID - Ensure not null/empty before comparing
                                                $callerComputerRaw = $lockoutEvent.Properties[3].Value
                        $callerComputerDisplay = if (-not [string]::IsNullOrEmpty($callerComputerRaw) -and $callerComputerRaw.Trim() -ieq 'S-1-5-18') {
                            # Display more descriptive text without parentheses
                                                        'Local System'
                        } else {
                            # Otherwise, use the raw value (could be name or other SID)
                                                        $callerComputerRaw
                        }

                        # Apply username filter if provided
                        if (-not [string]::IsNullOrEmpty($UserName)) {
                            if ($lockedOutUser -notlike "*$UserName*") {
                                # Skip if username doesn't match
                                                                continue
                            }
                        }

                        $lockoutDetail = [PSCustomObject]@{
                            TimeLockedUTC = $eventTime
                            UserName = $lockedOutUser
                            # Use the processed display name
                                                        CallerComputer = $callerComputerDisplay
                            DomainController = $dc
                        }
                        $allLockoutEvents += $lockoutDetail
                    }
                } else {
                     # Handle case where no events were found (no error thrown now)
                     # Only write this if Get-WinEvent didn't fail for other reasons
                     if ($LASTEXITCODE -eq 0 -and $Error.Count -eq 0) {
                        Write-Host "No lockout events found on $dc within the specified timeframe." -ForegroundColor DarkGray
                     }
                }
            } catch {
                # This catch block handles unexpected errors within the loop iteration
                Write-Warning "An unexpected error occurred while processing $dc. Error: $($_.Exception.Message)"
            }
        }

        if ($allLockoutEvents.Count -gt 0) {
            Write-Host "-----------------------------------------" -ForegroundColor White
            Write-Host "Recent Account Lockout Events Found:" -ForegroundColor Green
            Write-Host "-----------------------------------------" -ForegroundColor White
            $allLockoutEvents | Sort-Object TimeLockedUTC -Descending | Format-Table -AutoSize
            Write-Host "Successfully retrieved $($allLockoutEvents.Count) lockout events." -ForegroundColor Green
        } else {
            Write-Host "-----------------------------------------" -ForegroundColor White
            Write-Host "No matching account lockout events found in the last $HoursAgo hours" -ForegroundColor Yellow
            if (-not [string]::IsNullOrEmpty($UserName)) {
                Write-Host "(Filtered for user: $UserName)" -ForegroundColor Yellow
            }
            Write-Host "-----------------------------------------" -ForegroundColor White
        }

        Write-Host "Script finished." -ForegroundColor Cyan    } catch {
        # Catch block for errors in the main try block (e.g., Get-ADDomainController failure)
        Write-Error "An error occurred during script execution: $($_.Exception.Message)"
        # Consider adding more specific error handling if needed
    } finally {
        # Stop Transcript using the safe function, but don't rely solely on this in case of unexpected termination
        Stop-TranscriptSafely
    }
}

end {
    # Ensure transcript is properly stopped and resources are cleaned up
    if ($script:transcriptActive) {
        Write-Verbose "Stopping transcript in end block to ensure proper cleanup"
        Stop-TranscriptSafely
    }

    # Unregister event handler to prevent memory leaks
    Get-EventSubscriber -SourceIdentifier ([System.Management.Automation.PsEngineEvent]::Exiting) -ErrorAction SilentlyContinue |
    Unregister-Event -ErrorAction SilentlyContinue

    # Final garbage collection
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
