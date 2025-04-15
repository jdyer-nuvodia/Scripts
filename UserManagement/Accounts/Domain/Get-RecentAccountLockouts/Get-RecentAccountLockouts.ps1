# =============================================================================
# Script: Get-RecentAccountLockouts.ps1
# Created: 2025-04-15 22:28:00 UTC
# Author: GitHub Copilot
# Last Updated: 2025-04-15 22:46:00 UTC # Approximate current time
# Updated By: GitHub Copilot
# Version: 1.2.0
# Additional Info: Retrieves recent account lockout events (Event ID 4740) from Domain Controllers. Includes transcript logging directly in script directory. Displays '(Local System)' for S-1-5-18 caller SID.
# =============================================================================

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

#Requires -Modules ActiveDirectory

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, HelpMessage="Specify the number of hours back to search for lockout events. Default is 24.")]
    [ValidateRange(1, 720)] # Limit search range for performance
    [int]$HoursAgo = 24,

    [Parameter(Mandatory=$false, HelpMessage="Filter lockouts for a specific username.")]
    [string]$UserName
)

process {
    # Define Log Path and Start Transcript
    $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
    # Log file will be in the same directory as the script
    $logFile = Join-Path -Path $scriptPath -ChildPath "Get-RecentAccountLockouts_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    try {
        Start-Transcript -Path $logFile -Append -ErrorAction Stop
    } catch {
        Write-Warning "Failed to start transcript logging to '$logFile'. Error: $($_.Exception.Message)"
        # Continue execution without transcript logging
    }

    try {
        Write-Host "Starting search for account lockout events (ID 4740)..." -ForegroundColor Cyan
        $startTime = (Get-Date).AddHours(-$HoursAgo)
        $dcs = Get-ADDomainController -Filter * | Select-Object -ExpandProperty HostName

        if ($null -eq $dcs) {
            Write-Error "Could not retrieve list of Domain Controllers. Ensure the Active Directory module is available and you have permissions."
            exit 1
        }

        Write-Host "Searching on Domain Controllers: $($dcs -join ', ')" -ForegroundColor DarkGray
        Write-Host "Searching for events since: $($startTime.ToString('yyyy-MM-dd HH:mm:ss')) UTC" -ForegroundColor DarkGray

        $allLockoutEvents = @()

        foreach ($dc in $dcs) {
            Write-Host "Querying Domain Controller: $dc" -ForegroundColor Cyan
            try {
                $filterHashTable = @{
                    LogName   = 'Security'
                    ID        = 4740
                    StartTime = $startTime
                }

                $events = Get-WinEvent -ComputerName $dc -FilterHashtable $filterHashTable -ErrorAction Stop

                if ($null -ne $events) {
                    Write-Host "Found $($events.Count) potential lockout events on $dc since $startTime." -ForegroundColor White

                    foreach ($event in $events) {
                        # Extract details from the event message or properties
                        # Property indices based on typical Event ID 4740 structure:
                        # Index 0: Target User Name
                        # Index 1: Target Domain Name (often part of user name)
                        # Index 2: Target SID (not always needed directly)
                        # Index 3: Caller Computer Name
                        # Index 4: Caller User Name (often N/A or SYSTEM)
                        # Index 5: Caller Domain Name
                        # Index 6: Caller Logon ID

                        $eventTime = $event.TimeCreated.ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss UTC')
                        $lockedOutUser = $event.Properties[0].Value
                        $callerComputerRaw = $event.Properties[3].Value

                        # Check if the caller computer is the LOCAL SYSTEM SID
                        $callerComputerDisplay = if ($callerComputerRaw -eq 'S-1-5-18') {
                            '(Local System)' # Display more descriptive text
                        } else {
                            $callerComputerRaw # Otherwise, use the raw value
                        }

                        # Apply username filter if provided
                        if (-not [string]::IsNullOrEmpty($UserName)) {
                            if ($lockedOutUser -notlike "*$UserName*") {
                                continue # Skip if username doesn't match
                            }
                        }

                        $lockoutDetail = [PSCustomObject]@{
                            TimeLockedUTC  = $eventTime
                            UserName       = $lockedOutUser
                            CallerComputer = $callerComputerDisplay # Use the processed display name
                            DomainController = $dc
                        }
                        $allLockoutEvents += $lockoutDetail
                    }
                } else {
                     Write-Host "No lockout events found on $dc within the specified timeframe." -ForegroundColor DarkGray
                }
            } catch {
                Write-Warning "Failed to query $dc. Error: $($_.Exception.Message)"
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

        Write-Host "Script finished." -ForegroundColor Cyan

    } catch {
        # Existing catch block for main script logic errors
        Write-Error "An error occurred during script execution: $($_.Exception.Message)"
        # Consider adding more specific error handling if needed
    } finally {
        # Stop Transcript
        if ($global:Transcript) { # Check if transcript is active before stopping
            Write-Host "Attempting to stop transcript..." -ForegroundColor DarkGray
            Stop-Transcript
            Write-Host "Transcript stopped." -ForegroundColor DarkGray
        }
    }
}
