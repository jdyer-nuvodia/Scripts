# =============================================================================
# Script: Send-PasswordExpiryNotification.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-10-03 15:48:08 UTC
# Updated By: jdyer-nuvodia
# Version: 2.3.0
# Additional Info: Switched to Microsoft Graph sendMail with mandatory PS1 config; removed SMTP/PSD1 config paths
# =============================================================================

<#
.SYNOPSIS
Sends password expiry notifications to Active Directory users whose passwords are nearing expiration.

.DESCRIPTION
This script queries Active Directory for enabled users with expiring passwords and sends email notifications
to warn them about upcoming password expiration. The script includes filtering to exclude users in
SharePoint-related OUs from receiving notifications.

Email delivery uses Microsoft Graph API with application permissions via a mandatory external configuration file
(`Send-PasswordExpiryNotifications-Config.ps1`) dot-sourced from the script directory. No other input methods are
supported.

Features:
- Configurable expiration warning period (default: 14 days)
- CSV logging of all notification activities with automatic cleanup of entries older than 90 days
- Testing mode for validation
- Fine-grained password policy support
- SharePoint OU exclusion filtering

.PARAMETER expireindays
Number of days before expiration to start sending notifications (configured in script variables)

.PARAMETER logging
Enable or disable CSV logging functionality (configured in script variables)

.PARAMETER testing
Enable testing mode to send all emails to test recipient (configured in script variables)
# Enabled by default in config file in order to avoid accidental notifications

.PARAMETER logRetentionDays
Number of days to retain log entries before automatic cleanup (configured in script variables, default: 90)

.EXAMPLE
.\Send-PasswordExpiryNotification.ps1
Runs the script with default settings to check for expiring passwords and send notifications

.NOTES
- Users in OUs containing "SharePoint" (case-insensitive) will be excluded from notifications
- Requires Active Directory PowerShell module
- Requires Microsoft Graph application credentials provided in the config PS1 file
- Script should be run with appropriate AD permissions
#>

##################################################################################################################
# Mandatory configuration import (no other input methods supported)
$configPath = Join-Path -Path $PSScriptRoot -ChildPath 'Send-PasswordExpiryNotifications-Config.ps1'
if (-not (Test-Path -Path $configPath)) {
    Write-Error -Message "Required config file not found at: $configPath. Aborting."
    return
}

. $configPath

# Top-level configuration variables (all expected from the config PS1)
# General notification settings
$expireindays = $expireindays
$logging = $logging
$logFile = $logFile
$testing = $testing
$testRecipient = $testRecipient
$logRetentionDays = $logRetentionDays

# Microsoft Graph configuration
$clientId = $client_id
$clientSecret = $client_secret
$tenantId = $tenant_ID
$fromEml = $fromEml
${companyName} = $companyName

# Validate required Graph configuration
$missing = @()
if ([string]::IsNullOrWhiteSpace($clientId)) { $missing += 'client_id' }
if ([string]::IsNullOrWhiteSpace($clientSecret)) { $missing += 'client_secret' }
if ([string]::IsNullOrWhiteSpace($tenantId)) { $missing += 'tenant_ID' }
if ([string]::IsNullOrWhiteSpace($fromEml)) { $missing += 'fromEml' }
if ($missing.Count -gt 0) {
    Write-Error -Message ("Missing required configuration values in config file: {0}. Aborting." -f ($missing -join ', '))
    return
}

# Validate testing recipient if testing is enabled
if ($testing -eq 'Enabled' -and [string]::IsNullOrWhiteSpace($testRecipient)) {
    Write-Error -Message 'Testing is Enabled but testRecipient is not set in the config file. Aborting.'
    return
}

###################################################################################################################

# Function to get access token using Microsoft identity platform v2.0 endpoint
function Get-AccessToken {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ClientId,
        [Parameter(Mandatory = $true)]
        [string]$ClientSecret,
        [Parameter(Mandatory = $true)]
        [string]$TenantId
    )

    $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{
        grant_type    = 'client_credentials'
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = 'https://graph.microsoft.com/.default'
    }

    $response = Invoke-RestMethod -Uri $tokenEndpoint -Method Post -Body $body
    return $response.access_token
}

# Function to send email using Microsoft Graph API
function Send-Email {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$AccessToken,
        [Parameter(Mandatory = $true)]
        [string]$RecipientEmail,
        [Parameter(Mandatory = $true)]
        [string]$Subject,
        [Parameter(Mandatory = $true)]
        [string]$Body,
        [Parameter(Mandatory = $true)]
        [string]$FromEmail
    )

    $graphApiEndpoint = "https://graph.microsoft.com/v1.0/users/$FromEmail/sendMail"
    $headers = @{
        Authorization = "Bearer $AccessToken"
        'Content-Type' = 'application/json'
    }

    $emailData = @{
        message = @{
            subject = $Subject
            body = @{
                contentType = 'HTML'
                content = $Body
            }
            toRecipients = @(
                @{
                    emailAddress = @{
                        address = $RecipientEmail
                    }
                }
            )
            from = @{
                emailAddress = @{
                    address = $FromEmail
                }
            }
        }
    }

    $emailJson = $emailData | ConvertTo-Json -Depth 100
    Invoke-RestMethod -Uri $graphApiEndpoint -Method Post -Headers $headers -Body $emailJson -ContentType 'application/json' | Out-Null
}

# Check Logging Settings
if ($logging -eq 'Enabled') {
    # Ensure log directory exists
    $logDir = Split-Path -Path $logFile -Parent
    if (-not [string]::IsNullOrWhiteSpace($logDir) -and -not (Test-Path -Path $logDir)) {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    }

    # Ensure log file exists with header
    if (-not (Test-Path -Path $logFile)) {
        New-Item -Path $logFile -ItemType File -Force | Out-Null
        Add-Content -Path $logFile -Value 'Date,Name,EmailAddress,DaystoExpire,ExpiresOn,Notified'
    }
}
# End Logging Check

# Function to clean up old log entries
function Remove-OldLogEntry {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogFilePath,
        [Parameter(Mandatory = $true)]
        [int]$RetentionDays
    )

    if (-not (Test-Path -Path $LogFilePath)) {
        return
    }

    try {
        $cutoffDate = (Get-Date).AddDays(-$RetentionDays)
        $logContent = Get-Content -Path $LogFilePath

        if ($logContent.Count -le 1) {
            # Only header exists or file is empty
            return
        }

        $header = $logContent[0]
        $filteredEntries = @($header)

        for ($i = 1; $i -lt $logContent.Count; $i++) {
            $line = $logContent[$i]
            if ([string]::IsNullOrWhiteSpace($line)) {
                continue
            }

            $fields = $line -split ','
            if ($fields.Count -ge 1) {
                $dateString = $fields[0]
                try {
                    # Parse date in ddMMyyyy format
                    if ($dateString.Length -eq 8) {
                        $day = $dateString.Substring(0, 2)
                        $month = $dateString.Substring(2, 2)
                        $year = $dateString.Substring(4, 4)
                        $entryDate = Get-Date -Year $year -Month $month -Day $day

                        if ($entryDate -ge $cutoffDate) {
                            $filteredEntries += $line
                        }
                    } else {
                        # Keep entries with invalid date format
                        $filteredEntries += $line
                    }
                } catch {
                    # Keep entries with unparseable dates
                    $filteredEntries += $line
                }
            } else {
                # Keep malformed entries
                $filteredEntries += $line
            }
        }

        # Write filtered content back to file
        if ($filteredEntries.Count -gt 1 -and $PSCmdlet.ShouldProcess($LogFilePath, "Update log file with filtered entries")) {
            Set-Content -Path $LogFilePath -Value $filteredEntries
        }
    } catch {
        # Silently continue if log cleanup fails to avoid disrupting main functionality
        Write-Warning "Failed to clean up old log entries: $($_.Exception.Message)"
    }
}

# Clean up old log entries if logging is enabled
if ($logging -eq 'Enabled') {
    Remove-OldLogEntry -LogFilePath $logFile -RetentionDays $logRetentionDays
}

# System Settings
$date = Get-Date -Format ddMMyyyy
# End System Settings

# Get Users From AD who are Enabled, Passwords Expire and are Not Currently Expired
Import-Module -Name ActiveDirectory
$users = Get-ADUser -Filter * -Properties Name, PasswordNeverExpires, PasswordExpired, PasswordLastSet, EmailAddress, DistinguishedName |
    Where-Object { $_.Enabled -eq $true } |
    Where-Object { $_.PasswordNeverExpires -eq $false } |
    Where-Object { $_.PasswordExpired -eq $false }
$DefaultmaxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge

# Acquire Microsoft Graph token once per run
$script:GraphAccessToken = Get-AccessToken -ClientId $clientId -ClientSecret $clientSecret -TenantId $tenantId

# Process Each User for Password Expiry
foreach ($user in $users) {
    $Name = $user.Name
    $emailaddress = $user.emailaddress
    $passwordSetDate = $user.PasswordLastSet
    $PasswordPol = Get-ADUserResultantPasswordPolicy -Identity $user
    # Reset Sent Flag
    $sent = ""

    # Check if user is in a SharePoint OU - skip if found
    if ($user.DistinguishedName -match 'SharePoint') {
        $sent = "Skipped - SharePoint OU"
        # If Logging is Enabled Log Details
        if ($logging -eq 'Enabled') {
            $today = (Get-Date)
            $date = Get-Date -Format ddMMyyyy
            Add-Content -Path $logFile -Value "$date,$Name,$emailaddress,N/A,N/A,$sent"
        }
        continue
    }

    # Check for Fine Grained Password
    if ($null -ne $PasswordPol) {
        $maxPasswordAge = ($PasswordPol).MaxPasswordAge
    } else {
        # No FGP set to Domain Default
        $maxPasswordAge = $DefaultmaxPasswordAge
    }

    $expireson = $passwordsetdate + $maxPasswordAge
    $today = (Get-Date)
    $daystoexpire = (New-TimeSpan -Start $today -End $Expireson).Days

    # Set Greeting based on Number of Days to Expiry.

    # Check Number of Days to Expiry
    $messageDays = $daystoexpire

    if ($daystoexpire -eq 0) {
        $messageDays = 'today.'
    } elseif ($daystoexpire -eq 1) {
        $messageDays = 'in 1 day.'
    } elseif ($daystoexpire -gt 1) {
        $messageDays = "in $daystoexpire days."
    }

    # If Testing Is Enabled - redirect to test recipient
    if ($testing -eq 'Enabled') {
        $emailaddress = $testRecipient
    } elseif ([string]::IsNullOrWhiteSpace($emailaddress)) {
        # If a user has no email address and not in testing, skip sending
        $sent = 'Skipped - No Email'
        if ($logging -eq 'Enabled') {
            Add-Content -Path $logFile -Value "$date,$Name,,N/A,N/A,$sent"
        }
        continue
    }

    # Email Subject Set Here
    $subject = "Your password will expire $messageDays"

    # Email Body Set Here, Note You can use HTML.
    $body = @"
    <p>Dear $($name),</p>
    <p> Your $([string]::IsNullOrWhiteSpace($companyName) ? 'organization' : $companyName) Domain password will expire $($messageDays)<p>
    In order to prevent a disruption of services (e.i. VPN, Windows Sign-in, Email, SharePoint) you will need to reset your password before that time:<p>
    <p><u><i>Employees</i></u><br>
    1. If you are off property, engage the FortiClient VPN.<br>
    2. Once connected to the $([string]::IsNullOrWhiteSpace($companyName) ? 'company' : $companyName) network, either via VPN or while on property, press CTRL+ALT+Delete on your keyboard and choose Change a Password.<br>
    3. Follow the prompts to enter your current password and then your desired password - twice.<br></p>
    <p>Thanks in advance,<br>
    Nuvodia Automation</p>
"@

    # Send Email Message
    if (($daystoexpire -ge 0) -and ($daystoexpire -lt $expireindays)) {
        $sent = "Yes"
        # If Logging is Enabled Log Details
        if ($logging -eq 'Enabled') {
            Add-Content -Path $logFile -Value "$date,$Name,$emailaddress,$daystoExpire,$expireson,$sent"
        }
        # Send Email Message via Microsoft Graph
        Send-Email -AccessToken $script:GraphAccessToken -RecipientEmail $emailaddress -Subject $subject -Body $body -FromEmail $fromEml
    } else {
        # Log Non Expiring Password
        $sent = "No"
        # If Logging is Enabled Log Details
        if ($logging -eq 'Enabled') {
            Add-Content -Path $logFile -Value "$date,$Name,$emailaddress,$daystoExpire,$expireson,$sent"
        }
    }
    # End Send Message
}
