# =============================================================================
# Script: Send-PasswordExpiryNotification.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-10-03 21:35:24 UTC
# Updated By: GitHub Copilot
# Version: 2.6.2
# Additional Info: Include day 14 in notification window; retained fromLocal -> fromUser aliases
# =============================================================================

<#
.SYNOPSIS
Sends password expiry notifications to Active Directory users whose passwords are nearing expiration.

.DESCRIPTION
This script queries Active Directory for enabled users with expiring passwords and sends email notifications
to warn them about upcoming password expiration. The script includes filtering to exclude users in
SharePoint-related OUs from receiving notifications.

Email delivery uses Microsoft Graph API with application permissions. Configuration values must be provided via
Datto RMM automation job variables or as named parameters. No external config files are used.

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

.EXAMPLE
.\Send-PasswordExpiryNotification.ps1 -expireindays 14 -logging Enabled -testing Enabled -testRecipient "me@example.com" `
-client_id "<appId>" -client_secret "<secret>" -tenant_ID "<tenant>" -fromEml "noreply@contoso.com" -companyName "Contoso"
Runs the script with explicit values (useful for Datto RMM component variables)

.NOTES
- Users in OUs containing "SharePoint" (case-insensitive) will be excluded from notifications
- Requires Active Directory PowerShell module
- Requires Microsoft Graph application credentials provided in the config PS1 file
- Script should be run with appropriate AD permissions
#>

##################################################################################################################
# Parameter and Datto RMM variable handling (no external config)
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    $expireindays,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Enabled','Disabled')]
    [string]$logging,

    [Parameter(Mandatory = $false)]
    [string]$logFile,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Enabled','Disabled')]
    [string]$testing,

    [Parameter(Mandatory = $false)]
    [string]$testRecipient,

    [Parameter(Mandatory = $false)]
    $logRetentionDays,

    [Parameter(Mandatory = $false)]
    [Alias('clientId')]
    [string]$client_id,

    [Parameter(Mandatory = $false)]
    [Alias('clientSecret')]
    [string]$client_secret,

    [Parameter(Mandatory = $false)]
    [Alias('tenantId','tenantIdUpper')]
    [string]$tenant_ID,

    [Parameter(Mandatory = $false)]
    [Alias('fromEmail')]
    [string]$fromEml,

    # Optional two-part email construction variables for Datto RMM UI convenience
    [Parameter(Mandatory = $false)]
    [Alias('fromLocal')]
    [string]$fromUser,

    [Parameter(Mandatory = $false)]
    [string]$fromDomain,

    [Parameter(Mandatory = $false)]
    [string]$companyName
)

# Helper to resolve values with precedence: parameter -> env var -> default
function Resolve-ConfigValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $false)]
        [object]$ParamValue,
        [Parameter(Mandatory = $false)]
        [object]$Default,
        [Parameter(Mandatory = $false)]
        [string[]]$AltNames
    )

    # Skip unresolved Datto placeholders like $$var$$ if ever present
    function IsSet {
        param([object]$v)
        if ($null -eq $v) { return $false }
        if ($v -is [string]) {
            if ([string]::IsNullOrWhiteSpace($v)) { return $false }
            if ($v -match '^\$\$.*\$\$$') { return $false }
        }
        return $true
    }

    $candidates = @()
    $candidates += $ParamValue

    # Environment variables (Datto RMM can inject as env vars). Try name and alt names with case variants
    $envVal = $null
    $namesToTry = @($Name)
    if ($AltNames) { $namesToTry += $AltNames }
    foreach ($baseName in $namesToTry) {
        foreach ($n in @($baseName, $baseName.ToUpper(), $baseName.ToLower())) {
            $tmp = (Get-Item -Path ("env:{0}" -f $n) -ErrorAction SilentlyContinue).Value
            if (IsSet $tmp) { $envVal = $tmp; break }
        }
        if (IsSet $envVal) { break }
    }
    $candidates += $envVal

    $candidates += $Default

    foreach ($c in $candidates) {
        if (IsSet $c) { return $c }
    }
    return $Default
}

# Safely convert a possibly stringy value (e.g., "0-Default") to an integer with a fallback default
function Convert-ToIntSafe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Value,
        [Parameter(Mandatory = $true)]
        [int]$Default
    )

    if ($null -eq $Value) { return $Default }
    try {
        if ($Value -is [int]) { return [int]$Value }
        $s = [string]$Value
        if ([string]::IsNullOrWhiteSpace($s)) { return $Default }
        # Extract first integer-like token if the string contains extra text (e.g., "0-Default")
        $m = [regex]::Match($s, '-?\d+')
        if ($m.Success) { return [int]$m.Value }
        return [int]$s
    } catch { return $Default }
}

# Resolve top-level configuration values
$expireindays     = Convert-ToIntSafe -Value (Resolve-ConfigValue -Name 'expireindays' -ParamValue $expireindays -Default 14) -Default 14
$logging          = Resolve-ConfigValue -Name 'logging'         -ParamValue $logging         -Default 'Enabled'
$logFile          = Resolve-ConfigValue -Name 'logFile'         -ParamValue $logFile         -Default (Join-Path -Path $PSScriptRoot -ChildPath 'logs\Send-PasswordExpiryNotifications.csv')
$testing          = Resolve-ConfigValue -Name 'testing'         -ParamValue $testing         -Default 'Enabled'
$testRecipient    = Resolve-ConfigValue -Name 'testRecipient'   -ParamValue $testRecipient   -Default ''
$logRetentionDays = Convert-ToIntSafe -Value (Resolve-ConfigValue -Name 'logRetentionDays' -ParamValue $logRetentionDays -Default 90) -Default 90

# Microsoft Graph configuration
$clientId    = Resolve-ConfigValue -Name 'client_id'     -ParamValue $client_id     -Default $null -AltNames @('clientId')
$clientSecret= Resolve-ConfigValue -Name 'client_secret' -ParamValue $client_secret -Default $null -AltNames @('clientSecret')
$tenantId    = Resolve-ConfigValue -Name 'tenant_ID'     -ParamValue $tenant_ID     -Default $null -AltNames @('tenant_id','tenantId')
$fromEml     = Resolve-ConfigValue -Name 'fromEml'       -ParamValue $fromEml       -Default $null -AltNames @('fromEmail','fromemail','from','from_addr','from_address','fromaddress','sender','senderEmail','sender_email','mailFrom','mail_from','mailfrom')

# If fromEml not supplied, try constructing from two-part variables
if ([string]::IsNullOrWhiteSpace($fromEml)) {
    # Resolve both legacy (fromLocal) and current (fromUser) names
    $fromUser   = Resolve-ConfigValue -Name 'fromUser'   -ParamValue $fromUser   -Default $null -AltNames @('from_user','fromLocal','from_local','senderUser','sender_user','senderLocal','sender_local')
    $fromDomain = Resolve-ConfigValue -Name 'fromDomain' -ParamValue $fromDomain -Default $null -AltNames @('from_domain','senderDomain','sender_domain','mailDomain','mail_domain')
    if (-not [string]::IsNullOrWhiteSpace($fromUser) -and -not [string]::IsNullOrWhiteSpace($fromDomain)) {
        $fromEml = ("{0}@{1}" -f $fromUser.Trim('@ '), $fromDomain.Trim('@ '))
    }
}

# Light email validation (basic pattern)
function Test-EmailFormat {
    param([string]$Email)
    if ([string]::IsNullOrWhiteSpace($Email)) { return $false }
    return [bool]([regex]::IsMatch($Email, '^[^@\s]+@[^@\s]+\.[^@\s]+$'))
}
${companyName} = Resolve-ConfigValue -Name 'companyName' -ParamValue $companyName -Default ''

# Validate required Graph configuration
$missing = @()
if ([string]::IsNullOrWhiteSpace($clientId)) { $missing += 'client_id' }
if ([string]::IsNullOrWhiteSpace($clientSecret)) { $missing += 'client_secret' }
if ([string]::IsNullOrWhiteSpace($tenantId)) { $missing += 'tenant_ID' }
if (-not (Test-EmailFormat -Email $fromEml)) { $missing += 'fromEml (invalid or missing)' }
if ($missing.Count -gt 0) {
    Write-Error -Message ("Missing required configuration values: {0}. Aborting." -f ($missing -join ', '))
    return
}

# Validate testing recipient if testing is enabled
if ($testing -eq 'Enabled' -and [string]::IsNullOrWhiteSpace($testRecipient)) {
    Write-Error -Message 'Testing is Enabled but testRecipient is not set. Aborting.'
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

    # Prepare display strings (avoid PS7-only ternary)
    $orgDisplay = if ([string]::IsNullOrWhiteSpace($companyName)) { 'organization' } else { $companyName }
    $companyDisplay = if ([string]::IsNullOrWhiteSpace($companyName)) { 'company' } else { $companyName }

    # Email Body Set Here, Note You can use HTML.
    $body = @"
    <p>Dear $($name),</p>
    <p> Your $orgDisplay Domain password will expire $($messageDays)<p>
    In order to prevent a disruption of services (e.i. VPN, Windows Sign-in, Email, SharePoint) you will need to reset your password before that time:<p>
    <p><u><i>Employees</i></u><br>
    1. If you are off property, engage the FortiClient VPN.<br>
    2. Once connected to the $companyDisplay network, either via VPN or while on property, press CTRL+ALT+Delete on your keyboard and choose Change a Password.<br>
    3. Follow the prompts to enter your current password and then your desired password - twice.<br></p>
    <p>Thanks in advance,<br>
    Nuvodia Automation</p>
"@

    # Send Email Message
    if (($daystoexpire -ge 0) -and ($daystoexpire -le $expireindays)) {
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
