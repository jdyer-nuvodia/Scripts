# =============================================================================
# Script: Send-PasswordExpiryNotifications-Config.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-10-03 15:48:08 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.0
# Additional Info: Configuration file for Send-PasswordExpiryNotification.ps1 (Microsoft Graph + settings)
# =============================================================================

<#
.SYNOPSIS
Configuration values for Send-PasswordExpiryNotification.ps1.

.DESCRIPTION
Populate these variables and save the file in the same folder as the script. The main script will dot-source
this file and only run if it is present with the required values. Do not commit secrets to source control.
Use secure secret stores for client secrets when possible.

.NOTES
- Required fields: $client_id, $client_secret, $tenant_ID, $fromEml
- Testing mode will redirect all emails to $testRecipient if set to "Enabled"
#>

# General settings
$script:expireindays = 14                  # Integer: Number of days ahead to notify users
$script:logging = "Enabled"                # "Enabled" or "Disabled"
$script:logFile = "C:\\Password_Change_Notification\\Password_Change_Notification_14d_log.csv"  # CSV log path
$script:testing = "Enabled"               # "Enabled" or "Disabled". If Enabled, all mail goes to $testRecipient
$script:testRecipient = 'jdyer@nuvodia.com'                 # Your test email address when testing is Enabled
$script:logRetentionDays = 90              # Integer: days of log retention

# Microsoft Graph application credentials (Application permissions)
$script:client_id = ''                     # Azure App Registration (Application ID)
$script:client_secret = ''                 # Client secret for the App Registration (secure appropriately)
$script:tenant_ID = ''                     # Azure AD tenant ID (GUID)
$script:fromEml = ''                       # Mailbox/user principal to send as (e.g., no-reply@contoso.com)
${script:companyName} = ''                 # Optional: Company name for email body (e.g., "Contoso")

# Export values to calling scope when dot-sourced
Set-Variable -Name expireindays -Scope Script -Value $script:expireindays
Set-Variable -Name logging -Scope Script -Value $script:logging
Set-Variable -Name logFile -Scope Script -Value $script:logFile
Set-Variable -Name testing -Scope Script -Value $script:testing
Set-Variable -Name testRecipient -Scope Script -Value $script:testRecipient
Set-Variable -Name logRetentionDays -Scope Script -Value $script:logRetentionDays
Set-Variable -Name client_id -Scope Script -Value $script:client_id
Set-Variable -Name client_secret -Scope Script -Value $script:client_secret
Set-Variable -Name tenant_ID -Scope Script -Value $script:tenant_ID
Set-Variable -Name fromEml -Scope Script -Value $script:fromEml
Set-Variable -Name companyName -Scope Script -Value $script:companyName
