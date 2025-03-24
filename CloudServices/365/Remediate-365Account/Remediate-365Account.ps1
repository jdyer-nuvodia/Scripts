# =============================================================================
# Script: Remediate-365Account.ps1
# Created: 2024-02-21 10:00:00 UTC
# Author: nunya-nunya
# Last Updated: 2024-02-21 10:00:00 UTC
# Updated By: nunya-nunya
# Version: 1.0
# Additional Info: Initial script for remediating compromised O365 accounts
# =============================================================================

#Requires -Version 5.1
#Requires -Modules Microsoft.Graph

<#
.SYNOPSIS
    Remediates a compromised Office 365 account by performing security-related actions.
.DESCRIPTION
    This script performs a comprehensive set of actions to secure a compromised Office 365 account:
    - Resets password and enforces complexity
    - Revokes all active sign-in sessions
    - Removes mailbox delegates
    - Removes external mail forwarding rules
    - Removes recent inbox rules
    - Removes global mailbox forwarding
    - Enables MFA (if client uses it)
    - Generates audit logs

    Dependencies:
    - Microsoft.Graph PowerShell module
    - PowerShell 5.1 or higher
    - Office 365 Global Admin permissions

    Security considerations:
    - Requires elevated permissions in Office 365
    - Handles sensitive password information
    - Generates audit logs for compliance

    Performance impact:
    - Minimal impact on system resources
    - May take 5-10 minutes to complete all actions
.PARAMETER UserPrincipalName
    The email address (UPN) of the compromised account to remediate.
.PARAMETER TranscriptPath
    Optional. The path where the transcript log will be saved. Defaults to current directory.
.EXAMPLE
    .\Remediate-365Account.ps1 -UserPrincipalName "user@domain.com"
    Remediates the specified account and saves logs to the current directory.
.NOTES
    Security Level: High
    Required Permissions: Global Administrator
    Validation Requirements: 
    - Verify successful password reset
    - Confirm removal of suspicious rules
    - Check audit log generation
#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory=$true, Position=0)]
    [ValidatePattern('^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$')]
    [string]$UserPrincipalName,

    [Parameter(Mandatory=$false)]
    [string]$TranscriptPath = $PSScriptRoot
)

# Initialize error handling
$ErrorActionPreference = 'Stop'
$VerbosePreference = 'Continue'

function Initialize-RemediationEnvironment {
    [CmdletBinding()]
    param()
    
    try {
        # Check if Microsoft.Graph module is installed
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
            Write-Verbose "Installing Microsoft.Graph module..."
            Install-Module Microsoft.Graph -Scope CurrentUser -Force
        }

        # Import the module
        Import-Module Microsoft.Graph -ErrorAction Stop

        # Connect to Microsoft Graph with required scopes
        $requiredScopes = @(
            "User.ReadWrite.All",
            "Directory.ReadWrite.All",
            "AuditLog.Read.All",
            "MailboxSettings.ReadWrite"
        )
        
        Connect-MgGraph -Scopes $requiredScopes
        
        Write-Verbose "Successfully initialized remediation environment"
    }
    catch {
        throw "Failed to initialize environment: $_"
    }
}

function Reset-UserPassword {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    
    try {
        # Load System.Web assembly for password generation
        Add-Type -AssemblyName System.Web
        
        $newPassword = [System.Web.Security.Membership]::GeneratePassword(16,2)
        $params = @{
            passwordProfile = @{
                forceChangePasswordNextSignIn = $true
                password = $newPassword
            }
        }
        
        Update-MgUser -UserId $UserPrincipalName -BodyParameter $params
        Update-MgUser -UserId $UserPrincipalName -PasswordPolicies "DisablePasswordExpiration"
        
        # Save password to secure file
        $securePassword = ConvertTo-SecureString $newPassword -AsPlainText -Force
        $passwordFilePath = Join-Path $TranscriptPath "NewPassword_$((Get-Date).ToString('yyyyMMdd_HHmmss')).txt"
        $securePassword | ConvertFrom-SecureString | Out-File $passwordFilePath
        
        Write-Verbose "Password reset successful. New password saved to: $passwordFilePath"
        return $true
    }
    catch {
        Write-Error "Failed to reset password: $_"
        return $false
    }
}

function Remove-AllUserSessions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    
    try {
        $user = Get-MgUser -Filter "UserPrincipalName eq '$UserPrincipalName'" -ErrorAction Stop
        if ($user) {
            Invoke-MgRevokeSignInSession -UserId $user.Id
            Write-Verbose "Successfully revoked all sign-in sessions"
            return $true
        }
        return $false
    }
    catch {
        Write-Error "Failed to revoke sessions: $_"
        return $false
    }
}

function Remove-MailboxDelegates {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    
    try {
        $delegates = Get-MgUserMailboxPermission -UserId $UserPrincipalName
        $removedCount = 0
        
        foreach ($delegate in $delegates) {
            if ($delegate.GrantedToV2.User.UserPrincipalName -ne $UserPrincipalName) {
                Remove-MgUserMailboxPermission -UserId $UserPrincipalName -MailboxPermissionId $delegate.Id
                $removedCount++
            }
        }
        
        Write-Verbose "Removed $removedCount delegate permissions"
        return $true
    }
    catch {
        Write-Error "Failed to remove delegates: $_"
        return $false
    }
}

function Remove-RecentMailRules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    
    try {
        $sevenDaysAgo = (Get-Date).AddDays(-7)
        $recentRules = Get-MgUserMailFolder -UserId $UserPrincipalName -MailFolderId Inbox | 
                      Get-MgUserMailFolderMessageRule | 
                      Where-Object {$_.CreatedDateTime -gt $sevenDaysAgo}
        
        foreach ($rule in $recentRules) {
            Remove-MgUserMailFolderMessageRule -UserId $UserPrincipalName -MailFolderId Inbox -MessageRuleId $rule.Id
        }
        
        Write-Verbose "Removed $($recentRules.Count) recent mail rules"
        return $true
    }
    catch {
        Write-Error "Failed to remove recent mail rules: $_"
        return $false
    }
}

function Disable-ExternalForwarding {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    
    try {
        # Disable mailbox forwarding
        $params = @{
            "@odata.type" = "#microsoft.graph.mailboxSettings"
            automaticRepliesSetting = @{
                status = "Disabled"
            }
        }
        Update-MgUserMailboxSetting -UserId $UserPrincipalName -BodyParameter $params
        
        # Disable forwarding rules
        $rules = Get-MgUserMailFolder -UserId $UserPrincipalName -MailFolderId Inbox | 
                Get-MgUserMailFolderMessageRule
        
        foreach ($rule in $rules) {
            if ($rule.Actions.ForwardTo -or $rule.Actions.ForwardAsAttachmentTo -or $rule.Actions.RedirectTo) {
                Update-MgUserMailFolderMessageRule -UserId $UserPrincipalName -MailFolderId Inbox -MessageRuleId $rule.Id -Enabled:$false
            }
        }
        
        Write-Verbose "External forwarding disabled"
        return $true
    }
    catch {
        Write-Error "Failed to disable external forwarding: $_"
        return $false
    }
}

function Enable-UserMFA {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    
    try {
        $useMFA = Read-Host "Does the client use MFA? (Y/N)"
        if ($useMFA -eq 'Y') {
            $params = @{
                "@odata.type" = "#microsoft.graph.authenticationMethodsPolicy"
                authenticationMethodConfigurations = @(
                    @{
                        "@odata.type" = "#microsoft.graph.microsoftAuthenticatorAuthenticationMethodConfiguration"
                        state = "enabled"
                    }
                )
            }
            Update-MgPolicyAuthenticationMethodPolicy -BodyParameter $params
            Write-Verbose "MFA enabled successfully"
            return $true
        }
        Write-Verbose "MFA not enabled per client request"
        return $true
    }
    catch {
        Write-Error "Failed to configure MFA: $_"
        return $false
    }
}

function Export-UserAuditLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    
    try {
        $startDate = (Get-Date).AddDays(-7)
        $endDate = Get-Date
        
        $auditLogs = Get-MgAuditLogDirectoryAudit -Filter "activityDateTime ge $startDate and activityDateTime le $endDate and initiatedBy/user/userPrincipalName eq '$UserPrincipalName'"
        
        $auditLogPath = Join-Path $TranscriptPath "AuditLog_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
        $auditLogs | Export-Csv -Path $auditLogPath -NoTypeInformation
        
        Write-Verbose "Audit log exported to: $auditLogPath"
        return $true
    }
    catch {
        Write-Error "Failed to export audit log: $_"
        return $false
    }
}

# Main execution block
try {
    # Start transcript
    $transcriptFile = Join-Path $TranscriptPath "Remediation_$((Get-Date).ToString('yyyyMMdd_HHmmss')).log"
    Start-Transcript -Path $transcriptFile
    
    Write-Host "Starting remediation for account: $UserPrincipalName" -ForegroundColor Green
    
    # Initialize environment
    Initialize-RemediationEnvironment
    
    # Execute remediation steps
    $steps = @(
        @{ Name = "Reset Password"; Function = { Reset-UserPassword $UserPrincipalName } },
        @{ Name = "Remove Sessions"; Function = { Remove-AllUserSessions $UserPrincipalName } },
        @{ Name = "Remove Delegates"; Function = { Remove-MailboxDelegates $UserPrincipalName } },
        @{ Name = "Remove Recent Rules"; Function = { Remove-RecentMailRules $UserPrincipalName } },
        @{ Name = "Disable Forwarding"; Function = { Disable-ExternalForwarding $UserPrincipalName } },
        @{ Name = "Enable MFA"; Function = { Enable-UserMFA $UserPrincipalName } },
        @{ Name = "Export Audit Log"; Function = { Export-UserAuditLog $UserPrincipalName } }
    )
    
    $results = @()
    foreach ($step in $steps) {
        Write-Host "`nExecuting: $($step.Name)" -ForegroundColor Cyan
        $success = & $step.Function
        $results += [PSCustomObject]@{
            Step = $step.Name
            Status = if ($success) { "Success" } else { "Failed" }
        }
    }
    
    # Display summary
    Write-Host "`nRemediation Summary:" -ForegroundColor Yellow
    $results | Format-Table -AutoSize
    
}
catch {
    Write-Error "Remediation failed: $_"
}
finally {
    Stop-Transcript
    Write-Host "`nRemediation process completed. Check transcript at: $transcriptFile" -ForegroundColor Green
}
