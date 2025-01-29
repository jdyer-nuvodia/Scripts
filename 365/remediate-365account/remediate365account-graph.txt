#This script will allow you to execute a recommended set of steps to fully re-secure and remediate a known breached account in Office 365.
#It peroms the following actions:
# Reset password.
# Revokes all sign-ins.
# Remove mailbox delegates.
# Remove mailforwarding rules to external domains.
# Remove all inbox rules created in the last seven days.
# Remove global mailforwarding property on mailbox.
# Enable MFA on the user's account.
# Set password complexity on the account to be high.
# Produce Audit Log for the admin to review.

# Install and Import required modules
Install-Module Microsoft.Graph -Scope CurrentUser
Import-Module Microsoft.Graph

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All", "AuditLog.Read.All, MailboxSettings.ReadWrite"

$upn = Write-Host "Enter the email address of the compromised account."

[CmdletBinding()]
Param(
    [Parameter(Mandatory=$True,Position=0)][ValidateNotNullOrEmpty()]
		[string]$upn
		
	#[Parameter(Mandatory=$False)]
    #    [date]$startDate,
    
    #[Parameter(Mandatory=$False)]
    #    [date]$endDate,
    
    #[Parameter(Mandatory=$False)]
    #    [string]$fromFile	
)

#Get the username of the user from the UPN.
$userName = $upn -split "@"

#Set transcript path and start transcription.
$transcriptpath = ".\" + $userName[0] + "RemediationTranscript" + (Get-Date).ToString('yyyy-MM-dd') + ".txt"
Start-Transcript -Path $transcriptpath

#Notify the user the remediation is about to begin.
Write-Output "You are about to remediate this account: $upn"

# Load "System.Web" assembly in PowerShell console
[Reflection.Assembly]::LoadWithPartialName("System.Web")


function Reset-Password($upn) {
    $newPassword = ([System.Web.Security.Membership]::GeneratePassword(16,2))
    $params = @{
        passwordProfile = @{
            forceChangePasswordNextSignIn = $true
            password = $newPassword
        }
    }
    Update-MgUser -UserId $upn -BodyParameter $params
    Write-Output "Password reset for $upn. New password: $newPassword"
    
    Update-MgUser -UserId $upn -PasswordPolicies "DisablePasswordExpiration"
    Write-Output "Strong password requirement set for $upn."
}


function Logout-AllSessions($upn) {
	# Get the user's Object ID
	$objectId = (Get-MgUser -Filter "UserPrincipalName eq '$upn'").Id

	if ($objectId) {
		# Revoke all sign-in sessions for the user
		Invoke-MgRevokeSignInSession -UserId $objectId

		Write-Host "All sign-in sessions for the user with UPN $upn have been revoked."
	} else {
		Write-Host "User not found. Please check the UPN and try again."
	}
}


function Remove-MailboxDelegates($upn) {
    Write-Output "Removing mailbox delegate permissions for $upn"
    $delegates = Get-MgUserMailboxPermission -UserId $upn
    foreach ($delegate in $delegates) {
        if ($delegate.GrantedToV2.User.UserPrincipalName -ne $upn) {
            echo $delegate
			Remove-MgUserMailboxPermission -UserId $upn -MailboxPermissionId $delegate.Id
        }
    }
    Write-Output "Mailbox delegate permissions removed for $upn"
}

function Remove-RecentMailRules($upn) {
    Write-Output "Removing mail rules created in the last 7 days for $upn"
    $sevenDaysAgo = (Get-Date).AddDays(-7)
    $recentRules = Get-MgUserMailFolder -UserId $upn -MailFolderId Inbox | 
                   Get-MgUserMailFolderMessageRule | 
                   Where-Object {$_.CreatedDateTime -gt $sevenDaysAgo}
    
    foreach ($rule in $recentRules) {
        Remove-MgUserMailFolderMessageRule -UserId $upn -MailFolderId Inbox -MessageRuleId $rule.Id
    }
    Write-Output "Removed $(($recentRules | Measure-Object).Count) rules created in the last 7 days"
}


function Disable-MailforwardingRulesToExternalDomains($upn) {
    Write-Output "Disabling mailforwarding rules to external domains for $upn"
    $rules = Get-MgUserMailFolder -UserId $upn -MailFolderId Inbox | Get-MgUserMailFolderMessageRule
    foreach ($rule in $rules) {
        if ($rule.Actions.ForwardTo -or $rule.Actions.ForwardAsAttachmentTo -or $rule.Actions.RedirectTo) {
            Update-MgUserMailFolderMessageRule -UserId $upn -MailFolderId Inbox -MessageRuleId $rule.Id -Enabled:$false
        }
    }
    Write-Output "Mailforwarding rules disabled for $upn"
}


function Remove-MailboxForwarding($upn) {
    Write-Output "Removing mailbox forwarding for $upn"
    $params = @{
        "@odata.type" = "#microsoft.graph.mailboxSettings"
        automaticRepliesSetting = @{
            status = "Disabled"
        }
    }
    Update-MgUserMailboxSetting -UserId $upn -BodyParameter $params
    Write-Output "Mailbox forwarding removed for $upn"
}


function Enable-MFA($upn) {
    $useMFA = Read-Host "Does the client use MFA? (Yes/No)"
    
    if ($useMFA -eq "Yes") {
        Write-Output "Enabling MFA for $upn"
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
        Write-Output "MFA enabled for $upn"
    } else {
        Write-Output "MFA not enabled. Client does not use MFA."
    }
}


function Get-AuditLog($upn) {
    Write-Output "Retrieving audit log for $upn"
    $startDate = (Get-Date).AddDays(-7)
    $endDate = Get-Date
    $auditLogs = Get-MgAuditLogDirectoryAudit -Filter "activityDateTime ge $startDate and activityDateTime le $endDate and initiatedBy/user/userPrincipalName eq '$upn'"
    $auditLogPath = ".\" + $upn.Split("@")[0] + "AuditLog" + (Get-Date).ToString('yyyy-MM-dd') + ".csv"
    $auditLogs | Export-Csv -Path $auditLogPath
    Write-Output "Audit log exported to $auditLogPath"
}


Reset-Password $upn
Logout-AllSessions $upn
Remove-MailboxDelegates $upn
Remove-RecentMailRules $upn
Disable-MailforwardingRulesToExternalDomains $upn
Remove-MailboxForwarding $upn
Enable-MFA $upn
Get-AuditLog $upn

Stop-Transcript
Write-Output "Remediation complete for $upn"
