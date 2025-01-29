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
# Enable mailbox auditing.
# Block suspicious IPs.
# Produce Audit Log for the admin to review.

# Install and Import required modules
Install-Module MSOnline -Force
Install-Module ExchangeOnlineManagement -Force
Import-Module MSOnline
Import-Module ExchangeOnlineManagement


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

$userName = $upn -split "@"

$transcriptpath = ".\" + $userName[0] + "RemediationTranscript" + (Get-Date).ToString('yyyy-MM-dd') + ".txt"
Start-Transcript -Path $transcriptpath


Write-Output "You are about to remediate this account: $upn"
Write-Output "Let's get a credential and get connected to Office 365."

#Import the right module to talk with AAD
import-module MSOnline

#First, let's get us a cred!
$adminCredential = Get-Credential

    Write-Output "Connecting to Exchange Online Remote Powershell Service"
    $ExoSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $adminCredential -Authentication Basic -AllowRedirection
    if ($null -ne $ExoSession) { 
        Import-PSSession $ExoSession
    } else {
        Write-Output "  No EXO service set up for this account"
    }

    Write-Output "Connecting to EOP Powershell Service"
    $EopSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $adminCredential -Authentication Basic -AllowRedirection
    if ($null -ne $EopSession) { 
        Import-PSSession $EopSession -AllowClobber
    } else {
        Write-Output "  No EOP service set up for this account"
    }

#This connects to Azure Active Directory
Connect-MsolService -Credential $adminCredential

#Load "System.Web" assembly in PowerShell console 
[Reflection.Assembly]::LoadWithPartialName("System.Web") 

function Reset-Password($upn) {
    $newPassword = ([System.Web.Security.Membership]::GeneratePassword(16,2))
    Set-MsolUserPassword –UserPrincipalName $upn –NewPassword $newPassword -ForceChangePassword $True
    Write-Output "We've set the password for the account $upn to be $newPassword. Make sure you record this and share with the user, or be ready to reset the password again. They will have to reset their password on the next logon."
    
    Set-MsolUser -UserPrincipalName $upn -StrongPasswordRequired $True
    Write-Output "We've also set this user's account to require a strong password."
}

# Force sign-out
Revoke-AzureADUserAllRefreshToken -ObjectId $upn

function Enable-MailboxAuditing($upn) {
    Write-Output "##############################################################"
    Write-Output "We are going to enable mailbox auditing for this user to ensure we can monitor activity going forward."

    #Let's enable auditing for the mailbox in question.
    Set-Mailbox $upn -AuditEnabled $true -AuditLogAgeLimit 365

    Write-Output "##############################################################"
    Write-Output "Done! Here's the current configuration for auditing."    
    #Double-Check It!
    Get-Mailbox -Identity $upn | Select Name, AuditEnabled, AuditLogAgeLimit
}

function Remove-MailboxDelegates($upn) {
    Write-Output "##############################################################"
    Write-Output "Removing Mailbox Delegate Permissions for the affected user $upn."

    $mailboxDelegates = Get-MailboxPermission -Identity $upn | Where-Object {($_.IsInherited -ne "True") -and ($_.User -notlike "*SELF*")}
    Get-MailboxPermission -Identity $upn | Where-Object {($_.IsInherited -ne "True") -and ($_.User -notlike "*SELF*")}
    
    foreach ($delegate in $mailboxDelegates) 
    {
        Remove-MailboxPermission -Identity $upn -User $delegate.User -AccessRights $delegate.AccessRights -InheritanceType All -Confirm:$false
    }    
}

function Disable-MailforwardingRulesToExternalDomains($upn) {
    Write-Output "##############################################################"
    Write-Output "Disabling mailforwarding rules to external domains for the affected user $upn."
    Write-Output "We found the following rules that forward or redirect mail to other accounts: "
    Get-InboxRule -Mailbox $upn | Select Name, Description, Enabled, Priority, ForwardTo, ForwardAsAttachmentTo, RedirectTo, DeleteMessage, SendTextMessageNotificationTo | Where-Object {(($_.Enabled -eq $true) -and (($_.ForwardTo -ne $null) -or ($_.ForwardAsAttachmentTo -ne $null) -or ($_.RedirectTo -ne $null) -or ($_.SendTextMessageNotificationTo -ne $null)))} | Format-Table
    Get-InboxRule -Mailbox $upn | Where-Object {(($_.Enabled -eq $true) -and (($_.ForwardTo -ne $null) -or ($_.ForwardAsAttachmentTo -ne $null) -or ($_.RedirectTo -ne $null) -or ($_.SendTextMessageNotificationTo -ne $null)))} | Disable-InboxRule -Confirm:$false

    #Clean-up disabled rules
    #Get-InboxRule -Mailbox $upn | Where-Object {((($_.ForwardTo -ne $null) -or ($_.ForwardAsAttachmentTo -ne $null) -or ($_.RedirectTo -ne $null) -or ($_.SendTextMessageNotificationTo -ne $null)))} | Remove-InboxRule -Confirm:$false

    Write-Output "##############################################################"
    Write-Output "Aight. We've disabled all the rules that move your email to other mailboxes. "
}

function Delete-AllInboxRules7DaysOld
	$7DaysAgo = (Get-Date).AddDays(-7)

	Get-InboxRule -Mailbox $upn | Where-Object {$_.WhenCreated -gt $7DaysAgo} | Remove-InboxRule -Confirm:$false


function Remove-MailboxForwarding($upn) {
    Write-Output "##############################################################"
    Write-Output "Removing Mailbox Forwarding configurations for the affected user $upn. Current configuration is:"
    Get-Mailbox -Identity $upn | Select Name, DeliverToMailboxAndForward, ForwardingSmtpAddress

    Set-Mailbox -Identity $upn -DeliverToMailboxAndForward $false -ForwardingSmtpAddress $null

    Write-Output "##############################################################"
    Write-Output "Mailbox forwarding removal completed. Current configuration is:"
    Get-Mailbox -Identity $upn | Select Name, DeliverToMailboxAndForward, ForwardingSmtpAddress
}

function Enable-MFA ($upn) {

    #Create the StrongAuthenticationRequirement object and insert required settings
    $mf = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
    $mf.RelyingParty = "*"
    $mfa = @($mf)
    #Enable MFA for a user
    Set-MsolUser -UserPrincipalName $upn -StrongAuthenticationRequirements $mfa

    Write-Output "##############################################################"
    Write-Output "Aight. We've enabled MFA required for $upn. Let them know they'll need to setup their additional auth token the next time they logon."

    #Find all MFA enabled users
    Get-MsolUser -UserPrincipalName $upn | select UserPrincipalName,StrongAuthenticationMethods,StrongAuthenticationRequirements
}

function Block-SuspiciousIPs ($upn) {
	Write-Output "##############################################################"
	Write-Output "K. Now we're going to block some hacker IPs."
	$continue = $true

	while ($continue) {
		$response = Read-Host "Is there a suspicious IP address to block? (Y/N)"
		
		if ($response -eq "Y" -or $response -eq "y") {
			$ip = Read-Host "Enter the suspicious IP address"
			Set-OrganizationConfig -IPListBlocked @{Add=$ip}
			Write-Host "IP address $ip has been blocked."
		}
		elseif ($response -eq "N" -or $response -eq "n") {
			$continue = $false
			Write-Host "Continuing with the rest of the script..."
		}
		else {
			Write-Host "Invalid input. Please enter Y or N."
		}
	}
}

function Get-AuditLog ($upn) {
    Write-Output "##############################################################"
    Write-Output "We've remediated the account, but there might be things we missed. Review the audit transcript for this user to be super-sure you've got everything."

    $userName = $upn -split "@"
    $auditLogPath = ".\" + $userName[0] + "AuditLog" + (Get-Date).ToString('yyyy-MM-dd') + ".csv"
    
    $startDate = (Get-Date).AddDays(-7).ToString('MM/dd/yyyy') 
    $endDate = (Get-Date).ToString('MM/dd/yyyy')
    $results = Search-UnifiedAuditLog -StartDate $startDate -EndDate $endDate -UserIds $upn
    $results | Export-Csv -Path $auditLogPath

    Write-Output "##############################################################"
    Write-Output "We've written the log to $auditLogPath. You can also review the activity below."
    Write-Output "##############################################################"
    $results | Format-Table    
}


Reset-Password $upn
Enable-MailboxAuditing $upn
Remove-MailboxDelegates $upn
Disable-MailforwardingRulesToExternalDomains $upn
Remove-MailboxForwarding $upn
Enable-MFA $upn
Block-SuspiciousIPs $upn
Get-AuditLog $upn

Stop-Transcript

Write-Output "Remediation complete for $upn"