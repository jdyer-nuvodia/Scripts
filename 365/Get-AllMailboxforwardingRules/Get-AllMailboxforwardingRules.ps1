# Connect to Exchange Online
Connect-ExchangeOnline

# Get all mailboxes
$mailboxes = Get-Mailbox -ResultSize Unlimited

# Initialize an array to store results
$results = @()

foreach ($mailbox in $mailboxes) {
    $forwardingInfo = [PSCustomObject]@{
        UserPrincipalName = $mailbox.UserPrincipalName
        ForwardingAddress = $mailbox.ForwardingAddress
        ForwardingSmtpAddress = $mailbox.ForwardingSmtpAddress
        DeliverToMailboxAndForward = $mailbox.DeliverToMailboxAndForward
        InboxRules = $null
    }

    # Check for inbox rules with forwarding
    $inboxRules = Get-InboxRule -Mailbox $mailbox.UserPrincipalName | Where-Object {
        $_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo
    }

    if ($inboxRules) {
        $forwardingInfo.InboxRules = $inboxRules | ForEach-Object {
            "Rule: $($_.Name), Forward To: $($_.ForwardTo), Forward As Attachment: $($_.ForwardAsAttachmentTo), Redirect To: $($_.RedirectTo)"
        }
    }

    $results += $forwardingInfo
}

# Display results
$results | Format-Table -AutoSize -Wrap

# Optionally, export results to CSV
$results | Export-Csv -Path "C:\users\jdyer\OneDrive - Nuvodia\Documents\WindowsPowershell\Scripts\getAllMailboxForwardingRules\MailboxForwardingRules.csv" -NoTypeInformation
