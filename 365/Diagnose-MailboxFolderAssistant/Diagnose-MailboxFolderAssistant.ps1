$mailbox = "leadership@leadershipspokane.org"

[xml]$diag = (Export-MailboxDiagnosticLogs $mailbox -ExtendedProperties).MailboxLog
$diag.Properties.MailboxTable.Property | Where-Object {$_.Name -like "ELC*"}

Export-MailboxDiagnosticLogs $mailbox -ComponentName MRM
