# =============================================================================
# Script: Get-AllMailboxforwardingRules.ps1
# Created: 2024-02-21 10:00:00 UTC
# Author: jdyer-nuvodia
# Last Updated: 2024-02-21 11:00:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.1
# Additional Info: Added error handling, progress tracking, and parameter support
# =============================================================================

<#
.SYNOPSIS
    Retrieves and reports all mailbox forwarding rules in Exchange Online.
.DESCRIPTION
    This script connects to Exchange Online and retrieves forwarding information for all mailboxes.
    It checks both mailbox-level forwarding and inbox rules that contain forwarding actions.
    Key actions:
    - Connects to Exchange Online
    - Retrieves all mailboxes
    - Checks forwarding settings and inbox rules
    - Exports results to CSV
    Dependencies:
    - Exchange Online PowerShell Module
    - Active Exchange Online connection
.PARAMETER ExportPath
    Optional. Specify custom export path for the CSV file.
    If not specified, defaults to script directory.
.EXAMPLE
    .\Get-AllMailboxforwardingRules.ps1
    Retrieves forwarding rules for all mailboxes and exports to CSV.
.NOTES
    Security Level: Medium
    Required Permissions: Exchange Online Administrator
    Validation Requirements: Verify CSV output contains expected mailbox data
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [string]$ExportPath
)

# Verify Exchange Online connection
try {
    Write-Host "Checking Exchange Online connection..." -ForegroundColor Cyan
    $organization = Get-OrganizationConfig -ErrorAction Stop
    Write-Host "Successfully connected to Exchange Online" -ForegroundColor Green
} catch {
    Write-Error "Not connected to Exchange Online. Please run Connect-ExchangeOnline first."
    return
}

# Set export path
if (-not $ExportPath) {
    $ExportPath = Join-Path $PSScriptRoot "MailboxForwardingRules_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
}

Write-Host "Retrieving mailboxes..." -ForegroundColor Cyan
$mailboxes = Get-Mailbox -ResultSize Unlimited
$totalMailboxes = $mailboxes.Count
Write-Host "Found $totalMailboxes mailboxes to process" -ForegroundColor Green

$results = @()
$processedCount = 0
$forwardingFound = 0

foreach ($mailbox in $mailboxes) {
    $processedCount++
    $percentComplete = [math]::Round(($processedCount / $totalMailboxes) * 100, 2)
    Write-Progress -Activity "Processing Mailboxes" -Status "Processing $($mailbox.UserPrincipalName)" -PercentComplete $percentComplete

    try {
        $forwardingInfo = [PSCustomObject]@{
            UserPrincipalName = $mailbox.UserPrincipalName
            ForwardingAddress = $mailbox.ForwardingAddress
            ForwardingSmtpAddress = $mailbox.ForwardingSmtpAddress
            DeliverToMailboxAndForward = $mailbox.DeliverToMailboxAndForward
            InboxRules = $null
        }

        # Check for inbox rules with forwarding
        $inboxRules = Get-InboxRule -Mailbox $mailbox.UserPrincipalName -ErrorAction Stop | Where-Object {
            $_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo
        }

        if ($inboxRules -or $mailbox.ForwardingAddress -or $mailbox.ForwardingSmtpAddress) {
            $forwardingFound++
            Write-Host "Forwarding configuration found for $($mailbox.UserPrincipalName)" -ForegroundColor Yellow
        }

        if ($inboxRules) {
            $forwardingInfo.InboxRules = $inboxRules | ForEach-Object {
                "Rule: $($_.Name), Forward To: $($_.ForwardTo), Forward As Attachment: $($_.ForwardAsAttachmentTo), Redirect To: $($_.RedirectTo)"
            }
        }

        $results += $forwardingInfo
    } catch {
        Write-Warning "Error processing mailbox $($mailbox.UserPrincipalName): $_"
    }
}

Write-Progress -Activity "Processing Mailboxes" -Completed

# Export and display summary
try {
    $results | Export-Csv -Path $ExportPath -NoTypeInformation
    Write-Host "`nResults exported to: $ExportPath" -ForegroundColor Green
} catch {
    Write-Error "Failed to export results: $_"
}

# Display summary
Write-Host "`nSummary:" -ForegroundColor Cyan
Write-Host "Total mailboxes processed: $totalMailboxes" -ForegroundColor Green
Write-Host "Mailboxes with forwarding: $forwardingFound" -ForegroundColor Yellow
Write-Host "Export location: $ExportPath" -ForegroundColor Green
