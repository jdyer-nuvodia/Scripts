# =============================================================================
# Script: Disable-InactiveUsers.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-09-29 21:16:08 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.1
# Additional Info: Standards refactor and PSScriptAnalyzer PASS (no warnings)
# =============================================================================

<#
.SYNOPSIS
Automates disabling inactive AD users and performs post-disable hygiene.

.DESCRIPTION
This unattended script will:
 - Disable accounts inactive for a specified number of days
 - Ensure primary group is set to a specified group (default: DisabledPrimary)
 - Move accounts into a target Disabled Users OU
 - Update Description to "User disabled <UTC date>"

Notes:
 - Uses CmdletBinding with -WhatIf support and ConfirmImpact 'High'
 - All actions are idempotent and wrapped with ShouldProcess
 - Logging is captured via transcript in a Logs folder next to this script, including computer name and UTC timestamp
 - No interactive prompts or UI elements; suitable for Task Scheduler

.PARAMETER PrimaryGroupName
The AD group to set as the user's primary group. Default: DisabledPrimary

.PARAMETER TargetOU
The distinguished name of the OU where disabled users should be moved.

.PARAMETER TargetSearchOU
The distinguished name of the search base used to find inactive/disabled users.

.PARAMETER ExcludedOUs
An array of distinguished names (OUs) to exclude from processing.

.PARAMETER ExemptUsers
An array of SamAccountNames to exclude from processing.

.PARAMETER DaysInactive
Number of days of inactivity (LastLogonTimeStamp) before disabling accounts. Default: 30

.PARAMETER LogDirectory
Directory path where logs will be written. Default: <script folder>\Logs

.EXAMPLE
Disable-InactiveUsers.ps1 -TargetOU "OU=Disabled Users,OU=Disabled Accounts,DC=DOMAIN,DC=local" -TargetSearchOU "DC=DOMAIN,DC=local" -ExcludedOUs @("OU=Service Accounts,OU=MyBusiness,DC=DOMAIN,DC=local") -WhatIf
Previews actions without making changes.

.EXAMPLE
Disable-InactiveUsers.ps1 -TargetOU "OU=Disabled Users,OU=Disabled Accounts,DC=DOMAIN,DC=local" -TargetSearchOU "DC=DOMAIN,DC=local" -ExcludedOUs @("OU=Sharepoint,OU=Users,OU=MyBusiness,DC=DOMAIN,DC=local","OU=VendorStandard,OU=Users,OU=MyBusiness,DC=DOMAIN,DC=local") -DaysInactive 45
Runs unattended and logs results.

.NOTES
Security Level: High
Required Permissions: AD permissions to disable accounts, modify groups, move objects
Dependencies: ActiveDirectory module
Validation: Review the log file for summary and per-user actions
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$PrimaryGroupName = 'DisabledPrimary',

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$TargetOU,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$TargetSearchOU,

    [Parameter(Mandatory = $false)]
    [string[]]$ExcludedOUs = @(),

    [Parameter(Mandatory = $false)]
    [string[]]$ExemptUsers = @(),

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 3650)]
    [int]$DaysInactive = 30,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$LogDirectory = (Join-Path -Path $PSScriptRoot -ChildPath 'Logs')
)

# Color support variables and Write-ColorOutput function
$Script:UseAnsiColors = $PSVersionTable.PSVersion.Major -ge 7
$Script:Colors = if ($Script:UseAnsiColors) {
    @{
        'White'    = "`e[37m"
        'Cyan'     = "`e[36m"
        'Green'    = "`e[32m"
        'Yellow'   = "`e[33m"
        'Red'      = "`e[31m"
        'Magenta'  = "`e[35m"
        'DarkGray' = "`e[90m"
        'Reset'    = "`e[0m"
    }
} else {
    @{
        'White'    = [ConsoleColor]::White
        'Cyan'     = [ConsoleColor]::Cyan
        'Green'    = [ConsoleColor]::Green
        'Yellow'   = [ConsoleColor]::Yellow
        'Red'      = [ConsoleColor]::Red
        'Magenta'  = [ConsoleColor]::Magenta
        'DarkGray' = [ConsoleColor]::DarkGray
        'Reset'    = ''
    }
}

function Write-ColorOutput {
    <#
    .SYNOPSIS
    Outputs colored text in a way that is compatible with PSScriptAnalyzer requirements.

    .DESCRIPTION
    Uses Write-Output only. For PowerShell 7+, ANSI codes are embedded; PowerShell 5.1 temporarily sets console color and resets.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$Color = 'White'
    )

    if ($Script:UseAnsiColors) {
        $colorCode = $Script:Colors[$Color]
        $resetCode = $Script:Colors.Reset
        Write-Output "${colorCode}${Message}${resetCode}"
    } else {
        $originalColor = $Host.UI.RawUI.ForegroundColor
        try {
            if ($Script:Colors[$Color] -and $Script:Colors[$Color] -ne '') {
                $Host.UI.RawUI.ForegroundColor = $Script:Colors[$Color]
            }
            Write-Output $Message
        } finally {
            $Host.UI.RawUI.ForegroundColor = $originalColor
        }
    }
}

# Prepare logging
try {
    if (-not (Test-Path -Path $LogDirectory -PathType Container)) {
        $null = New-Item -Path $LogDirectory -ItemType Directory -Force
    }
    $UtcStamp   = (Get-Date -AsUTC -Format 'yyyy-MM-dd_HH-mm-ss')
    $Computer   = $env:COMPUTERNAME
    $LogPath    = Join-Path -Path $LogDirectory -ChildPath ("Disable-InactiveUsers_${Computer}_${UtcStamp}.log")
    Start-Transcript -Path $LogPath -ErrorAction Stop
} catch {
    Write-ColorOutput -Message "[SYSTEM ERROR DETECTED] Failed to initialize transcript logging: $($_.Exception.Message)" -Color 'Red'
    exit 1
}

Write-ColorOutput -Message "=======================================================" -Color 'White'
Write-ColorOutput -Message "Disable-InactiveUsers.ps1" -Color 'White'
Write-ColorOutput -Message "=======================================================" -Color 'White'

Write-ColorOutput -Message "CONFIGURATION:" -Color 'White'
Write-ColorOutput -Message "PrimaryGroupName: $PrimaryGroupName" -Color 'Cyan'
Write-ColorOutput -Message "TargetOU: $TargetOU" -Color 'Cyan'
Write-ColorOutput -Message "TargetSearchOU: $TargetSearchOU" -Color 'Cyan'
Write-ColorOutput -Message "ExcludedOUs: $([string]::Join('; ', $ExcludedOUs))" -Color 'Cyan'
Write-ColorOutput -Message "DaysInactive: $DaysInactive" -Color 'Cyan'
Write-ColorOutput -Message "Log file: $LogPath" -Color 'DarkGray'
Write-ColorOutput -Message "-------------------------------------------------------" -Color 'DarkGray'

# Import ActiveDirectory module
try {
    Import-Module -Name ActiveDirectory -ErrorAction Stop
} catch {
    Write-ColorOutput -Message "ERROR: ActiveDirectory module is not available. $($_.Exception.Message)" -Color 'Red'
    Stop-Transcript | Out-Null
    exit 1
}

# Calculate inactivity cutoff
$CutoffDate = (Get-Date).AddDays(-1 * $DaysInactive)

# Disable inactive users (idempotent, gated by ShouldProcess)
try {
    $InactiveUsers = Get-ADUser -Filter { LastLogonTimeStamp -lt $CutoffDate } -SearchBase $TargetSearchOU -Properties SamAccountName, LastLogonTimeStamp -ErrorAction Stop

    foreach ($IU in $InactiveUsers) {
        $dn = $IU.DistinguishedName
        $skip = $false
        foreach ($ex in $ExcludedOUs) {
            if ([string]::IsNullOrWhiteSpace($ex)) { continue }
            if ($dn -like ("*${ex}*")) { $skip = $true; break }
        }
        if ($skip) { continue }

        if ($ExemptUsers -contains $IU.SamAccountName) { continue }

        if ($PSCmdlet.ShouldProcess($IU.SamAccountName, 'Disable AD account')) {
            try {
                Disable-ADAccount -Identity $IU.SamAccountName -ErrorAction Stop -Confirm:$false
                Write-ColorOutput -Message "Disabled account: $($IU.SamAccountName)" -Color 'Green'
            } catch {
                Write-ColorOutput -Message "Error disabling account $($IU.SamAccountName): $($_.Exception.Message)" -Color 'Red'
            }
        }
    }
} catch {
    Write-ColorOutput -Message "Error during inactivity disable phase: $($_.Exception.Message)" -Color 'Red'
}

# Resolve primary group token id
try {
    $PGInfo = Get-ADGroup -Identity $PrimaryGroupName -Properties primaryGroupToken -ErrorAction Stop
    $PrimaryGroupToken = [int]$PGInfo.primaryGroupToken
} catch {
    Write-ColorOutput -Message "ERROR: Failed to get primary group token for '$PrimaryGroupName' - $($_.Exception.Message)" -Color 'Red'
    Stop-Transcript | Out-Null
    exit 1
}

# Find all disabled users within scope
try {
    $DisabledUsers = Search-ADAccount -AccountDisabled -UsersOnly -ResultPageSize 2000 -ResultSetSize $null -SearchBase $TargetSearchOU -SearchScope Subtree -ErrorAction Stop
    if ($ExcludedOUs.Count -gt 0) {
        $DisabledUsers = $DisabledUsers | Where-Object -FilterScript {
            $dn = $_.DistinguishedName
            -not ($ExcludedOUs | ForEach-Object -Process { if (-not [string]::IsNullOrWhiteSpace($_)) { $dn -like ("*$_*") } })
        }
    }
    if ($ExemptUsers.Count -gt 0) {
        $DisabledUsers = $DisabledUsers | Where-Object -FilterScript { $_.SamAccountName -notin $ExemptUsers }
    }
    $DisabledUsers = $DisabledUsers | Sort-Object -Property SamAccountName
    $DisabledUsersCount = ($DisabledUsers | Measure-Object).Count
} catch {
    Write-ColorOutput -Message "ERROR: Failed to enumerate disabled users - $($_.Exception.Message)" -Color 'Red'
    Stop-Transcript | Out-Null
    exit 1
}

Write-ColorOutput -Message "PROCESSING:" -Color 'White'
Write-ColorOutput -Message "Identified $DisabledUsersCount disabled user accounts" -Color 'Cyan'
Write-ColorOutput -Message "-------------------------------------------------------" -Color 'DarkGray'

# Counters
$ProcessedUsers = 0

function Set-PrimaryGroup {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory = $true)]
        [string]$User,
        [Parameter(Mandatory = $true)]
        [int]$CurrentPrimaryGroupId,
        [Parameter(Mandatory = $true)]
        [int]$TargetPrimaryGroupId,
        [Parameter(Mandatory = $true)]
        [string]$TargetPrimaryGroupName
    )
    if ($CurrentPrimaryGroupId -ne $TargetPrimaryGroupId) {
        Write-ColorOutput -Message "Updating Primary Group to $TargetPrimaryGroupName ($TargetPrimaryGroupId) for $User" -Color 'Yellow'
        try {
            if ($PSCmdlet.ShouldProcess($User, "Add to '$TargetPrimaryGroupName' and set PrimaryGroupID")) {
                Add-ADGroupMember -Identity $TargetPrimaryGroupName -Members $User -Confirm:$false -ErrorAction Stop
                Set-ADUser -Identity $User -Replace @{ primaryGroupID = $TargetPrimaryGroupId } -ErrorAction Stop
            }
        } catch {
            Write-ColorOutput -Message "Error setting primary group for ${User}: $($_.Exception.Message)" -Color 'Red'
        }
    } else {
        Write-ColorOutput -Message "$User already has primary group '$TargetPrimaryGroupName'" -Color 'DarkGray'
    }
}

function Move-UserToOu {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory = $true)]
        [string]$User,
        [Parameter(Mandatory = $true)]
        [string]$CurrentOu,
        [Parameter(Mandatory = $true)]
        [string]$DestinationOu
    )
    if ($CurrentOu -ne $DestinationOu) {
        try {
            if ($PSCmdlet.ShouldProcess($User, "Move account to '$DestinationOu'")) {
                $userObj = Get-ADUser -Identity $User -ErrorAction Stop
                Move-ADObject -Identity $userObj.DistinguishedName -TargetPath $DestinationOu -Confirm:$false -ErrorAction Stop
                Write-ColorOutput -Message "Moved $User to $DestinationOu" -Color 'Green'
            }
        } catch {
            Write-ColorOutput -Message "Error moving $User to ${DestinationOu}: $($_.Exception.Message)" -Color 'Red'
        }
    } else {
        Write-ColorOutput -Message "$User already in target OU" -Color 'DarkGray'
    }
}

function Set-UserDescription {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [string]$CurrentDescription,
        [Parameter(Mandatory = $true)]
        [string]$User
    )
    $DisabledDate = Get-Date -AsUTC
    if (($null -eq $CurrentDescription) -or (-not $CurrentDescription.StartsWith('User disabled', [System.StringComparison]::OrdinalIgnoreCase))) {
        try {
            if ($PSCmdlet.ShouldProcess($User, "Update Description to 'User disabled $DisabledDate'")) {
                Set-ADUser -Identity $User -Description "User disabled $DisabledDate" -ErrorAction Stop
                Write-ColorOutput -Message "Updated description for $User" -Color 'Green'
            }
        } catch {
            Write-ColorOutput -Message "Error updating description for ${User}: $($_.Exception.Message)" -Color 'Red'
        }
    } else {
        Write-ColorOutput -Message "Description already indicates disabled status for $User" -Color 'DarkGray'
    }
}

# Main processing loop
try {
    foreach ($DU in $DisabledUsers) {
        $User = $DU.SamAccountName
        if ([string]::IsNullOrWhiteSpace($User)) { continue }
        if ($ExemptUsers -contains $User) {
            Write-ColorOutput -Message "Skipping exempt user $User" -Color 'Yellow'
            continue
        }

        $ProcessedUsers++
        $UserDetails = Get-ADUser -Identity $User -Properties DistinguishedName, PrimaryGroupID, Description -ErrorAction Stop

        # Derive current OU name from DN
        $dn = $UserDetails.DistinguishedName
        $ouStart = $dn.IndexOf('OU=', [System.StringComparison]::CurrentCultureIgnoreCase)
        $currentOu = if ($ouStart -ge 0) { $dn.Substring($ouStart) } else { $null }

        # Primary group step
        Set-PrimaryGroup -User $User -CurrentPrimaryGroupId $UserDetails.PrimaryGroupID -TargetPrimaryGroupId $PrimaryGroupToken -TargetPrimaryGroupName $PrimaryGroupName

        # Move to target OU step
        if ($null -ne $currentOu) {
            Move-UserToOu -User $User -CurrentOu $currentOu -DestinationOu $TargetOU
        }

        # Description update
        Set-UserDescription -CurrentDescription $UserDetails.Description -User $User

        Write-ColorOutput -Message "-------------------------------------------------------" -Color 'DarkGray'
    }
} catch {
    Write-ColorOutput -Message "Critical error in main processing loop: $($_.Exception.Message)" -Color 'Red'
}

# Summary
Write-ColorOutput -Message "SUMMARY:" -Color 'White'
Write-ColorOutput -Message "Processed users: $ProcessedUsers of $DisabledUsersCount" -Color 'Cyan'
Write-ColorOutput -Message "Log file saved to: $LogPath" -Color 'DarkGray'
Write-ColorOutput -Message "=======================================================" -Color 'White'

try { Stop-Transcript | Out-Null } catch { Write-Error -Message "Failed to stop transcript: $($_.Exception.Message)" }

exit 0
