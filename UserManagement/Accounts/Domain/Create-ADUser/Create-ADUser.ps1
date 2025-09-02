# =============================================================================
# Script: Create-ADUser.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 16:50:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.4.2
# Additional Info: Aligned operators vertically for PSScriptAnalyzer compliance
# =============================================================================

<#
.SYNOPSIS
    Creates a new Active Directory user and adds them to specified groups.
.DESCRIPTION
    This script creates a new Active Directory user account with specified parameters
    and optionally adds them to designated AD groups. It includes:

    - Secure password handling using SecureString
    - Parameter validation
    - Group membership management
    - Error handling and logging
    - Support for -WhatIf to preview changes

    Dependencies:
    - Active Directory PowerShell module
    - Domain Admin or Account Operator permissions
    - Network connectivity to Domain Controller

    The script performs the following actions:
    1. Validates all input parameters
    2. Processes secure password
    3. Creates new AD user account
    4. Adds user to specified groups
    5. Logs all actions and results
.PARAMETER Name
    Full name of the user (display name in AD)
.PARAMETER GivenName
    First name of the user
.PARAMETER Surname
    Last name of the user
.PARAMETER SamAccountName
    SAM account name for the user (pre-Windows 2000 login name)
    Must be unique in the domain
.PARAMETER UserPrincipalName
    User Principal Name in email format (username@domain.com)
    Must be unique in the domain
.PARAMETER Password
    Initial password for the user as SecureString
    Must meet domain password complexity requirements
    Use: Read-Host -AsSecureString "Enter Password"
.PARAMETER OUPath
    Distinguished Name of the OU where the user will be created
    Example: "OU = Users, OU = Company, DC = domain, DC = com"
.PARAMETER Groups
    Array of group names to add the user to
    Optional - if not specified, user will only have default group memberships
.EXAMPLE
    $securePass = Read-Host -AsSecureString "Enter Password"
    .\Create-ADUser.ps1 -Name "John Doe" -GivenName "John" -Surname "Doe" `
        -SamAccountName "jdoe" -UserPrincipalName "jdoe@domain.com" `
        -Password $securePass -OUPath "OU = Users, DC = domain, DC = com"

    Creates a new user account for John Doe with minimal parameters using secure password input
.EXAMPLE
    $securePass = ConvertTo-SecureString "InitialP@ss123" -AsPlainText -Force
    .\Create-ADUser.ps1 -Name "Jane Smith" -GivenName "Jane" -Surname "Smith" `
        -SamAccountName "jsmith" -UserPrincipalName "jsmith@domain.com" `
        -Password $securePass -OUPath "OU = Sales, OU = Users, DC = domain, DC = com" `
        -Groups @("Sales Team", "Remote Users", "VPN Users")

    Creates a new user account for Jane Smith and adds her to multiple groups
.NOTES
    Security Level: High
    Required Permissions: Domain Admin or delegated user creation rights
    Validation Requirements:
    - Verify AD module is available
    - Verify OU path exists
    - Verify groups exist
    - Verify account name uniqueness
    - Verify password meets complexity requirements
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Name,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$GivenName,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Surname,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[a-zA-Z0-9\-\.]{ 1, 20}$')]
    [string]$SamAccountName,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[\w-]+(\.[\w-]+)*@([\w-]+\.)+[\w-]+$')]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $true)]
    [ValidateNotNull()]
    [System.Security.SecureString]$Password,

    [Parameter(Mandatory = $true)]
    [ValidateScript({
            if (-not (Test-Path "AD:$_")) {
                throw "OU path does not exist in Active Directory"
            }
            return $true
        })]
    [string]$OUPath,

    [Parameter(Mandatory = $false)]
    [string[]]$Groups = @()
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
    Outputs colored text in a way that's compatible with PSScriptAnalyzer requirements.

    .DESCRIPTION
    This function provides colored output while maintaining compatibility with PSScriptAnalyzer
    by using only Write-Output and standard PowerShell cmdlets.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$Color = "White"
    )

    # Always use Write-Output to satisfy PSScriptAnalyzer
    # For PowerShell 7+, include ANSI color codes in the output
    if ($Script:UseAnsiColors) {
        $colorCode = $Script:Colors[$Color]
        $resetCode = $Script:Colors.Reset
        Write-Output "${colorCode}${Message}${resetCode}"
    } else {
        # For PowerShell 5.1, just output the message
        # Color formatting will be handled by the terminal/host if supported
        Write-Output $Message
    }
}


# Initialize logging
$LogPath                = Join-Path $PSScriptRoot "ADUserCreation_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$ErrorActionPreference  = "Stop"

function Write-LogMessage {
    param($Message, $Level = "Information")

    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC"
    $LogMessage = "$TimeStamp [$Level] $Message"
    Add-Content -Path $LogPath -Value $LogMessage

    switch ($Level) {
        "Information" { Write-ColorOutput -Message $Message -Color 'White' }
        "Success" { Write-ColorOutput -Message $Message -Color 'Green' }
        "Warning" { Write-ColorOutput -Message $Message -Color 'Yellow' }
        "Error" { Write-ColorOutput -Message $Message -Color 'Red' }
    }
}

try {
    Write-LogMessage "Starting user creation process for: $Name" "Information"

    # Verify AD module is loaded
    if (-not (Get-Module -Name ActiveDirectory)) {
        Import-Module ActiveDirectory -ErrorAction Stop
        Write-LogMessage "Loaded Active Directory module" "Success"
    }

    # Verify SamAccountName is unique
    if (Get-ADUser -Filter { SamAccountName -eq $SamAccountName } -ErrorAction SilentlyContinue) {
        throw "SamAccountName '$SamAccountName' already exists"
    }

    # Create the new user
    Write-LogMessage "Creating new AD user account..." "Information"
    if ($PSCmdlet.ShouldProcess($Name, "Create new AD user with SamAccountName '$SamAccountName'")) {
        New-ADUser -Name $Name `
            -GivenName $GivenName `
            -Surname $Surname `
            -SamAccountName $SamAccountName `
            -UserPrincipalName $UserPrincipalName `
            -AccountPassword $Password `
            -Enabled $true `
            -Path $OUPath `
            -ErrorAction Stop

        Write-LogMessage "Successfully created user account: $Name" "Success"

        # Add user to specified groups
        if ($Groups.Count -gt 0) {
            Write-LogMessage "Adding user to specified groups..." "Information"
            foreach ($Group in $Groups) {
                try {
                    if ($PSCmdlet.ShouldProcess($SamAccountName, "Add to group '$Group'")) {
                        Add-ADGroupMember -Identity $Group -Members $SamAccountName -ErrorAction Stop
                        Write-LogMessage "Added to group: $Group" "Success"
                    }
                } catch {
                    Write-LogMessage "Failed to add to group $Group`: $_" "Warning"
                }
            }
        }
    }

    Write-LogMessage "User creation process completed successfully" "Success"
} catch {
    Write-LogMessage "Error creating user: $_" "Error"
    throw
} finally {
    Write-LogMessage "Script execution finished" "Information"
}
