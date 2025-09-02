# =============================================================================
# Script: Copy-ADUser.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-17 17:00:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.3.0
# Additional Info: Converted Write-Host to Write-ColorOutput for PSScriptAnalyzer compliance
# =============================================================================

<#
.SYNOPSIS
    Copies an existing AD user's group memberships to a new user account.
.DESCRIPTION
    This script creates a new Active Directory user account and copies all group
    memberships from a specified source user. The script:
    - Creates new user with specified properties
    - Copies group memberships from source user
    - Enables the account with a specified password
    Dependencies:
    - Active Directory PowerShell module
    - Domain Admin or appropriate AD delegation rights
.PARAMETER SourceUser
    The username of the existing AD user to copy from
.PARAMETER NewUserName
    The new username to be created
.PARAMETER NewUserGivenName
    The given name for the new user
.PARAMETER NewUserSurname
    The surname for the new user
.PARAMETER NewUserPassword
    The initial password for the new user (SecureString)
.PARAMETER NewUserDescription
    The description for the new user account
.EXAMPLE
    $SecurePass = ConvertTo-SecureString "P@ssw0rd123!" -AsPlainText -Force
    .\Copy-ADUser.ps1 -SourceUser "john.doe" -NewUserName "jane.doe" -NewUserGivenName "Jane" -NewUserSurname "Doe" -NewUserPassword $SecurePass -NewUserDescription "Sales Department"
.NOTES
    Security Level: High
    Required Permissions: Domain Admin or delegated AD user creation rights
    Validation Requirements: Verify source user exists, new username doesn't exist
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$SourceUser,

    [Parameter(Mandatory = $true)]
    [string]$NewUserName,

    [Parameter(Mandatory = $true)]
    [string]$NewUserGivenName,

    [Parameter(Mandatory = $true)]
    [string]$NewUserSurname,

    [Parameter(Mandatory = $true)]
    [SecureString]$NewUserPassword,

    [Parameter(Mandatory = $true)]
    [string]$NewUserDescription
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


# Load the Active Directory module
Import-Module ActiveDirectory

# Verify source user exists
try {
    Write-ColorOutput -Message "Verifying source user exists..." -Color 'Cyan'
    $sourceUserDetails = Get-ADUser -Identity $SourceUser -Properties * -ErrorAction Stop
} catch {
    Write-Error "Source user '$SourceUser' not found. Please verify the username and try again."
    exit 1
}

# Verify new username doesn't exist
if (Get-ADUser -Filter "SamAccountName -eq '$NewUserName'" -ErrorAction SilentlyContinue) {
    Write-Error "User '$NewUserName' already exists. Please choose a different username."
    exit 1
}

Write-ColorOutput -Message "Creating new user account..." -Color 'Cyan'

# Create the new user with the different name properties and description
New-ADUser `
    -Name "$NewUserGivenName $NewUserSurname" `
    -GivenName $NewUserGivenName `
    -Surname $NewUserSurname `
    -SamAccountName $NewUserName `
    -UserPrincipalName "$NewUserName@$(($sourceUserDetails.UserPrincipalName).Split('@')[1])" `
    -Path $sourceUserDetails.DistinguishedName `
    -Enabled $true `
    -AccountPassword $NewUserPassword `
    -Description $NewUserDescription

# Add the new user to the same groups as the source user
$sourceUserGroups = Get-ADUser -Identity $SourceUser -Properties MemberOf | Select-Object -ExpandProperty MemberOf
foreach ($group in $sourceUserGroups) {
    try {
        Add-ADGroupMember -Identity $group -Members $NewUserName
        Write-Output "Added $NewUserName to group $group"
    } catch {
        Write-Warning "Failed to add user to group $group"
    }
}

Write-ColorOutput -Message "New user $NewUserName created successfully!" -Color 'Green'
Write-ColorOutput -Message "Group memberships copied from $SourceUser" -Color 'Green'
