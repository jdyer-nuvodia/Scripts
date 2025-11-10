# =============================================================================
# Script: Add-ScriptsToProfile.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-11-10 23:53:59 UTC
# Updated By: GitHub Copilot
# Version: 1.0.8
# Additional Info: Normalized profile script paths to UNC for mapped drive compatibility
# =============================================================================

<#
.SYNOPSIS
Adds command shortcuts for repository scripts into the user PowerShell profile.

.DESCRIPTION
Generates or updates a managed region within the user PowerShell profile that maps each
PowerShell script in the specified directory to a global function. The script enforces
logging, idempotent updates, duplicate detection, and organizational formatting
requirements while supporting WhatIf semantics.

.PARAMETER ScriptDirectory
UNC path or local directory that contains PowerShell scripts to expose as profile functions.

.PARAMETER ProfilePath
Target PowerShell profile file to update. Defaults to the current user profile for the
running host.

.PARAMETER IncludeSubdirectories
When provided, scripts in child folders of the target directory are also exposed.

.PARAMETER ProfileRegionName
Logical name for the managed profile region. Defaults to "Scripts AutoLoad".

.PARAMETER LogDirectory
Directory where execution logs are stored. Defaults to the script directory Logs folder.

.EXAMPLE
.\Add-ScriptsToProfile.ps1 -ScriptDirectory "C:\Users\jdyer\OneDrive - Nuvodia\Documents\Github\Scripts" -IncludeSubdirectories -WhatIf
Demonstrates the profile update and logging actions without modifying the profile.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ScriptDirectory = '\\WYK-FP01\HomeFolders$\aegis\Documents\Github\Scripts',

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ProfilePath = $PROFILE,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeSubdirectories,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ProfileRegionName = 'Scripts AutoLoad',

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$LogDirectory
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not $PSBoundParameters.ContainsKey('LogDirectory') -or [string]::IsNullOrWhiteSpace($LogDirectory)) {
    $LogDirectory = Join-Path -Path $PSScriptRoot -ChildPath 'Logs'
}

$script:UseAnsiColors = $PSVersionTable.PSVersion.Major -ge 7
$ansiEscapeCharacter = [char]27
$script:Colors = @{
    White    = if ($script:UseAnsiColors) { "$ansiEscapeCharacter[37m" } else { [ConsoleColor]::White }
    Cyan     = if ($script:UseAnsiColors) { "$ansiEscapeCharacter[36m" } else { [ConsoleColor]::Cyan }
    Green    = if ($script:UseAnsiColors) { "$ansiEscapeCharacter[32m" } else { [ConsoleColor]::Green }
    Yellow   = if ($script:UseAnsiColors) { "$ansiEscapeCharacter[33m" } else { [ConsoleColor]::Yellow }
    Red      = if ($script:UseAnsiColors) { "$ansiEscapeCharacter[31m" } else { [ConsoleColor]::Red }
    Magenta  = if ($script:UseAnsiColors) { "$ansiEscapeCharacter[35m" } else { [ConsoleColor]::Magenta }
    DarkGray = if ($script:UseAnsiColors) { "$ansiEscapeCharacter[90m" } else { [ConsoleColor]::DarkGray }
    Reset    = if ($script:UseAnsiColors) { "$ansiEscapeCharacter[0m" } else { '' }
}

function Write-ColorOutput {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$Color = 'White'
    )

    if ($script:UseAnsiColors) {
        $colorCode = $script:Colors[$Color]
        $resetCode = $script:Colors.Reset
        Write-Output "${colorCode}${Message}${resetCode}"
    } else {
        $originalColor = $Host.UI.RawUI.ForegroundColor
        try {
            if ($script:Colors[$Color] -and $script:Colors[$Color] -ne '') {
                $Host.UI.RawUI.ForegroundColor = $script:Colors[$Color]
            }
            Write-Output $Message
        } finally {
            $Host.UI.RawUI.ForegroundColor = $originalColor
        }
    }
}

function New-LogFile {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$DestinationDirectory
    )

    if (-not (Test-Path -Path $DestinationDirectory -PathType Container)) {
        if ($PSCmdlet.ShouldProcess($DestinationDirectory, 'Create log directory')) {
            [void](New-Item -Path $DestinationDirectory -ItemType Directory -Force)
        } else {
            return $null
        }
    }

    $timestampSuffix = (Get-Date).ToUniversalTime().ToString('yyyyMMddHHmmss')
    $fileName = '{0}_{1}_UTC.log' -f $env:COMPUTERNAME, $timestampSuffix
    $logPath = Join-Path -Path $DestinationDirectory -ChildPath $fileName

    if ($PSCmdlet.ShouldProcess($logPath, 'Create log file')) {
        [void](New-Item -Path $logPath -ItemType File -Force)
        return $logPath
    }

    return $null
}

$script:LogFilePath = $null

function Write-LogEntry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Debug')]
        [string]$Severity = 'Info'
    )

    $timestamp = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $entry = '{0}`t{1}`t{2}' -f $timestamp, $Severity, $Message

    if (-not [string]::IsNullOrEmpty($script:LogFilePath)) {
        Add-Content -Path $script:LogFilePath -Value $entry -Encoding UTF8
    }

    switch ($Severity) {
        'Success' { Write-ColorOutput -Message $Message -Color 'Green' }
        'Warning' { Write-ColorOutput -Message $Message -Color 'Yellow' }
        'Error' { Write-ColorOutput -Message $Message -Color 'Red' }
        'Debug' { Write-ColorOutput -Message $Message -Color 'Magenta' }
        default { Write-ColorOutput -Message $Message -Color 'White' }
    }
}

function Get-ScriptFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$RootPath,
        [Parameter(Mandatory = $false)]
        [switch]$IncludeSubdirectories
    )

    $parameters = @{
        Path        = $RootPath
        Filter      = '*.ps1'
        File        = $true
        ErrorAction = 'Stop'
    }

    if ($IncludeSubdirectories.IsPresent) {
        $parameters['Recurse'] = $true
    }

    return Get-ChildItem @parameters | Sort-Object -Property FullName
}

function Get-ScriptDuplicateSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$ScriptFiles
    )

    $duplicateGroups = $ScriptFiles | Group-Object -Property BaseName | Where-Object -FilterScript { $_.Count -gt 1 }
    foreach ($group in $duplicateGroups) {
        $paths = $group.Group | Sort-Object -Property FullName | ForEach-Object -Process { $_.FullName }
        $message = 'Duplicate script name detected: {0}. Using last enumerated instance when profile loads. Paths: {1}' -f $group.Name, ([string]::Join('; ', $paths))
        Write-LogEntry -Message $message -Severity 'Warning'
    }
}

function Get-ProfileRegionContent {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$RegionName,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$TargetDirectory,
        [Parameter(Mandatory = $false)]
        [switch]$IncludeSubdirectories
    )

    $escapedDirectory = $TargetDirectory.Replace('"', '`"')
    $builder = [System.Text.StringBuilder]::new()
    [void]$builder.AppendLine("#region $RegionName - Managed by Add-ScriptsToProfile.ps1")
    [void]$builder.AppendLine('$script:ScriptsRoot = "' + $escapedDirectory + '"')
    [void]$builder.AppendLine('if (-not (Test-Path -Path $script:ScriptsRoot -PathType Container)) {')
    [void]$builder.AppendLine('    return')
    [void]$builder.AppendLine('}')
    [void]$builder.AppendLine('$scriptLoaderParameters = @{')
    [void]$builder.AppendLine('    Path = $script:ScriptsRoot')
    [void]$builder.AppendLine('    Filter = ''*.ps1''')
    [void]$builder.AppendLine('    File = $true')
    [void]$builder.AppendLine('    ErrorAction = ''Stop''')
    [void]$builder.AppendLine('}')
    if ($IncludeSubdirectories.IsPresent) {
        [void]$builder.AppendLine('$scriptLoaderParameters.Recurse = $true')
    }

    [void]$builder.AppendLine('Get-ChildItem @scriptLoaderParameters | Sort-Object -Property FullName | ForEach-Object -Process {')
    [void]$builder.AppendLine('    $functionName = $_.BaseName')
    [void]$builder.AppendLine('    $scriptPath = $_.FullName')
    [void]$builder.AppendLine('    if ($_.PSDrive -and -not [string]::IsNullOrEmpty($_.PSDrive.DisplayRoot)) {')
    [void]$builder.AppendLine('        $driveRoot = $_.PSDrive.Root')
    [void]$builder.AppendLine('        if ($scriptPath.StartsWith($driveRoot, [System.StringComparison]::OrdinalIgnoreCase)) {')
    [void]$builder.AppendLine('            $relativePath = $scriptPath.Substring($driveRoot.Length).TrimStart([char]92)')
    [void]$builder.AppendLine('            $scriptPath = Join-Path -Path $_.PSDrive.DisplayRoot.TrimEnd([char]92) -ChildPath $relativePath')
    [void]$builder.AppendLine('        }')
    [void]$builder.AppendLine('    }')
    [void]$builder.AppendLine("    `$escapedPath = `$scriptPath.Replace(""'"", ""''"")")
    [void]$builder.AppendLine("    `$scriptBlockText = ""& '`$escapedPath' @args""")
    [void]$builder.AppendLine('    $invocationBlock = [ScriptBlock]::Create($scriptBlockText)')
    [void]$builder.AppendLine('    $targetPath = "Function:\\global:" + $functionName')
    [void]$builder.AppendLine('    Set-Item -Path $targetPath -Value $invocationBlock -Force')
    [void]$builder.AppendLine('    $extensionTarget = "Function:\\global:" + $functionName + ".ps1"')
    [void]$builder.AppendLine('    Set-Item -Path $extensionTarget -Value $invocationBlock -Force')
    [void]$builder.AppendLine('}')
    [void]$builder.AppendLine("#endregion $RegionName - Managed by Add-ScriptsToProfile.ps1")

    return $builder.ToString().TrimEnd() + [Environment]::NewLine
}

function Backup-ProfileFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ProfilePath
    )

    if (-not (Test-Path -Path $ProfilePath -PathType Leaf)) {
        return $null
    }

    $timestampSuffix = (Get-Date).ToUniversalTime().ToString('yyyyMMddHHmmss')
    $directory = Split-Path -Path $ProfilePath -Parent
    $fileName = '{0}.{1}.bak' -f (Split-Path -Path $ProfilePath -Leaf), $timestampSuffix
    $destination = Join-Path -Path $directory -ChildPath $fileName
    Copy-Item -Path $ProfilePath -Destination $destination -Force
    return $destination
}

function Update-ProfileFile {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ProfilePath,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$RegionName,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$RegionContent
    )

    $profileDirectory = Split-Path -Path $ProfilePath -Parent
    if (-not (Test-Path -Path $profileDirectory -PathType Container)) {
        if ($PSCmdlet.ShouldProcess($profileDirectory, 'Create profile directory')) {
            [void](New-Item -Path $profileDirectory -ItemType Directory -Force)
        } else {
            Write-LogEntry -Message ('Profile directory creation skipped for {0} due to WhatIf preference.' -f $profileDirectory) -Severity 'Info'
            return
        }
    }

    if (-not (Test-Path -Path $ProfilePath -PathType Leaf)) {
        if ($PSCmdlet.ShouldProcess($ProfilePath, 'Create profile file')) {
            [void](New-Item -Path $ProfilePath -ItemType File -Force)
        } else {
            Write-LogEntry -Message ('Profile file creation skipped for {0} due to WhatIf preference.' -f $ProfilePath) -Severity 'Info'
            return
        }
    }

    $profileContent = Get-Content -Path $ProfilePath -Raw -Encoding UTF8
    if ($null -eq $profileContent) {
        $profileContent = [string]::Empty
    }
    $regionStart = "#region $RegionName - Managed by Add-ScriptsToProfile.ps1"
    $regionEnd = "#endregion $RegionName - Managed by Add-ScriptsToProfile.ps1"
    $regexPattern = "(?s){0}.*?{1}" -f [System.Text.RegularExpressions.Regex]::Escape($regionStart), [System.Text.RegularExpressions.Regex]::Escape($regionEnd)
    $match = [System.Text.RegularExpressions.Regex]::Match($profileContent, $regexPattern)

    $normalizedNewContent = $RegionContent.Trim()

    if ($match.Success) {
        $existingContent = $match.Value.Trim()
        if ($existingContent -eq $normalizedNewContent) {
            Write-LogEntry -Message 'Profile already contains the latest script loader region. No action required.' -Severity 'Info'
            return
        }

        if ($PSCmdlet.ShouldProcess($ProfilePath, 'Update managed script loader region')) {
            $updatedContent = [System.Text.RegularExpressions.Regex]::Replace($profileContent, $regexPattern, $RegionContent)
            $backupPath = Backup-ProfileFile -ProfilePath $ProfilePath
            Set-Content -Path $ProfilePath -Value $updatedContent -Encoding UTF8
            $successMessage = if ($backupPath) { 'Profile updated. Backup saved to {0}' -f $backupPath } else { 'Profile updated.' }
            Write-LogEntry -Message $successMessage -Severity 'Success'
        } else {
            Write-LogEntry -Message 'Profile update skipped due to WhatIf preference.' -Severity 'Info'
        }
    } else {
        if ($PSCmdlet.ShouldProcess($ProfilePath, 'Add managed script loader region')) {
            $contentToWrite = if ([string]::IsNullOrWhiteSpace($profileContent)) {
                $RegionContent
            } else {
                $profileContent.TrimEnd() + [Environment]::NewLine + [Environment]::NewLine + $RegionContent
            }

            $backupPath = Backup-ProfileFile -ProfilePath $ProfilePath
            Set-Content -Path $ProfilePath -Value $contentToWrite -Encoding UTF8
            $successMessage = if ($backupPath) { 'Profile updated with new region. Backup saved to {0}' -f $backupPath } else { 'Profile updated with new region.' }
            Write-LogEntry -Message $successMessage -Severity 'Success'
        } else {
            Write-LogEntry -Message 'Profile update skipped due to WhatIf preference.' -Severity 'Info'
        }
    }
}

$script:ExitCode = 1

try {
    $script:LogFilePath = New-LogFile -DestinationDirectory $LogDirectory -WhatIf:$WhatIfPreference
    Write-LogEntry -Message ('Log initialized at {0}' -f $script:LogFilePath) -Severity 'Debug'
    Write-LogEntry -Message ('Starting Add-ScriptsToProfile for {0}' -f $ScriptDirectory) -Severity 'Info'

    if (-not (Test-Path -Path $ScriptDirectory -PathType Container)) {
        throw [System.IO.DirectoryNotFoundException]::new('Script directory not found: ' + $ScriptDirectory)
    }

    $scriptFiles = Get-ScriptFile -RootPath $ScriptDirectory -IncludeSubdirectories:$IncludeSubdirectories
    if (-not $scriptFiles -or $scriptFiles.Count -eq 0) {
        Write-LogEntry -Message 'No PowerShell scripts were found in the specified directory. Profile region will still be created for future use.' -Severity 'Warning'
    } else {
        Get-ScriptDuplicateSummary -ScriptFiles $scriptFiles
        Write-LogEntry -Message ('Discovered {0} PowerShell script(s) for profile exposure.' -f $scriptFiles.Count) -Severity 'Info'
    }

    $regionContent = Get-ProfileRegionContent -RegionName $ProfileRegionName -TargetDirectory $ScriptDirectory -IncludeSubdirectories:$IncludeSubdirectories
    Update-ProfileFile -ProfilePath $ProfilePath -RegionName $ProfileRegionName -RegionContent $regionContent -WhatIf:$WhatIfPreference
    $script:ExitCode = 0
    Write-LogEntry -Message 'Add-ScriptsToProfile execution completed successfully.' -Severity 'Success'
} catch {
    $script:ExitCode = 1
    Write-LogEntry -Message ('[SYSTEM ERROR DETECTED] {0}' -f $_.Exception.Message) -Severity 'Error'
    if ($_.InvocationInfo -and $null -ne $_.InvocationInfo.PositionMessage) {
        Write-LogEntry -Message $_.InvocationInfo.PositionMessage -Severity 'Debug'
    }
} finally {
    if (-not [string]::IsNullOrEmpty($script:LogFilePath)) {
        Write-LogEntry -Message ('Execution completed with exit code {0}' -f $script:ExitCode) -Severity 'Info'
        Write-LogEntry -Message ('Log file saved to {0}' -f $script:LogFilePath) -Severity 'Debug'
    }

    exit $script:ExitCode
}
