# =============================================================================
# Script: New-FakePSTFile.ps1
# Created: 2025-06-03 17:05:53 UTC
# Author: jdyer-nuvodia
# Last Updated: 2025-06-03 17:23:00 UTC
# Updated By: jdyer-nuvodia
# Version: 1.0.3
# Additional Info: Fixed PST creation method and improved store mounting reliability
# =============================================================================

<#
.SYNOPSIS
    Generate a new PST file with fake email messages for testing purposes.

.DESCRIPTION
    This script creates a new Outlook PST file and populates it with randomly generated email messages
    for testing, development, or training purposes. The script uses Outlook COM automation to create
    folders within the PST and generate realistic-looking email messages with randomized senders,
    recipients, subjects, bodies, and timestamps. All generated data is fake and safe for testing
    environments without exposing real email data.

.PARAMETER PstPath
    The full path including filename where the new PST file will be created.
    Default: C:\Temp\FakeEmails.pst

.PARAMETER PstDisplayName
    The display name to assign to the PST within Outlook.
    Default: FakeEmailArchive

.PARAMETER FolderNames
    Array of folder names to create within the PST. Multiple folders can be specified.
    Default: @('TestInbox', 'TestSent', 'TestArchive')

.PARAMETER MessagesPerFolder
    Number of fake messages to generate in each folder.
    Default: 50

.PARAMETER DateRangeDays
    Number of days in the past from which to randomly select message dates.
    Default: 60

.PARAMETER WhatIf
    Shows what would happen if the cmdlet runs without actually executing the changes.

.PARAMETER Verbose
    Provides detailed output during script execution.

.EXAMPLE
    .\New-FakePSTFile.ps1
    Creates a PST file at C:\Temp\FakeEmails.pst with default settings.

.EXAMPLE
    .\New-FakePSTFile.ps1 -PstPath "D:\TestData\MyTest.pst" -MessagesPerFolder 100 -Verbose
    Creates a PST file at the specified path with 100 messages per folder and verbose output.

.EXAMPLE
    .\New-FakePSTFile.ps1 -FolderNames @('Inbox', 'Sent Items', 'Archive', 'Projects') -DateRangeDays 365
    Creates a PST with custom folder names and message dates spanning one year.

.EXAMPLE
    .\New-FakePSTFile.ps1 -WhatIf
    Shows what the script would do without actually creating the PST file.

.EXAMPLE
    .\New-FakePSTFile.ps1 -PstPath "C:\Temp\CompanyEmails_2025.pst" -PstDisplayName "Company Test Archive" -FolderNames @('Inbox', 'Sent Items', 'Projects', 'HR Communications', 'IT Support', 'Sales Leads') -MessagesPerFolder 250 -DateRangeDays 365 -Verbose
    Creates a comprehensive test PST file with 6 folders containing 250 messages each (1500 total), spanning one year of message dates, with verbose output for detailed progress tracking. This example demonstrates all available parameters to create a realistic corporate email archive for testing purposes.

.NOTES
    Requires:
    - Microsoft Outlook installed and configured on the local machine
    - PowerShell execution policy that allows script execution
    - Sufficient disk space for the PST file
    
    The script uses Outlook COM automation, so Outlook must be installed but does not need
    to have an active Exchange connection. The script will create a standalone PST file.
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory = $false)]
    [ValidateScript({
        $directory = Split-Path -Path $_ -Parent
        if (-not (Test-Path -Path $directory)) {
            throw "Directory '$directory' does not exist. Please create the directory first or specify a valid path."
        }
        $true
    })]
    [string]$PstPath = "C:\Temp\FakeEmails.pst",

    [Parameter(Mandatory = $false)]
    [ValidateLength(1, 50)]
    [string]$PstDisplayName = "FakeEmailArchive",

    [Parameter(Mandatory = $false)]
    [ValidateCount(1, 20)]
    [string[]]$FolderNames = @('TestInbox', 'TestSent', 'TestArchive'),

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 10000)]
    [int]$MessagesPerFolder = 50,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 3650)]
    [int]$DateRangeDays = 60
)

# Initialize logging
$ScriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ComputerName = $env:COMPUTERNAME
$LogTimestamp = (Get-Date).ToUniversalTime().ToString("yyyyMMdd-HHmmss")
$LogFileName = "${ScriptName}_${ComputerName}_${LogTimestamp}.log"
$LogPath = Join-Path -Path $ScriptPath -ChildPath $LogFileName

# Logging function
function Write-LogMessage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Warning', 'Error', 'Success', 'Debug')]
        [string]$Level = 'Info'
    )
    
    $Timestamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss UTC")
    $LogEntry = "[$Timestamp] [$Level] $Message"
    
    # Write to log file
    Add-Content -Path $LogPath -Value $LogEntry -Encoding UTF8
      # Write to console with appropriate color
    switch ($Level) {
        'Info' { Write-Information $Message -InformationAction Continue }
        'Warning' { Write-Warning $Message }
        'Error' { Write-Error $Message }
        'Success' { Write-Information $Message -InformationAction Continue }
        'Debug' { Write-Verbose $Message }
    }
}

# Sample data arrays for randomization
$SampleRecipients = @(
    'support@contoso.com',
    'marketing@acme.org',
    'user1@example.com',
    'user2@example.com',
    'qa@testdomain.net',
    'admin@fakeco.com',
    'info@samplecorp.com',
    'contact@demo.org',
    'help@placeholder.org',
    'team@testco.com'
)

$SampleSubjects = @(
    'Monthly Report - {0}',
    'Reminder: Meeting scheduled for {0}',
    'Re: Your Request #{0}',
    'Action Required: Policy Update',
    'Welcome to the Team, {0}',
    'System Maintenance Notice - {0}',
    'Invoice #{0} - Payment Due',
    'Project Update: Phase {0} Complete',
    'Security Alert: Account Activity on {0}',
    'Newsletter: Week of {0}'
)

$SampleBodies = @(
    "Hello,`r`n`r`nThis is a test email generated on {0}. Please disregard this message as it is for testing purposes only.`r`n`r`nBest regards,`r`nAutomated Test System",
    "Dear Recipient,`r`n`r`nYour ticket #{0} has been received and is currently being processed by our team. Expected resolution: {1}.`r`n`r`nThank you for your patience,`r`nSupport Team",
    "Team,`r`n`r`nWe wanted to inform you about the upcoming system maintenance scheduled for {0}. Please plan accordingly.`r`n`r`nRegards,`r`nOperations Team",
    "Hello,`r`n`r`nCongratulations! Your order #{0} has been processed and shipped. Expected delivery date: {1}.`r`n`r`nThank you for your business,`r`nSales Team",
    "Dear Team Members,`r`n`r`nPlease find attached the Q{0} performance metrics. Review the data and let me know if you have any questions.`r`n`r`nBest,`r`nManagement",
    "Hi there,`r`n`r`nThis is a friendly reminder about the meeting scheduled for {0}. Please confirm your attendance.`r`n`r`nThanks,`r`nAdmin Assistant",
    "Dear User,`r`n`r`nYour account password will expire on {0}. Please update your password to avoid any service interruption.`r`n`r`nSecurity Team",
    "Hello,`r`n`r`nWe have detected unusual activity on your account on {0}. If this was not you, please contact support immediately.`r`n`r`nSecurity Department"
)

# Helper function to get random item from array
function Get-RandomArrayItem {
    param([array]$Array)
    return $Array | Get-Random
}

# Start logging
Write-LogMessage -Message "Starting New-FakePSTFile script execution" -Level "Info"
Write-LogMessage -Message "PST Path: $PstPath" -Level "Debug"
Write-LogMessage -Message "Display Name: $PstDisplayName" -Level "Debug"
Write-LogMessage -Message "Folders: $($FolderNames -join ', ')" -Level "Debug"
Write-LogMessage -Message "Messages per folder: $MessagesPerFolder" -Level "Debug"

try {
    # Check if target directory exists
    $TargetDirectory = Split-Path -Path $PstPath -Parent
    if (-not (Test-Path -Path $TargetDirectory)) {
        if ($PSCmdlet.ShouldProcess($TargetDirectory, "Create directory")) {
            Write-LogMessage -Message "Creating target directory: $TargetDirectory" -Level "Info"
            New-Item -Path $TargetDirectory -ItemType Directory -Force | Out-Null
        }
    }

    # Check if PST already exists
    if (Test-Path -Path $PstPath) {
        if ($PSCmdlet.ShouldProcess($PstPath, "Remove existing PST file")) {
            Write-LogMessage -Message "Removing existing PST file: $PstPath" -Level "Warning"
            Remove-Item -Path $PstPath -Force
        }
    }

    if ($PSCmdlet.ShouldProcess($PstPath, "Create new PST file with $($FolderNames.Count) folders and $MessagesPerFolder messages per folder")) {
        # Initialize Outlook COM object
        Write-LogMessage -Message "Initializing Outlook COM object" -Level "Info"
        $Outlook = New-Object -ComObject Outlook.Application
        $Namespace = $Outlook.GetNamespace("MAPI")        # Create new PST file
        Write-LogMessage -Message "Creating new PST file: $PstPath" -Level "Info"
        
        # Use Outlook's proper method to create a new PST data file
        # The AddStoreEx method with olStoreDefault (1) creates a new PST if it doesn't exist
        $Namespace.AddStoreEx($PstPath, 1)  # 1 = olStoreDefault (creates new PST)        # Wait for PST to be mounted and available
        Write-LogMessage -Message "Waiting for PST to be mounted in Outlook..." -Level "Info"
        Start-Sleep -Seconds 5

        # Find the newly created store with retry logic
        $FakeStore = $null
        $MaxRetries = 10
        $RetryCount = 0
        
        do {
            $RetryCount++
            Write-LogMessage -Message "Attempting to locate PST store (attempt $RetryCount of $MaxRetries)" -Level "Debug"
            
            foreach ($Store in $Namespace.Stores) {
                Write-LogMessage -Message "Checking store: $($Store.DisplayName) at path: $($Store.FilePath)" -Level "Debug"
                if ($Store.FilePath -ieq $PstPath) {
                    $FakeStore = $Store
                    break
                }
            }
            
            if (-not $FakeStore -and $RetryCount -lt $MaxRetries) {
                Write-LogMessage -Message "PST not found yet, waiting 2 more seconds..." -Level "Debug"
                Start-Sleep -Seconds 2
            }
        } while (-not $FakeStore -and $RetryCount -lt $MaxRetries)

        if (-not $FakeStore) {
            throw "Failed to locate the newly created PST in Outlook stores"
        }        # Set PST display name
        Write-LogMessage -Message "Setting PST display name to: $PstDisplayName" -Level "Info"
        try {
            $FakeStore.DisplayName = $PstDisplayName
        }
        catch {
            Write-LogMessage -Message "Warning: Could not set PST display name. Using default name." -Level "Warning"
        }        # Create folders
        Write-LogMessage -Message "Creating $($FolderNames.Count) folders in PST" -Level "Info"
        $RootFolder = $FakeStore.GetRootFolder()
        $FolderObjects = @{}
        foreach ($FolderName in $FolderNames) {
            Write-LogMessage -Message "Creating folder: $FolderName" -Level "Debug"
            $NewFolder = $RootFolder.Folders.Add($FolderName, 6)  # 6 = olFolderInbox
            $FolderObjects[$FolderName] = $NewFolder
        }

        # Generate fake messages
        $TotalMessages = $MessagesPerFolder * $FolderNames.Count
        Write-LogMessage -Message "Generating $TotalMessages fake messages ($MessagesPerFolder per folder)" -Level "Info"
        
        $MessageCount = 0
        for ($i = 1; $i -le $MessagesPerFolder; $i++) {
            foreach ($FolderName in $FolderNames) {
                $MessageCount++
                if ($MessageCount % 10 -eq 0) {
                    Write-LogMessage -Message "Generated $MessageCount of $TotalMessages messages" -Level "Debug"
                }

                $Folder = $FolderObjects[$FolderName]
                
                # Create new mail item
                $Mail = $Outlook.CreateItem(0)  # 0 = olMailItem                # Generate random data
                $ToAddress = Get-RandomArrayItem -Array $SampleRecipients
                
                # Generate subject with placeholder
                $SubjectTemplate = Get-RandomArrayItem -Array $SampleSubjects
                $SubjectPlaceholder = Get-Random -Minimum 1000 -Maximum 9999
                $Subject = $SubjectTemplate -f $SubjectPlaceholder

                # Generate random date within specified range
                $RandomDate = (Get-Date).AddDays(-1 * (Get-Random -Minimum 0 -Maximum $DateRangeDays))
                $FormattedDate = $RandomDate.ToString("MM/dd/yyyy")
                
                # Generate body with placeholders
                $BodyTemplate = Get-RandomArrayItem -Array $SampleBodies
                $BodyPlaceHolder1 = Get-Random -Minimum 100000 -Maximum 999999
                $BodyPlaceHolder2 = (Get-Date).AddDays((Get-Random -Minimum 1 -Maximum 30)).ToString("MM/dd/yyyy")
                $Body = $BodyTemplate -f $FormattedDate, $BodyPlaceHolder1, $BodyPlaceHolder2

                # Set mail properties
                $Mail.Subject = $Subject
                $Mail.Body = $Body
                $Mail.To = $ToAddress

                # Set received time using property accessor
                $PropertyAccessor = $Mail.PropertyAccessor
                $PR_CLIENT_SUBMIT_TIME = "http://schemas.microsoft.com/mapi/proptag/0x00390040"
                $PropertyAccessor.SetProperty($PR_CLIENT_SUBMIT_TIME, $RandomDate.ToFileTime())

                # Move to folder and save
                $Mail.Move($Folder) | Out-Null
                $Mail.Save()
                
                # Release COM object
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Mail) | Out-Null
            }
        }

        Write-LogMessage -Message "Successfully generated all $TotalMessages fake messages" -Level "Success"

        # Clean up: remove store from Outlook session
        Write-LogMessage -Message "Dismounting PST from Outlook session" -Level "Info"
        $Namespace.RemoveStore($RootFolder)

        # Release COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($RootFolder) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($FakeStore) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Namespace) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null        Write-LogMessage -Message "Fake PST file created successfully at: $PstPath" -Level "Success"
        Write-Information "`nSUCCESS: Fake PST file created successfully!" -InformationAction Continue
        Write-Information "Location: $PstPath" -InformationAction Continue
        Write-Information "Folders created: $($FolderNames.Count)" -InformationAction Continue
        Write-Information "Total messages: $TotalMessages" -InformationAction Continue
        Write-Information "Log file: $LogPath" -InformationAction Continue
    }
}
catch {
    $ErrorMessage = "Error creating fake PST file: $($_.Exception.Message)"
    Write-LogMessage -Message $ErrorMessage -Level "Error"
    Write-LogMessage -Message "Stack trace: $($_.ScriptStackTrace)" -Level "Error"
    
    # Clean up COM objects in case of error
    try {
        if ($Outlook) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
        }
    }
    catch {
        Write-LogMessage -Message "Error cleaning up COM objects: $($_.Exception.Message)" -Level "Warning"
    }
    
    throw $_
}
finally {
    Write-LogMessage -Message "Script execution completed" -Level "Info"
    
    # Force garbage collection to clean up COM objects
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
