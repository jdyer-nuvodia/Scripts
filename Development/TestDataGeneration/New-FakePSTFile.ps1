# =============================================================================
# Script: New-FakePSTFile.ps1
# Author: jdyer-nuvodia
# Last Updated: 2025-07-16 20:48:00 UTC
# Updated By: jdyer-nuvodia
# Version: 2.9.4
# Additional Info: Fixed header corruption and PSScriptAnalyzer syntax issues
# =============================================================================

<#
.SYNOPSIS
Creates a new PST file with random test data for development and testing purposes.

.DESCRIPTION
This script creates a new Outlook PST file at a specified location and populates it with random email data.
The script creates an Inbox folder and adds a configurable number of test emails with random subjects and bodies.
This is useful for testing applications that process PST files or for creating sample data for development.
The script includes -WhatIf functionality to preview actions before execution and comprehensive logging.

PREREQUISITES:
- Microsoft Outlook must be installed and properly configured
- Outlook Interop Assembly must be available
- Script must be run with appropriate permissions to create files

The script will automatically check for these prerequisites before attempting to create the PST file.

.PARAMETER PSTPath
The full path where the new PST file will be created. Default is "C:\Temp\RandomData.pst"

.PARAMETER EmailCount
The total number of random emails to create and distribute across all folders in the PST file. Default is 10

.PARAMETER WhatIf
Shows what would happen if the cmdlet runs without actually executing the operations

.PARAMETER AttachmentPercentage
The percentage of emails that should include random attachments (0-100). Default is 30

.PARAMETER AttachmentTempPath
The temporary directory path where attachment files will be created before being added to emails.
Files are automatically cleaned up after use. Default is "C:\Temp\PST_Attachments"

.PARAMETER Verbose
Enables verbose output for detailed operation information

.EXAMPLE
.\New-FakePSTFile.ps1
Creates a PST file at the default location with 10 random emails distributed across folders

.EXAMPLE
.\New-FakePSTFile.ps1 -AttachmentPercentage 50
Creates a PST file with 50% of emails containing random attachments (2-5MB each)

.EXAMPLE
.\New-FakePSTFile.ps1 -PSTPath "C:\Temp\Sample.pst" -EmailCount 1000 -AttachmentPercentage 25
Creates a PST file with 1000 emails, where 25% include random attachments

.EXAMPLE
.\New-FakePSTFile.ps1 -WhatIf
Shows what would happen without actually creating the PST file
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $false)]
    [string]$PSTPath = "C:\Temp\RandomData.pst",

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 1000000)]
    [int]$EmailCount = 10,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 100)]
    [int]$AttachmentPercentage = 30,

    [Parameter(Mandatory = $false)]
    [string]$AttachmentTempPath = "C:\Temp\PST_Attachments"
)

# Initialize variables
$scriptName = $MyInvocation.MyCommand.Name
$scriptPath = $MyInvocation.MyCommand.Path | Split-Path -Parent
$systemName = $env:COMPUTERNAME
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$logFileName = "{0}_{1}_{2}_UTC.log" -f $scriptName.Replace('.ps1', ''), $systemName, $timestamp
$logPath = Join-Path -Path $scriptPath -ChildPath $logFileName
$startTime = Get-Date

# Function to write colored output and log
function Write-ColoredOutput {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [switch]$NoLog
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC"
    $logMessage = "[$timestamp] $Message"

    # Write to console
    Write-Output $Message

    if (-not $NoLog) {
        Add-Content -Path $logPath -Value $logMessage -ErrorAction SilentlyContinue
    }
}

# Function to create a random file with specified size
function New-RandomFile {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $true)]
        [long]$SizeInBytes
    )
    try {
        # Ensure directory exists
        $directory = Split-Path -Path $FilePath -Parent
        if (-not (Test-Path -Path $directory)) {
            if ($PSCmdlet.ShouldProcess($directory, "Create directory")) {
                New-Item -Path $directory -ItemType Directory -Force | Out-Null
            } else {
                return $false
            }
        }

        if ($PSCmdlet.ShouldProcess($FilePath, "Create random file ($([Math]::Round($SizeInBytes / 1MB, 2)) MB)")) {
            # Generate random content efficiently for large files
            # 64KB buffer for efficiency
            $buffer = New-Object byte[] 65536
            $random = New-Object System.Random
            $remainingBytes = $SizeInBytes

            $fileStream = [System.IO.File]::Create($FilePath)

            while ($remainingBytes -gt 0) {
                $bytesToWrite = [Math]::Min($buffer.Length, $remainingBytes)

                if ($bytesToWrite -eq $buffer.Length) {
                    $random.NextBytes($buffer)
                    $fileStream.Write($buffer, 0, $buffer.Length)
                } else {
                    $smallBuffer = New-Object byte[] $bytesToWrite
                    $random.NextBytes($smallBuffer)
                    $fileStream.Write($smallBuffer, 0, $smallBuffer.Length)
                }

                $remainingBytes -= $bytesToWrite
            }

            $fileStream.Close()
            $fileStream.Dispose()

            return $true
        } else {
            return $false
        }
    } catch {
        Write-ColoredOutput -Message "WARNING: Failed to create random file '$FilePath': $($_.Exception.Message)"
        return $false
    }
}

# Function to generate a random attachment for an email
function New-RandomAttachment {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TempDirectory
    )

    try {
        # Random file types that are commonly attached to emails
        $fileTypes = @(
            @{ Extension = ".pdf"; Name = "Document" },
            @{ Extension = ".docx"; Name = "Report" },
            @{ Extension = ".xlsx"; Name = "Spreadsheet" },
            @{ Extension = ".pptx"; Name = "Presentation" },
            @{ Extension = ".txt"; Name = "Notes" },
            @{ Extension = ".jpg"; Name = "Image" },
            @{ Extension = ".png"; Name = "Photo" },
            @{ Extension = ".zip"; Name = "Archive" },
            @{ Extension = ".log"; Name = "LogFile" },
            @{ Extension = ".csv"; Name = "DataFile" }
        )

        # Select random file type
        $selectedType = $fileTypes | Get-Random

        # Generate random size between 2MB and 5MB
        # 2MB
        $minSizeBytes = 2 * 1024 * 1024
        # 5MB
        $maxSizeBytes = 5 * 1024 * 1024
        $randomSize = Get-Random -Minimum $minSizeBytes -Maximum $maxSizeBytes

        # Generate unique filename with timestamp
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss_fff"
        $randomNumber = Get-Random -Minimum 1000 -Maximum 9999
        $fileName = "$($selectedType.Name)_$timestamp`_$randomNumber$($selectedType.Extension)"
        # Create the random file
        $fullPath = Join-Path -Path $TempDirectory -ChildPath $fileName
        $success = New-RandomFile -FilePath $fullPath -SizeInBytes $randomSize

        if ($success) {
            return @{
                FilePath = $fullPath
                FileName = $fileName
                SizeMB = [Math]::Round($randomSize / 1MB, 2)
                Success = $true
            }
        } else {
            return @{ Success = $false }
        }
    } catch {
        Write-ColoredOutput -Message "WARNING: Failed to generate random attachment: $($_.Exception.Message)"
        return @{ Success = $false }
    }
}

# Function to test Outlook prerequisite
function Test-OutlookPrerequisite {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $false)]
        [bool]$SkipCOMTest = $false
    )

    Write-ColoredOutput -Message "Checking Outlook prerequisites..."

    # Comprehensive check for Outlook installation in Program Files directories
    $programFilesPaths = @(
        $env:ProgramFiles,
        ${ env:ProgramFiles(x86)}
    )

    $outlookPaths = @()

    # Build comprehensive list of potential Outlook paths
    foreach ($programPath in $programFilesPaths) {
        if ($programPath -and (Test-Path -Path $programPath)) {
            # Office versions and their typical paths
            $officeVersions = @(
                # Office 2016/2019/2021/365 Click-to-Run
                "Microsoft Office\root\Office16",
                # Office 2016 MSI
                "Microsoft Office\Office16",
                # Office 2013 Click-to-Run
                "Microsoft Office\root\Office15",
                # Office 2013 MSI
                "Microsoft Office\Office15",
                # Office 2010 Click-to-Run
                "Microsoft Office\root\Office14",
                # Office 2010 MSI
                "Microsoft Office\Office14",
                # Office 2007
                "Microsoft Office\OFFICE12",
                # Office 2003
                "Microsoft Office\OFFICE11"
            )

            foreach ($version in $officeVersions) {
                $fullPath = Join-Path -Path $programPath -ChildPath "$version\OUTLOOK.EXE"
                $outlookPaths += $fullPath
            }
        }
    }

    # Also check for standalone Outlook installations
    foreach ($programPath in $programFilesPaths) {
        if ($programPath -and (Test-Path -Path $programPath)) {
            $standalonePaths = @(
                "Microsoft Outlook\OUTLOOK.EXE",
                "Outlook\OUTLOOK.EXE"
            )

            foreach ($standalone in $standalonePaths) {
                $fullPath = Join-Path -Path $programPath -ChildPath $standalone
                $outlookPaths += $fullPath
            }
        }
    }

    # Check each potential path
    $outlookInstalled = $false

    Write-ColoredOutput -Message "Searching for Outlook executable in Program Files directories..."

    foreach ($path in $outlookPaths) {
        Write-Verbose "Checking path: $path"
        if (Test-Path -Path $path -PathType Leaf) {
            $outlookInstalled = $true
            $fileInfo = Get-Item -Path $path
            Write-ColoredOutput -Message "Found Outlook installation at: $path"
            Write-ColoredOutput -Message "Version: $($fileInfo.VersionInfo.FileVersion)"
            Write-ColoredOutput -Message "Product: $($fileInfo.VersionInfo.ProductName)"
            break
        }
    }

    if (-not $outlookInstalled) {
        Write-ColoredOutput -Message "WARNING: Outlook executable not found in standard Program Files locations"
        Write-ColoredOutput -Message "Checked paths in:"
        Write-ColoredOutput -Message "  - $env:ProgramFiles"
        if ($env:ProgramFiles -ne ${ env:ProgramFiles(x86)}) {
            Write-ColoredOutput -Message "  - ${ env:ProgramFiles(x86)}"
        }
    }

    # Check Windows Registry for Outlook installation
    Write-ColoredOutput -Message "Checking Windows Registry for Outlook registration..."

    $registryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
        "HKLM:\SOFTWARE\Microsoft\Office\16.0\Outlook",
        "HKLM:\SOFTWARE\Microsoft\Office\15.0\Outlook",
        "HKLM:\SOFTWARE\Microsoft\Office\14.0\Outlook",
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Outlook",
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\Outlook",
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\14.0\Outlook"
    )

    $registryFound = $false
    foreach ($regPath in $registryPaths) {
        if (Test-Path -Path $regPath) {
            $registryFound = $true
            Write-ColoredOutput -Message "Found Outlook registry entry at: $regPath"
            break
        }
    }

    if (-not $registryFound) {
        Write-ColoredOutput -Message "WARNING: No Outlook registry entries found"
    }

    # Test COM interface if not skipped
    if (-not $SkipCOMTest) {
        Write-ColoredOutput -Message "Testing Outlook COM interface..."

        try {
            $testOutlook = New-Object -ComObject Outlook.Application
            Write-ColoredOutput -Message "Outlook COM interface test successful"

            # Clean up test object
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($testOutlook) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()

            return $true
        } catch {
            Write-ColoredOutput -Message "ERROR: Outlook COM interface test failed: $($_.Exception.Message)"
            return $false
        }
    } else {
        Write-ColoredOutput -Message "COM interface test skipped as requested"
        return ($outlookInstalled -or $registryFound)
    }
}

# Function to generate random string
function Get-RandomString {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $false)]
        [int]$Length = 8
    )

    $chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    $randomString = ""

    for ($i = 1; $i -le $Length; $i++) {
        $randomString += $chars[(Get-Random -Maximum $chars.Length)]
    }

    return $randomString
}

# Main function to create PST file with data
function New-PSTFileWithData {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([void])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $true)]
        [int]$EmailCount,

        [Parameter(Mandatory = $false)]
        [bool]$IsWhatIfMode = $false,

        [Parameter(Mandatory = $false)]
        [int]$AttachmentPercentage = 30,

        [Parameter(Mandatory = $false)]
        [string]$AttachmentTempPath = "C:\Temp\PST_Attachments"
    )

    try {
        if (-not $IsWhatIfMode) {
            if ($PSCmdlet.ShouldProcess($FilePath, "Create PST file with $EmailCount emails")) {
                Write-ColoredOutput -Message "Creating PST file with $EmailCount emails at: $FilePath"

                # Initialize attachment generation if needed
                $attachmentTempFiles = @()
                if ($AttachmentPercentage -gt 0) {
                    Write-ColoredOutput -Message "Attachment generation enabled: $AttachmentPercentage% of emails will include random attachments (2-5MB each)"

                    # Ensure temp directory exists
                    if (-not (Test-Path -Path $AttachmentTempPath)) {
                        Write-ColoredOutput -Message "Creating temporary attachment directory: $AttachmentTempPath"
                        New-Item -Path $AttachmentTempPath -ItemType Directory -Force | Out-Null
                    }
                } else {
                    Write-ColoredOutput -Message "No attachments will be generated (AttachmentPercentage = 0)"
                }

                # Test Outlook prerequisites first
                if (-not (Test-OutlookPrerequisite)) {
                    throw "Outlook prerequisites not met. Please install and configure Microsoft Outlook."
                }

                # Delete existing file if it exists
                if (Test-Path -Path $FilePath) {
                    Write-ColoredOutput -Message "Removing existing PST file..."
                    Remove-Item -Path $FilePath -Force
                }

                # Ensure target directory exists
                $targetDir = Split-Path -Path $FilePath -Parent
                if (-not (Test-Path -Path $targetDir)) {
                    Write-ColoredOutput -Message "Creating target directory: $targetDir"
                    New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
                }

                # Create Outlook application object
                Write-ColoredOutput -Message "Initializing Outlook COM interface..."
                $outlook = New-Object -ComObject Outlook.Application

                Write-ColoredOutput -Message "Outlook application object created successfully"

                # Get namespace
                $namespace = $outlook.GetNamespace("MAPI")

                # Create new PST file
                Write-ColoredOutput -Message "Adding new PST data store..."
                $namespace.AddStore($FilePath)

                # Find the newly created PST store
                $store = $namespace.Stores | Where-Object { $_.FilePath -eq $FilePath }
                if (-not $store) {
                    throw "Failed to create or access PST file: $FilePath"
                }

                Write-ColoredOutput -Message "PST store created successfully: $($store.DisplayName)"

                # Get root folder
                $rootFolder = $store.GetRootFolder()
                Write-ColoredOutput -Message "Root folder accessed: $($rootFolder.Name)"

                # Ensure default folders exist (create Inbox if it doesn't exist)
                Write-ColoredOutput -Message "Setting up default folder structure..."

                # Check if Inbox exists, create if not
                $inbox = $rootFolder.Folders | Where-Object { $_.Name -eq "Inbox" }
                if (-not $inbox) {
                    Write-ColoredOutput -Message "Creating Inbox folder..."
                    $inbox = $rootFolder.Folders.Add("Inbox")
                }

                # Create custom folders for sample organization
                Write-ColoredOutput -Message "PST folder structure initialized successfully"
                Write-ColoredOutput -Message "Creating custom sample folders..."
                $customFolders = @("Important", "Projects", "Personal", "Archive")
                $allFolders = @{ "Inbox" = $inbox }

                foreach ($customFolderName in $customFolders) {
                    try {
                        # Check if folder already exists
                        $existingCustomFolder = $rootFolder.Folders | Where-Object { $_.Name -eq $customFolderName }

                        if (-not $existingCustomFolder) {
                            Write-ColoredOutput -Message "Creating custom folder: $customFolderName"
                            $newCustomFolder = $rootFolder.Folders.Add($customFolderName)
                            $allFolders[$customFolderName] = $newCustomFolder
                        } else {
                            Write-ColoredOutput -Message "Custom folder already exists: $customFolderName"
                            $allFolders[$customFolderName] = $existingCustomFolder
                        }
                    } catch {
                        Write-ColoredOutput -Message "WARNING: Could not create custom folder '$customFolderName': $($_.Exception.Message)"
                    }
                }

                # Randomly distribute emails across all folders
                $folderNames = $allFolders.Keys
                Write-ColoredOutput -Message "Randomly distributing $EmailCount emails across all folders ($($folderNames -join ', '))..."

                # Create random distribution array
                $folderDistribution = @{}
                foreach ($folderName in $folderNames) {
                    $folderDistribution[$folderName] = 0
                }

                # Randomly assign each email to a folder
                for ($i = 1; $i -le $EmailCount; $i++) {
                    $randomFolder = $folderNames | Get-Random
                    $folderDistribution[$randomFolder]++
                    # Display distribution
                }
                Write-ColoredOutput -Message "Email distribution plan:"
                foreach ($folderName in $folderNames) {
                    $count = $folderDistribution[$folderName]
                    Write-ColoredOutput -Message "  $folderName`: $count emails"
                }

                # Sample data arrays for email generation
                $sampleSubjects = @(
                    "Meeting Follow-up", "Project Update", "Weekly Report", "Action Items",
                    "Budget Review", "Client Feedback", "Team Announcement", "Schedule Change",
                    "Important Notice", "Document Review", "Status Update", "Planning Session",
                    "Quarterly Results", "System Maintenance", "Policy Changes", "Training Session"
                )

                $sampleSenders = @(
                    "john.doe@company.com", "jane.smith@company.com", "mike.johnson@client.com",
                    "sarah.wilson@partner.org", "david.brown@vendor.net", "lisa.davis@consultant.biz",
                    "admin@system.local", "support@vendor.com", "manager@department.org"
                )

                $sampleBodies = @(
                    "Please find the attached document for your review. Let me know if you have any questions or concerns.",
                    "Following up on our meeting yesterday. Here are the action items we discussed.",
                    "This is the weekly status report for the project. All milestones are on track.",
                    "I wanted to update you on the latest developments regarding the budget review.",
                    "The client has provided feedback on the proposal. Please see the details below.",
                    "Please be aware of the schedule changes for next week's meetings.",
                    "The quarterly review is scheduled for next month. Please prepare your reports accordingly.",
                    "System maintenance is planned for this weekend. Please plan accordingly.",
                    "New company policies have been updated. Please review at your earliest convenience."
                )

                $totalSuccessfulEmails = 0

                # Create emails for each folder according to distribution
                foreach ($folderName in $folderNames) {
                    $emailsForThisFolder = $folderDistribution[$folderName]

                    if ($emailsForThisFolder -eq 0) {
                        Write-ColoredOutput -Message "Skipping $folderName (no emails assigned)"
                        continue
                    }

                    Write-ColoredOutput -Message "Creating $emailsForThisFolder emails for $folderName..."
                    $folderSuccessCount = 0
                    $targetFolder = $allFolders[$folderName]

                    for ($j = 1; $j -le $emailsForThisFolder; $j++) {
                        # Get random content
                        try {
                            $randomSubject = $sampleSubjects | Get-Random
                            $randomSender = $sampleSenders | Get-Random
                            $randomBody = $sampleBodies | Get-Random

                            # Create email item using CreateItem method
                            # 0 = olMailItem
                            $mailItem = $outlook.CreateItem(0)

                            # Set email properties with folder context
                            if ($folderName -eq "Inbox") {
                                $mailItem.Subject = "$randomSubject"
                                $mailItem.Body = "$randomBody`n`nGenerated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC')"
                            } else {
                                $mailItem.Subject = "[$folderName] $randomSubject"
                                $mailItem.Body = "$randomBody`n`nFolder Category: $folderName`nGenerated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC')"

                                # Add category to help identify the intended folder
                                try {
                                    $mailItem.Categories = $folderName
                                } catch {
                                    Write-Verbose "Could not set category: $($_.Exception.Message)"
                                }
                            }

                            $mailItem.To = "testuser@example.com"

                            # Set sender information
                            try {
                                $mailItem.SenderName = ($randomSender -split '@')[0] -replace '\.', ' '
                                $mailItem.SenderEmailAddress = $randomSender
                            } catch {
                                Write-Verbose "Could not set sender information: $($_.Exception.Message)"
                            }

                            # Set random past dates (SentOn should be before ReceivedTime)
                            $daysBack = Get-Random -Minimum 1 -Maximum 90
                            $hoursBack = Get-Random -Minimum 1 -Maximum 24
                            $minutesBack = Get-Random -Minimum 1 -Maximum 60
                            $randomSentDate = (Get-Date).AddDays(-$daysBack).AddHours(-$hoursBack).AddMinutes(-$minutesBack)
                            $randomReceivedDate = $randomSentDate.AddMinutes((Get-Random -Minimum 1 -Maximum 30))

                            try {
                                $mailItem.SentOn = $randomSentDate
                                $mailItem.ReceivedTime = $randomReceivedDate
                                $mailItem.CreationTime = $randomReceivedDate
                            } catch {
                                Write-Verbose "Could not set date properties: $($_.Exception.Message)"
                            }

                            # Set read status randomly for more realistic data
                            try {
                                # Randomly mark some as unread
                                $mailItem.UnRead = (Get-Random -Maximum 2) -eq 1
                            } catch {
                                Write-Verbose "Could not set read status: $($_.Exception.Message)"
                                # Set importance randomly for more realistic data
                            }
                            try {
                                # 0=Low, 1=Normal, 2=High
                                $mailItem.Importance = Get-Random -Minimum 0 -Maximum 3
                            } catch {
                                Write-Verbose "Could not set importance: $($_.Exception.Message)"
                            }

                            # Add random attachments based on percentage
                            $currentAttachmentFiles = @()
                            if ($AttachmentPercentage -gt 0) {
                                $shouldAddAttachment = (Get-Random -Minimum 1 -Maximum 101) -le $AttachmentPercentage

                                if ($shouldAddAttachment) {
                                    try {
                                        Write-Verbose "Generating attachment for email $j in $folderName..."
                                        $attachmentInfo = New-RandomAttachment -TempDirectory $AttachmentTempPath

                                        if ($attachmentInfo.Success) {
                                            Write-Verbose "Adding attachment: $($attachmentInfo.FileName) ($($attachmentInfo.SizeMB) MB)"
                                            $attachment = $mailItem.Attachments.Add($attachmentInfo.FilePath)
                                            $attachment.DisplayName = $attachmentInfo.FileName

                                            # Track temp files for cleanup
                                            $currentAttachmentFiles += $attachmentInfo.FilePath
                                            $attachmentTempFiles += $attachmentInfo.FilePath

                                            # Update email body to mention attachment
                                            $attachmentNote = "`n`n--- Attachment ---`nFile: $($attachmentInfo.FileName)`nSize: $($attachmentInfo.SizeMB) MB"
                                            $mailItem.Body += $attachmentNote

                                            Write-Verbose "Successfully added attachment: $($attachmentInfo.FileName)"
                                        } else {
                                            Write-Verbose "Failed to generate attachment for email $j"
                                        }
                                    } catch {
                                        Write-ColoredOutput -Message "  WARNING: Failed to add attachment to email $j in $folderName`: $($_.Exception.Message)"
                                    }
                                }

                            }
                            # Save the email first
                            $mailItem.Save()

                            # Move the email to the target folder
                            $mailItem = $mailItem.Move($targetFolder)

                            $folderSuccessCount++

                            # Clean up
                            $totalSuccessfulEmails++
                            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($mailItem) | Out-Null

                            # Clean up temporary attachment files for this email
                            foreach ($tempFile in $currentAttachmentFiles) {
                                try {
                                    if (Test-Path -Path $tempFile) {
                                        Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
                                        Write-Verbose "Cleaned up temp attachment file: $tempFile"
                                    }
                                } catch {
                                    Write-Verbose "Could not clean up temp file '$tempFile': $($_.Exception.Message)"
                                }
                            }

                            if ($j % 5 -eq 0 -or $j -eq $emailsForThisFolder) {
                                Write-ColoredOutput -Message "  Created email $j of $emailsForThisFolder for $folderName"
                            }
                        } catch {
                            Write-ColoredOutput -Message "  WARNING: Failed to create email $j for $folderName`: $($_.Exception.Message)"
                        }
                    }

                    if ($folderSuccessCount -gt 0) {
                        Write-ColoredOutput -Message "Successfully created $folderSuccessCount emails for $folderName"
                    }
                }            Write-ColoredOutput -Message "PST file creation completed: $FilePath"

                # Summary of what was created
                Write-Output ""
                Write-ColoredOutput -Message "=== Summary ==="
                Write-ColoredOutput -Message "[SUCCESS] PST file created with folder structure"
                Write-ColoredOutput -Message "[SUCCESS] $totalSuccessfulEmails of $EmailCount emails created and saved to PST folders"
                Write-ColoredOutput -Message "[SUCCESS] Emails distributed across PST folders (not in default mailbox)"
                if ($AttachmentPercentage -gt 0) {
                    $expectedAttachmentCount = [Math]::Ceiling($EmailCount * ($AttachmentPercentage / 100))
                    Write-ColoredOutput -Message "[SUCCESS] Random attachments (2-5MB) added to approximately $AttachmentPercentage% of emails"
                }
                Write-Output ""
                Write-ColoredOutput -Message "PST file ready for use: $FilePath"
                Write-ColoredOutput -Message "You can now import this PST file into Outlook or other applications"

                # Clean up any remaining temporary attachment files
                if ($AttachmentPercentage -gt 0) {
                    Write-ColoredOutput -Message "Cleaning up temporary attachment files..."
                    try {
                        if (Test-Path -Path $AttachmentTempPath) {
                            $remainingFiles = Get-ChildItem -Path $AttachmentTempPath -File -ErrorAction SilentlyContinue
                            foreach ($file in $remainingFiles) {
                                try {
                                    Remove-Item -Path $file.FullName -Force -ErrorAction SilentlyContinue
                                    Write-Verbose "Cleaned up remaining temp file: $($file.Name)"
                                } catch {
                                    Write-Verbose "Could not clean up temp file '$($file.Name)': $($_.Exception.Message)"
                                }
                            }

                            # Try to remove the temp directory if it's empty
                            $remainingItems = Get-ChildItem -Path $AttachmentTempPath -ErrorAction SilentlyContinue
                            if (-not $remainingItems) {
                                Remove-Item -Path $AttachmentTempPath -Force -ErrorAction SilentlyContinue
                                Write-ColoredOutput -Message "Temporary attachment directory cleaned up"
                            }
                        }
                    } catch {
                        Write-ColoredOutput -Message "WARNING: Could not fully clean up temporary attachment directory: $($_.Exception.Message)"
                    }
                    # Optional: Remove the PST from Outlook profile (keep file but remove from profile)
                }
                # This prevents the test PST from remaining in the user's Outlook
                Write-ColoredOutput -Message "Removing PST from Outlook profile (file remains on disk)..."
                try {
                    $namespace.RemoveStore($store.GetRootFolder())
                    Write-ColoredOutput -Message "PST removed from Outlook profile successfully"
                } catch {
                    Write-ColoredOutput -Message "WARNING: Could not remove PST from profile: $($_.Exception.Message)"
                    Write-ColoredOutput -Message "You may need to manually remove it from Outlook"
                }
            }
        } else {
            Write-ColoredOutput -Message "WHATIF: Would create PST file at $FilePath with $EmailCount emails randomly distributed across folders"
            if ($AttachmentPercentage -gt 0) {
                $expectedAttachmentCount = [Math]::Ceiling($EmailCount * ($AttachmentPercentage / 100))
                Write-ColoredOutput -Message "WHATIF: Would add random attachments (2-5MB) to approximately $expectedAttachmentCount emails ($AttachmentPercentage%)"
                Write-ColoredOutput -Message "WHATIF: Would use temporary attachment directory: $AttachmentTempPath"
            }
        }
    } catch {
        Write-ColoredOutput -Message "ERROR: $($_.Exception.Message)"
        throw
    } finally {
        # Clean up COM objects - be more careful with cleanup order
        try {
            if ($allFolders) {
                foreach ($folderObj in $allFolders.Values) {
                    try {
                        if ($folderObj -and [System.Runtime.Interopservices.Marshal]::IsComObject($folderObj)) {
                            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($folderObj) | Out-Null
                        }
                    } catch {
                        Write-Verbose "Folder cleanup warning: $($_.Exception.Message)"
                    }
                }
            }
            if ($inbox -and [System.Runtime.Interopservices.Marshal]::IsComObject($inbox)) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($inbox) | Out-Null
            }
            if ($rootFolder -and [System.Runtime.Interopservices.Marshal]::IsComObject($rootFolder)) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($rootFolder) | Out-Null
            }
            if ($store -and [System.Runtime.Interopservices.Marshal]::IsComObject($store)) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($store) | Out-Null
            }
            if ($namespace -and [System.Runtime.Interopservices.Marshal]::IsComObject($namespace)) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
            }
            if ($outlook -and [System.Runtime.Interopservices.Marshal]::IsComObject($outlook)) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
            }
        } catch {
            Write-Verbose "COM cleanup warning: $($_.Exception.Message)"
        } finally {
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
}

# Main execution block
try {
    # Initialize log file
    Write-ColoredOutput -Message "=== New-FakePSTFile.ps1 Execution Started ==="
    Write-ColoredOutput -Message "Script Version: 2.9.4"
    Write-ColoredOutput -Message "Execution Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC')"
    Write-ColoredOutput -Message "System: $systemName"
    Write-ColoredOutput -Message "Log File: $logPath"
    Write-ColoredOutput -Message "Target PST Path: $PSTPath"
    Write-ColoredOutput -Message "Number of Emails: $EmailCount"
    Write-ColoredOutput -Message "Attachment Percentage: $AttachmentPercentage%"
    Write-ColoredOutput -Message "Attachment Temp Path: $AttachmentTempPath"
    Write-ColoredOutput -Message "WhatIf Mode: $($PSBoundParameters.ContainsKey('WhatIf') -and $PSBoundParameters['WhatIf'])"
    Write-ColoredOutput -Message "==========================================="

    # Validate parameters
    if ([string]::IsNullOrWhiteSpace($PSTPath)) {
        throw "PST file path cannot be empty"
    }

    if ($EmailCount -lt 1 -or $EmailCount -gt 1000000) {
        throw "Number of emails must be between 1 and 1000000"
    }

    # Check if PST file already exists
    if (Test-Path -Path $PSTPath) {
        if ($PSCmdlet.ShouldProcess($PSTPath, "Overwrite existing PST file")) {
            Write-ColoredOutput -Message "WARNING: PST file already exists and will be overwritten: $PSTPath"
        } else {
            Write-ColoredOutput -Message "WHATIF: Would overwrite existing PST file: $PSTPath"
        }

    }
    # Create the PST file with data
    $isWhatIfMode = $PSBoundParameters.ContainsKey('WhatIf') -and $PSBoundParameters['WhatIf']
    New-PSTFileWithData -FilePath $PSTPath -EmailCount $EmailCount -IsWhatIfMode $isWhatIfMode -AttachmentPercentage $AttachmentPercentage -AttachmentTempPath $AttachmentTempPath

    $endTime = Get-Date
    $duration = $endTime - $startTime

    Write-ColoredOutput -Message "=== Execution Completed Successfully ==="
    Write-ColoredOutput -Message "Total Execution Time: $($duration.TotalSeconds) seconds"
} catch {
    $endTime = Get-Date
    $duration = $endTime - $startTime

    Write-ColoredOutput -Message "=== Execution Failed ==="
    Write-ColoredOutput -Message "ERROR: $($_.Exception.Message)"
    Write-ColoredOutput -Message "Total Execution Time: $($duration.TotalSeconds) seconds"

    exit 1
} finally {
    Write-ColoredOutput -Message "Log file saved to: $logPath"
}
