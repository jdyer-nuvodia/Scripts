<!--
=============================================================================
File: README.md
Author: jdyer-nuvodia
Last Updated: 2025-07-03 21:55:00 UTC
Updated By: jdyer-nuvodia
Version: 1.4.0
Additional Info: Enhanced script standards to match comprehensive copilot-instructions.md requirements and added development workflow section
=============================================================================
-->

# PowerShell Scripts Collection

This repository contains a collection of PowerShell scripts for various system administration tasks, cloud management, and maintenance operations.

## Directory Structure

- **CloudServices/** - Scripts for cloud service management
  - **365/** - Microsoft 365 management scripts
    - Add-UserListTo365Group
    - Delete-Mailboxes
    - Diagnose-MailboxFolderAssistant
    - Get-AllMailboxforwardingRules
    - Get-CalendarPermissions
    - Get-FullMailboxAttributes
    - Get-Mailboxes
    - Get-MailboxFolderList
    - Grant-CalendarPermissions
    - Grant-RMToMailboxEditCalendarPermissions
    - List-PNPFolderItemCount
    - Remediate-365Account
    - Remove-AllMailboxPermissions
  - **Azure/** - Azure resource management
    - Deploy-ActionGroup
    - Test-VMAutoShutdown

- **Development/** - Development-related utilities
  - **CodeQuality/** - Code analysis and improvement tools
    - Invoke-PowerShellCodeCleanup
  - **FileConversion/** - File format conversion utilities
    - Convert-MarkdownToText
  - **Git/** - Git repository management tools
    - Generate-GitHistory
  - **TestDataGeneration/** - Test data creation utilities
    - New-FakePSTFile

- **Network/** - Network administration and diagnostics
  - **Test-AdvancedNetworkConnectivity/** - Advanced network connectivity testing
  - **Test-NetworkConnectivity/** - Basic network connectivity testing

- **PersonalUtilities/** - Personal workflow optimization scripts
  - **Delete-OldScreenshots/** - Screenshot cleanup utility

- **Security/** - Security and compliance tools
  - **Auditing/** - Security audit scripts (placeholder for future expansion)
  - **Compliance/**
    - Get-SecurityConfigurationStatus
  - **Permissions/**
    - **NTFS/** - NTFS permission management
      - Get-NTFSFolderPermissions
      - Get-NTFSPermissionsForUser

- **Software/** - Software deployment and management
  - **Installation/**
    - Get-WizTreePortable
  - **Management**
    - Get-InstalledSoftware
  - **Reinstallation/** - Software reinstallation tools
  - **Removal/**
    - Remove-AdobeAcrobatReader

- **SystemManagement/** - System administration and maintenance
  - **FileSystem/**
    - **Cleanup/** - File system cleanup utilities
      - Clear-SystemStorage
      - Delete-AllFilesInDirectory
      - Delete-OldFiles
      - Delete-OldUserFiles
      - Delete-UserFiles
      - Find-LargestFolders
      - Get-DriveInfo
    - **Naming/** - File and folder naming utilities
      - Rename-FolderCase
      - Rename-ScriptToFolderName
    - **Search/** - File system search tools
      - Search-ContentRecursively
  - **Maintenance/**
    - Analyze-WindowsLogs
    - Get-EventLogs
    - Get-SystemHealthReport
    - **Printers/** - Printer management utilities
      - Clear-PrintQueue
    - Repair-WindowsOS
  - **PATHManagement/** - System PATH environment variable management
    - Add-FoldersToPath
    - Reset-PATH
  - **Performance/** - System performance monitoring and optimization
    - Get-SetInactivityTimers
    - Monitor-SystemResources
  - **Services/** - Windows service management
    - Monitor-CriticalServices
  - **Sessions/** - User session management
    - Manage-RDPSessions

- **UserManagement/** - User administration tools
  - **Accounts/**
    - **Domain/** - Active Directory user account management
      - Change-ADUserPassword
      - Copy-ADUser
      - Create-ADUser
      - Get-RecentAccountLockouts
      - Remove-GroupsFromDisabledUsers
    - **Local/** - Local user account management
      - Change-LocalUserPassword
      - Create-LocalUserAccount
  - **Groups/** - Group management utilities (placeholder for future expansion)
  - **Permissions/** - User permissions management (placeholder for future expansion)

## Getting Started

All scripts include detailed help information accessible via:

```powershell
Get-Help .\ScriptName.ps1 -Full
```

Most scripts include -WhatIf as a parameter to test changes before running in production.

## Development Workflow

This repository includes standardized GitHub Copilot prompts for consistent development:

- **`/create-powershell-script`** - Generate new scripts with complete organizational standards compliance
- **`/review-psscriptanalyzer-compliance`** - Comprehensive code review and compliance checking
- **`/run-code-quality-scripts`** - Execute all code quality scripts including alignment analysis, Write-Host conversion, cleanup, and PSScriptAnalyzer checks
- **`/run-powershell-code-cleanup`** - Safe whitespace and formatting fixes

These prompts ensure adherence to all organizational standards and automate the development workflow.

## Repository Infrastructure

- **GitHub Actions**: Automated PowerShell validation and dependency scanning
- **VS Code Configuration**: Complete workspace setup with PSScriptAnalyzer settings
- **Issue Templates**: Standardized templates for bug reports, feature requests, and security issues
- **Pull Request Templates**: Consistent formatting for code contributions

## Script Standards

All scripts in this repository follow these comprehensive standards:

### Critical Security Requirements

- **NEVER** use `-ExecutionPolicy Bypass` - this triggers security alerts and violates enterprise policies
- Use secure alternatives: `RemoteSigned` or `AllSigned` execution policies
- Digitally sign all scripts with valid certificates
- Never store or hardcode credentials in scripts
- Use secure credential mechanisms: `Get-Credential`, credential vaults, or encrypted files

### Documentation Standards

- Complete header documentation following standardized format
- Parameter descriptions and usage examples
- Version control information with UTC timestamps
- Clear synopsis and detailed descriptions
- No contractions in comments or documentation
- Use Verb-Noun format for functions and scripts

### Coding Standards

- **Error handling**: Use `try`, `catch`, and `finally` blocks with proper exit codes
- **Parameter validation**: Accept input via named parameters with sensible defaults
- **Idempotency**: Scripts must be safely executable multiple times
- **WhatIf support**: Implement `-WhatIf` for operations that modify system state
- **PSScriptAnalyzer compliance**: All scripts must pass without warnings or errors
- **No Write-Host**: Use `Write-ColorOutput` function for colored output instead
- **Named parameters**: Always use named parameters, never positional parameters
- **Variable scoping**: Use explicit scoping (`$script:ParameterName`) for script parameters within functions
- **Modularity**: Break logic into small, reusable functions or modules
- **Automation compatibility**: Compatible with Windows Task Scheduler, no interactive prompts

### Formatting and Style

- **Indentation**: Use 4 spaces per level (no tabs)
- **Operator spacing**: Exactly one space before and after all operators
- **Consistent formatting**: Properly aligned braces and code blocks
- **ASCII only**: No non-ASCII characters in script names or content
- **Full cmdlet names**: Avoid aliases, use complete PowerShell cmdlet names

### Logging and Auditing

- Implement comprehensive logging using `Write-Output` or `Write-Verbose`
- Save logs to script directory with system name and UTC timestamp
- Use `.log` extension for all log files
- Log all colored output to transcript files without color codes

### Color Scheme Standards

- **White**: Standard information
- **Cyan**: Process updates and background operations
- **Green**: Success messages and completed operations
- **Yellow**: Warnings and issues requiring attention
- **Red**: Errors and critical issues
- **Magenta**: Debug information and detailed troubleshooting
- **DarkGray**: Less important details and secondary information

### Version Control

- **Semantic versioning**: MAJOR.MINOR.PATCH format
- **UTC timestamps**: YYYY-MM-DD HH:MM:SS UTC format only
- **Commit messages**: `<type>(<scope>): <description>` format
- **Change documentation**: Update headers for every modification
- **Version increments**: Required for all changes with change summaries

## Contributors

- jdyer-nuvodia

## License

This repository is for internal use. All rights reserved.
