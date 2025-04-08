<!--
=============================================================================
File: README.md
Created: 2025-02-07 21:21:53 UTC
Author: jdyer-nuvodia
Last Updated: 2025-04-08 21:49:00 UTC
Updated By: jdyer-nuvodia
Version: 1.1.1
Additional Info: Added previously missing scripts to directory structure
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
    - Remediate-365Account
    - Remove-AllMailboxPermissions
  - **Azure/** - Azure resource management
    - Deploy-ActionGroup
    - Test-VMAutoShutdown

- **Development/** - Development-related utilities
  - **Git/** - Git repository management tools
    - Generate-GitHistory

- **Network/** - Network administration and diagnostics
  - **Test-NetworkConnectivity/** - Network connectivity testing

- **PersonalUtilities/** - Personal workflow optimization scripts
  - **Delete-OldScreenshots/** - Screenshot cleanup utility

- **Security/** - Security and compliance tools
  - **Auditing/** - Security audit scripts
  - **Compliance/** 
    - Get-SecurityConfigurationStatus
  - **Permissions/**
    - NTFS permission management

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
    - **Cleanup/**
      - Get-FolderSizes
    - **Search/**
      - Search-ContentRecursively
    - **Naming**
      - Rename-ScriptToFolderName
      - Rename-FolderCase
  - **Maintenance**
    - Analyze-WindowsLogs
    - Get-EventLogs
    - Get-SystemHealthReport
    - Repair-WindowsOS
  - **PATHManagement**
    - Add-FoldersToPath
    - Reset-MachinePATH
  - **Performance**
    - Get-SetInactivityTimers
    - Monitor-SystemResources
  - **Services**
    - Monitor-CriticalServices

- **UserManagement/** - User administration tools
  - **Accounts/**
    - **Domain/**
      - Change-ADUserPassword
      - Copy-ADUser
      - Create-ADUser
      - Remove-GroupsFromDisabledUsers
    - **Local/**
      - Change-LocalUserPassword
      - Create-LocalUserAccount
  - **Groups/** - Group management utilities
  - **Permissions/** - User permissions management

## Getting Started

All scripts include detailed help information accessible via:

```powershell
Get-Help .\ScriptName.ps1 -Full
```

Most scripts include -WhatIf as a parameter to test changes before running in production.

## Script Standards

All scripts in this repository follow these standards:

### Documentation

- Complete header documentation
- Parameter descriptions and examples
- Version control information
- UTC timestamps

### Coding Standards

- Error handling and logging
- Parameter validation
- PowerShell best practices
- Consistent color scheme for output:
  - White: Standard information
  - Cyan: Process updates
  - Green: Success messages
  - Yellow: Warnings
  - Red: Errors
  - Magenta: Debug information
  - DarkGray: Less important details

### Version Control

- Semantic versioning (MAJOR.MINOR.PATCH)
- Commit message format: [type](scope): description
- UTC timestamp format: YYYY-MM-DD HH:MM:SS UTC
- Change documentation in file headers

## Contributors

- jdyer-nuvodia

## License

This repository is for internal use. All rights reserved.
