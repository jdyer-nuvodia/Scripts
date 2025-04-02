# PowerShell Scripts Collection

This repository contains a collection of PowerShell scripts for various system administration tasks, cloud management, and maintenance operations.

## Directory Structure

- **CloudServices/** - Scripts for cloud service management
  - **365/** - Microsoft 365 management scripts (mailbox, calendar, and user management)
  - **Azure/** - Azure resource management and automation scripts

- **Development/** - Development-related utilities
  - **Git/** - Git repository management tools

- **FileSystem-Management/** - Scripts for file system operations, permissions, and path management
  - NTFS permissions management
  - File system cleanup and organization

- **Network/** - Network administration and diagnostics
  - Network connectivity testing
  - Network troubleshooting utilities

- **PersonalUtilities/** - Personal workflow optimization scripts
  - File management utilities
  - Workspace organization tools

- **Security/** - Security and compliance tools
  - **Auditing/** - Security audit scripts
  - **Compliance/** - Compliance checking and reporting
  - **Permissions/** - Permission management tools

- **Software/** - Software deployment and management
  - **Installation/** - Software installation scripts
  - **Management/** - Software management utilities
  - **Reinstallation/** - Software reinstallation tools
  - **Removal/** - Software removal scripts

- **SystemManagement/** - System administration and maintenance
  - **FileSystem/** - File system management tools
  - **Maintenance/** - System maintenance scripts
  - **PATHManagement/** - System PATH variable management
  - **Performance/** - Performance monitoring and optimization
  - **Services/** - Windows services management

- **UserManagement/** - User administration tools
  - **Accounts/** - Account management scripts
  - **Groups/** - Group management utilities
  - **Permissions/** - User permissions management

## Getting Started

Most scripts include detailed help information accessible via:

```powershell
Get-Help .\ScriptName.ps1 -Full
```

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
- Commit message format: <type>(<scope>): <description>
- UTC timestamp format: YYYY-MM-DD HH:MM:SS UTC
- Change documentation in file headers

## Contributors

- jdyer-nuvodia

## License

This repository is for internal use. All rights reserved.
