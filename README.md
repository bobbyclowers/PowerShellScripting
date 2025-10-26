<!-- omit from toc -->
# PowerShellScripting

A collection of PowerShell scripts for enterprise IT administration, covering Active Directory, Microsoft 365, Exchange Online, Entra ID, and Intune management tasks that i've created over the years.

[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)
[![PowerShell Gallery](https://img.shields.io/badge/PowerShell-7.0+-blue.svg)](https://github.com/PowerShell/PowerShell)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)

## Table of Contents

- [Table of Contents](#table-of-contents)
- [Features](#features)
- [Getting Started](#getting-started)
  - [Prerequisites](#prerequisites)
  - [Installation](#installation)
- [Usage](#usage)
  - [Basic Usage](#basic-usage)
  - [Script Categories](#script-categories)
  - [Interactive Scripts](#interactive-scripts)
- [Configuration](#configuration)
  - [Environment Variables](#environment-variables)
  - [Authentication](#authentication)
  - [Customisation](#customisation)
- [Folder Structure](#folder-structure)
- [Modules and Functions](#modules-and-functions)
  - [Core Functionality](#core-functionality)
  - [Key Scripts](#key-scripts)
- [Testing](#testing)
  - [Development Environment](#development-environment)
  - [Validation](#validation)
- [Logging and Troubleshooting](#logging-and-troubleshooting)
  - [Logging Standards](#logging-standards)
  - [Common Issues](#common-issues)
  - [Support Resources](#support-resources)
- [Accessibility](#accessibility)
- [Contributing](#contributing)
  - [Development Guidelines](#development-guidelines)
- [Changelog](#changelog)
  - [Recent Updates](#recent-updates)
  - [Version History](#version-history)
- [License](#license)
- [Like to say thank you?](#like-to-say-thank-you)
- [Contact and Support](#contact-and-support)
  - [Project Maintainer](#project-maintainer)
  - [Getting Help](#getting-help)
  - [Support Guidelines](#support-guidelines)

## Features

- **Active Directory Management**: User creation, group management, computer organisation, and bulk operations
- **Microsoft 365 Administration**: Exchange Online mailbox management, quarantine handling, and transport rules
- **Entra ID Integration**: External user management, compromised account remediation, and identity operations
- **Intune Device Management**: Bulk device synchronisation, remediation scripts, and compliance monitoring
- **General Utilities**: Password generation, module management, and script selection tools
- **OneDrive Administration**: User content download and management capabilities
- **Comprehensive Logging**: Standardised logging across all scripts with detailed audit trails
- **Error Handling**: Robust error handling and retry logic for enterprise environments
- **GUI Interfaces**: User-friendly forms for complex administrative tasks

## Getting Started

### Prerequisites

- PowerShell 7.0 or later
- Windows operating system
- Appropriate administrative permissions for target systems
- Required PowerShell modules (see individual scripts for specific requirements):
  - Active Directory Module
  - Exchange Online Management
  - Microsoft Graph PowerShell SDK
  - Microsoft.Graph.Intune
  - MSOnline (where applicable)

### Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/twcau/PowerShellScripting.git
   ```

2. Navigate to the project directory:

   ```bash
   cd PowerShellScripting
   ```

3. Review the script you want to use and install any required modules:

   ```powershell
   # Example: Install Exchange Online Management module
   Install-Module -Name ExchangeOnlineManagement -Force
   ```

4. Configure the scripts according to your environment (see Configuration section)

## Usage

### Basic Usage

Each script is designed to be run independently. Navigate to the appropriate folder and execute the script:

```powershell
# Example: Run user creation script
.\ad\user\creation\User-Creation.ps1

# Example: Run Intune bulk sync
.\intune\devices\Intune-BulkSync.ps1
```

### Script Categories

- **Active Directory**: Scripts for user and computer management in on-premises AD environments
- **Exchange 365**: Email and mailbox management for cloud and hybrid environments
- **Entra ID**: Identity and access management for Azure AD/Entra ID
- **Intune**: Mobile device management and compliance scripts
- **General**: Utility scripts for common administrative tasks

### Interactive Scripts

Many scripts include GUI interfaces for ease of use:

- User creation wizards with form-based input
- Device selection interfaces
- Progress indicators for long-running operations

## Configuration

### Environment Variables

Some scripts may require environment-specific configuration. Review each script's header for specific requirements.

### Authentication

- Ensure you have appropriate administrative credentials
- Some scripts require multi-factor authentication (MFA)
- Consider using application passwords where applicable

### Customisation

Scripts include configurable parameters at the top of each file. Common customisations include:

- Domain names and organisational units
- Email domains and Exchange settings
- Logging paths and retention policies
- Timeout values and retry attempts

## Folder Structure

```plaintext
PowerShellScripting/
├── ad/                                    # Active Directory scripts
│   ├── computer/
│   │   └── FindMachineOU.ps1             # Locate computer objects in AD
│   └── user/
│       ├── creation/                      # User account creation scripts
│       │   ├── AD-CopyGroups.ps1         # Copy group memberships
│       │   ├── User-Creation-Bulk.ps1    # Bulk user creation
│       │   ├── User-Creation.ps1         # Individual user creation with GUI
│       │   └── User-Departure.ps1        # User departure processing
│       └── reconcillation/               # User account reconciliation
│           ├── AD-Bulk-DepartedEmployeeReconcillation.ps1
│           ├── Employee-Departure-Reconciliation.ps1
│           └── Employee-Listing.ps1
├── e365/                                  # Exchange 365 scripts
│   ├── E365-Mailbox-ConvertToShared.ps1 # Convert mailboxes to shared
│   ├── E365-Quarantine-ExportRecord.ps1 # Export quarantine records
│   ├── Exchange-QuarantineTABL-DataDownload.ps1
│   └── NewTransportRuleExecName.ps1     # Transport rule management
├── entra/                                 # Entra ID (Azure AD) scripts
│   ├── AutomateCompromisedAccountRemediation.ps1
│   ├── Entra-UserExternal-Create.ps1    # External user creation
│   └── User-Management-External.ps1     # External user management
├── general/                               # General utility scripts
│   ├── ScriptSelector.ps1                # Interactive script launcher
│   ├── module-management/                # PowerShell module utilities
│   │   ├── Module-PowerShell7-Require.ps1
│   │   └── Update-Module.ps1
│   └── password-generation/              # Password generation tools
│       ├── Password-Generator-Silent.ps1
│       └── Password-Generator.ps1
├── intune/                                # Microsoft Intune scripts
│   ├── devices/
│   │   └── Intune-BulkSync.ps1          # Bulk device synchronisation
│   └── remediation/                      # Intune remediation scripts
│       ├── M365-VersionDetect.ps1       # M365 Apps version detection
│       ├── M365-VersionRemediate.ps1    # M365 Apps version remediation
│       ├── TeamsOld-Detect.ps1          # Legacy Teams detection
│       ├── TeamsOld-Remediate.ps1       # Legacy Teams remediation
│       ├── WinUpdate-23H2to24H2Force-Detect.ps1
│       ├── WinUpdate-23H2to24H2Force-Remediate.ps1
│       ├── WinUpdate-Detect.ps1         # Windows Update detection
│       ├── WinUpdate-Pause-Detect.ps1   # Windows Update pause detection
│       ├── WinUpdate-Pause-Remediate.ps1
│       └── WinUpdate-Remediate.ps1      # Windows Update remediation
├── m365/                                  # Microsoft 365 scripts
├── onedrive/                              # OneDrive management scripts
│   └── M365-OneDrive-DownloadUserContents.ps1
└── testing/                               # Development and testing scripts
```

## Modules and Functions

### Core Functionality

The scripts in this collection provide:

- **User Management**: Creation, modification, and departure processing
- **Group Management**: Membership copying and bulk operations
- **Device Management**: Synchronisation, detection, and remediation
- **Security Operations**: Compromised account handling and compliance monitoring
- **Utility Functions**: Password generation, module management, and system utilities

### Key Scripts

- **User-Creation.ps1**: Comprehensive user creation with GUI interface
- **Intune-BulkSync.ps1**: Mass device synchronisation for Intune environments
- **AutomateCompromisedAccountRemediation.ps1**: Automated security response
- **ScriptSelector.ps1**: Interactive menu system for script selection

## Testing

### Development Environment

Testing scripts are located in the `testing/` folder and include:

- Proof-of-concept implementations
- Version comparisons
- Experimental features

### Validation

Before using scripts in production:

1. Review the script header for version information and changelog
2. Test in a non-production environment
3. Verify all required modules are installed
4. Check logging output for any warnings or errors

## Logging and Troubleshooting

### Logging Standards

All scripts follow consistent logging practices:

- Log files stored in `$env:TEMP` with timestamps
- Comprehensive error logging with context
- Success and failure reporting
- Progress indicators for long-running operations

### Common Issues

- **Module Import Errors**: Ensure required PowerShell modules are installed
- **Authentication Failures**: Verify credentials and MFA settings
- **Permission Errors**: Check administrative rights for target systems
- **Network Connectivity**: Ensure access to required cloud services

### Support Resources

- Check script headers for specific documentation links
- Review Microsoft documentation for API changes
- Consult PowerShell Gallery for module updates

## Accessibility

This project is committed to accessibility and inclusive design:

- Scripts include progress indicators and clear status messages
- Documentation uses descriptive text for all functionality
- Error messages provide actionable guidance
- GUI interfaces follow accessibility best practices
- All documentation supports screen readers

## Contributing

Contributions to improve and expand this script collection are welcome. Please read the contribution guidelines:

1. **Code Standards**: Follow PowerShell best practices and existing code style
2. **Documentation**: Include comprehensive headers and inline comments
3. **Testing**: Validate scripts in appropriate test environments
4. **Security**: Ensure no hardcoded credentials or sensitive information

### Development Guidelines

- Use Australian English (EN-AU) for documentation and comments
- Include proper error handling and logging
- Follow the established folder structure
- Update this README when adding new functionality

## Changelog

### Recent Updates

- **6/06/2025**: Enhanced user creation script with group copying improvements
- **27/03/2025**: Added Clear Base User and Clear All User functionality
- **21/05/2025**: Implemented base group validation and management
- **4/03/2025**: Updated department listings for dynamic group memberships

### Version History

See individual script headers for detailed version history and changelog information.

## License

This project is licensed under the GNU General Public License v3.0 - see the [LICENSE](LICENSE) file for details.

Copyright (c) 2025, Michael Harris, All rights reserved.

## Like to say thank you?

If these scripts have helped you in your IT administration tasks, consider:

- ⭐ Starring this repository
- 🐛 Reporting issues or suggesting improvements
- 📖 Contributing to the documentation
- ☕ [Buy me a coffee](https://ko-fi.com/twcau) to support continued development

## Contact and Support

### Project Maintainer

- **Michael Harris** - [@twcau](https://github.com/twcau)

### Getting Help

- **Issues**: Report bugs or request features via [GitHub Issues](https://github.com/twcau/PowerShellScripting/issues)
- **Discussions**: Join the conversation in [GitHub Discussions](https://github.com/twcau/PowerShellScripting/discussions)
- **Documentation**: Review script headers and Microsoft documentation links

### Support Guidelines

- Provide clear descriptions of issues with relevant log files
- Include PowerShell version and module information
- Specify the target environment (on-premises, cloud, hybrid)
- Follow the issue templates when reporting problems

---

*This project follows Microsoft PowerShell best practices and maintains compatibility with enterprise IT environments.*
