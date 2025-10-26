<#
.SYNOPSIS
    Export Microsoft 365 Exchange quarantine data and Tenant Allow/Block List to CSV files for analysis.

.DESCRIPTION
    This script connects to Exchange Online to download quarantine message records and Tenant Allow/Block List data.
    It implements cross-platform compatibility for Windows, macOS, and Linux environments with appropriate file path handling.
    
    Key features:
    - Downloads up to 10,000 quarantine message records in paginated format
    - Exports Tenant Allow/Block List (Sender type)
    - Archives existing CSV files with date stamps before creating new ones
    - Provides cross-platform file path support
    - Requires PowerShell 7 for modern authentication support including FIDO2/WebAuthn
    - Implements robust error handling and progress indication

.PARAMETER None
    This script does not accept parameters.

.INPUTS
    None. This script does not accept pipeline input.

.OUTPUTS
    CSV files:
    - quarantine.csv: Quarantine message records
    - tabl.csv: Tenant Allow/Block List data
    
    Files are saved to platform-specific locations:
    - Windows: V:\Exchange-Blocking\
    - macOS: ~/Documents/Exchange-Blocking/
    - Linux: ~/Exchange-Blocking/

.EXAMPLE
    .\E365-Quarantine-ExportRecord.ps1
    
    Runs the script to export quarantine data and TABL to CSV files.

.NOTES
    Author: Michael HARRIS
    Created: 10/03/2025
    Updated: 28/08/2025
    Version: 3.0
    
    Requirements:
    - PowerShell 7.0 or later
    - ExchangeOnlineManagement module
    - Appropriate Exchange Online permissions
    - Modern authentication capability (FIDO2/WebAuthn support)

.DOCUMENTATION
    Exchange Online PowerShell: https://docs.microsoft.com/en-us/powershell/exchange/
    Get-QuarantineMessage: https://docs.microsoft.com/en-us/powershell/module/exchange/get-quarantinemessage
    Get-TenantAllowBlockListItems: https://docs.microsoft.com/en-us/powershell/module/exchange/get-tenantallowblocklistitems
    Connect-ExchangeOnline: https://docs.microsoft.com/en-us/powershell/module/exchange/connect-exchangeonline

.FILECREATED
    10/03/2025

.FILELASTUPDATED
    28/08/2025

TODO:
- Update header formatting to following method, to prevent formatting leakage onto next row:
  $($PSStyle.Foreground.White)$($PSStyle.Background.Blue)Headertext($PSStyle.Reset)
- See if PowerShell 7 Test Code, and PowerShell 7 module loading can be simplified into a single script

Update history:

28/08/2025
- Add cross-platform support for running the code
- Implement robust directory creation and error handling
- Update documentation to meet coding standards
9/06/2025
- Insert TODO list
4/06/2025
- Converted to PowerShell 7 for FIDO2 WebAuthn redirection support
- Adjusted end script method
14/04/2024
- If existing CSV file(s) in folder, first update file name(s) with creation date as suffix, and move them to archives folder
- Cease adding creation date to file names created in script as a result
- Designed to make running of Excel book easier, and not need to update the file name(s) in the Power Query each time
10/03/2025
- Initial release
#>

<# Core and initial variables #>

<#
Cross-Platform Path Configuration
---------------------------------
This section implements cross-platform compatibility by detecting the operating system
and setting appropriate file paths for Windows, macOS, and Linux environments.

Purpose:
- Ensures the script works consistently across different operating systems
- Maintains existing Windows V: drive functionality for enterprise environments
- Provides sensible defaults for personal/home use on macOS and Linux
- Includes robust fallback mechanism for unknown environments

Implementation Details:
- Uses PowerShell's built-in OS detection variables ($IsWindows, $IsMacOS, $IsLinux)
- Includes backwards compatibility check for PowerShell versions < 6 (assumes Windows)
- Uses [Environment]::GetFolderPath("UserProfile") for reliable home directory detection
- Employs Join-Path for proper cross-platform path construction
- Provides fallback to current working directory if OS detection fails

Path Strategy:
- Windows: Maintains original V:\Exchange-Blocking\ path for enterprise compatibility
- macOS: Uses ~/Documents/Exchange-Blocking/ (standard Documents folder location)
- Linux: Uses ~/Exchange-Blocking/ (directly in user home directory)
- Fallback: Uses current directory + Exchange-Blocking/ subfolder
#>

# Detect operating system and set appropriate paths based on platform
if ($IsWindows -or ($PSVersionTable.PSVersion.Major -lt 6)) {
    # Windows environment - use original enterprise V: drive path
    # Note: PowerShell versions < 6 are Windows PowerShell, so assume Windows
    $directoryPath = "V:\Exchange-Blocking\"
    $targetFolder = "V:\Exchange-Blocking\Archived"
}
elseif ($IsMacOS) {
    # macOS environment - use Documents folder for better user experience
    # Documents folder is the standard location for user-generated files on macOS
    $homeDirectory = [Environment]::GetFolderPath("UserProfile")
    $directoryPath = Join-Path $homeDirectory "Documents/Exchange-Blocking/"
    $targetFolder = Join-Path $homeDirectory "Documents/Exchange-Blocking/Archived"
}
elseif ($IsLinux) {
    # Linux environment - use home directory root for simplicity
    # Linux users typically organise files directly under home directory
    $homeDirectory = [Environment]::GetFolderPath("UserProfile")
    $directoryPath = Join-Path $homeDirectory "Exchange-Blocking/"
    $targetFolder = Join-Path $homeDirectory "Exchange-Blocking/Archived"
}
else {
    # Fallback scenario - if OS detection fails for any reason
    # Use current working directory as a safe default
    Write-Warning "Unable to detect operating system. Using current directory as fallback."
    $directoryPath = Join-Path (Get-Location) "Exchange-Blocking/"
    $targetFolder = Join-Path (Get-Location) "Exchange-Blocking/Archived"
}

# Ensure directories exist
# Create both main directory and archived subdirectory if they don't exist
# This prevents errors when the script tries to write files or move archived files
if (!(Test-Path $directoryPath)) {
    New-Item -ItemType Directory -Path $directoryPath -Force | Out-Null
}
if (!(Test-Path $targetFolder)) {
    New-Item -ItemType Directory -Path $targetFolder -Force | Out-Null
}

# Define the resulting files for the current day's data
# Using Join-Path ensures proper path construction regardless of operating system
$filePathQuarantine = Join-Path $directoryPath "quarantine.csv"
$filePathTABL = Join-Path $directoryPath "tabl.csv"

# Specify required modules for this script, to be tested for by the appropriate for each statement

$moduleNames = @(
    "ExchangeOnlineManagement"
)

#-------------------------------------------------------------------------------------------
# STEP 0: Start script
#-------------------------------------------------------------------------------------------

Write-Host ""
Write-Host "E365 - Export quarantine data for analysis" -Foregroundcolor White -BackgroundColor Blue

#-------------------------------------------------------------------------------------------
# STEP 1: Test if running in PowerShell 7
#-------------------------------------------------------------------------------------------

Write-Host ""
Write-Host "Step 1: Test if running in PowerShell 7" -Foregroundcolor White -BackgroundColor Blue

Write-Host ""

<#
PowerShell Version Validation
----------------------------
This section ensures the script runs in PowerShell 7 (Core) for modern authentication support.

Requirements:
- PowerShell 7.0 or later required for FIDO2/WebAuthn authentication
- Cross-platform compatibility requires PowerShell Core
- Enterprise security features require modern PowerShell capabilities

Documentation:
- PowerShell Installation Guide: https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell
- $PSVersionTable Variable: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_automatic_variables#psversiontable
- Start-Process Cmdlet: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/start-process
- Get-Command Cmdlet: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/get-command
#>

# Check if running in PowerShell 7 (Core)
if ($PSVersionTable.PSEdition -ne 'Core' -or $PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7. Relaunching in PowerShell 7..." -ForegroundColor Yellow

    # Construct the full path to this script
    $scriptPath = $MyInvocation.MyCommand.Path

    # Check if 'pwsh' is available (PowerShell 7)
    if (Get-Command pwsh -ErrorAction SilentlyContinue) {
        # Start PowerShell 7 and pass this script
        Start-Process -FilePath "pwsh" -ArgumentList "-NoExit", "-File", "`"$scriptPath`""
        exit
    }
    else {
        Write-Host "❌ Error - PowerShell 7 (pwsh) is not installed or not found in PATH." -Foregroundcolor White -BackgroundColor Red
        Write-Host ""
        Write-Host "Check that PowerShell 7 is properly installed and configured before running this script again."
        Write-Host "If you need to install PowerShell 7, go to the following URL:"
        Write-Host "Windows devices - https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.5"
        Write-Host ""
        Write-Host "Press the Enter key to end the script..."
        Read-Host
        exit 1
    }
}

# Safe to run PowerShell 7 code below this point
Write-Host "Running in PowerShell 7, continuing script..." -ForegroundColor Green

#-------------------------------------------------------------------------------------------
# STEP 2: Test for required modules
#-------------------------------------------------------------------------------------------

Write-Host ""
Write-Host "Step 2: Testing for required modules" -Foregroundcolor White -BackgroundColor Blue

Write-Host ""

<#
Module Dependency Management
---------------------------
This section ensures all required PowerShell modules are installed and imported.

Required Modules:
- ExchangeOnlineManagement: Provides cmdlets for connecting to and managing Exchange Online

Module Management Process:
1. Check if module is available in the system
2. Install module if not found (CurrentUser scope for security)
3. Import module with Force parameter to ensure latest version

Documentation:
- Get-Module Cmdlet: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/get-module
- Install-Module Cmdlet: https://docs.microsoft.com/en-us/powershell/module/powershellget/install-module
- Import-Module Cmdlet: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/import-module
- ExchangeOnlineManagement Module: https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
#>

foreach ($moduleName in $moduleNames) {
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        Write-Host "Module '{0}' not found. Installing..." -f $moduleName -ForegroundColor Yellow
        Install-Module -Name $moduleName -Scope CurrentUser -Force
    }
    else {
        Write-Host "Module '{0}' is already installed." -f $moduleName -ForegroundColor Green
    }

    # Import the module with Force to ensure latest version is loaded
    Write-Host "Importing Module '{0}'." -f $moduleName -ForegroundColor Green
    Import-Module -Name $moduleName -Force
}

#-------------------------------------------------------------------------------------------
# STEP 3: Archive existing CSV files
#-------------------------------------------------------------------------------------------

Write-Host ""
Write-Host "Step 3: Archive old .CSV files" -Foregroundcolor White -BackgroundColor Blue

Write-Host "Rename previous CSV files in folder, and move to \Archived"
Write-Host ""

<#
File Archiving Process
---------------------
This section archives existing CSV files to prevent data loss and maintain clean working directory.

Process:
1. Ensure archive directory exists (redundant safety check)
2. Locate all CSV files in the working directory
3. Rename files with creation date suffix (YYYYMMDD format)
4. Move renamed files to archive directory

Benefits:
- Prevents overwriting previous data exports
- Maintains consistent file names for Excel Power Query integration
- Provides historical data retention
- Enables easy data comparison across time periods

Documentation:
- Test-Path Cmdlet: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/test-path
- New-Item Cmdlet: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/new-item
- Get-ChildItem Cmdlet: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/get-childitem
- Rename-Item Cmdlet: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/rename-item
- Move-Item Cmdlet: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/move-item
- Join-Path Cmdlet: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/join-path
#>

# Ensure target directory exists before attempting to move files
# This provides additional safety in case directories were deleted after initial creation
if (!(Test-Path $targetFolder)) {
    Write-Host "Creating archive directory: {0}" -f $targetFolder -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $targetFolder -Force | Out-Null
}

# Get all CSV files in the working directory
$csvFiles = Get-ChildItem -Path $directoryPath -Filter *.csv

if ($csvFiles.Count -gt 0) {
    foreach ($file in $csvFiles) {
        # Get the file's creation date for archive naming
        $creationDate = $file.CreationTime

        # Format the date as YYYYMMDD for consistent file naming
        $dateSuffix = $creationDate.ToString("yyyyMMdd")

        # Extract file name components for reconstruction
        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($file)
        $fileExtension = $file.Extension

        # Create the new file name with date suffix
        $newFileName = "{0}-{1}{2}" -f $fileName, $dateSuffix, $fileExtension
        # Use Join-Path for cross-platform compatibility instead of [System.IO.Path]::Combine()
        $newFilePath = Join-Path $directoryPath $newFileName

        # Rename the file with date suffix
        Rename-Item -Path $file.FullName -NewName $newFileName

        # Move the renamed file to the archive directory
        Move-Item -Path $newFilePath -Destination $targetFolder

        Write-Output "File renamed to: {0} and moved to {1}" -f $newFileName, $targetFolder
    }
    Write-Output "All CSV files have been renamed and moved to archive."
}
else {
    Write-Output "No CSV files found in the directory."
}


#-------------------------------------------------------------------------------------------
# STEP 4: Exchange Online Authentication
#-------------------------------------------------------------------------------------------

Write-Host ""
Write-Host "Step 4: Authentication" -Foregroundcolor White -BackgroundColor Blue

Write-Host ""
Write-Host "When prompted - please authenticate with your permitted account, to connect to Exchange." -ForegroundColor Yellow

Write-Host ""

<#
Exchange Online Authentication
-----------------------------
This section establishes a connection to Exchange Online using modern authentication.

Authentication Features:
- Modern authentication with FIDO2/WebAuthn support
- Multi-factor authentication (MFA) compatibility
- Certificate-based authentication support
- Device-based conditional access compliance

Requirements:
- Valid Exchange Online administrator credentials
- Appropriate Exchange Online permissions (View-Only Organization Management or higher)
- MFA-enabled account (recommended for security)
- Network connectivity to Exchange Online endpoints

Documentation:
- Connect-ExchangeOnline: https://docs.microsoft.com/en-us/powershell/module/exchange/connect-exchangeonline
- Exchange Online PowerShell: https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
- Modern Authentication: https://docs.microsoft.com/en-us/microsoft-365/enterprise/modern-auth-for-office-2013-and-2016
- Exchange Online Permissions: https://docs.microsoft.com/en-us/exchange/permissions-exo/permissions-exo
#>

# Connect to Exchange Online using modern authentication
Connect-ExchangeOnline

#-------------------------------------------------------------------------------------------
# STEP 5: Data Retrieval and Export
#-------------------------------------------------------------------------------------------

<#
Progress Bar Function
--------------------
Creates a visual progress indicator for long-running operations.

Features:
- 50-character progress bar with filled and empty sections
- Percentage display for clear progress indication
- Carriage return for in-place updates

Documentation:
- Write-Host Cmdlet: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/write-host
- Math Class: https://docs.microsoft.com/en-us/dotnet/api/system.math
#>
function Show-ProgressBar {
    param (
        [int]$PercentComplete
    )

    $barLength = 50
    $filledLength = [Math]::Round($barLength * $PercentComplete / 100)
    $filled = '█' * $filledLength
    $empty = '-' * ($barLength - $filledLength)
    $bar = "`r[{0}{1}] {2}%" -f $filled, $empty, $PercentComplete

    Write-Host -NoNewline $bar
}

<#
Quarantine Data Retrieval
-------------------------
Downloads quarantine message records from Exchange Online in paginated format.

Configuration:
- Total Pages: 10 (retrieving up to 10,000 records)
- Page Size: 1,000 records per page
- Entity Type: Email messages only

Process:
1. Initialize progress tracking
2. Iterate through pages of quarantine data
3. Export first page as new CSV file
4. Append subsequent pages to existing CSV file
5. Update progress bar for user feedback

Documentation:
- Get-QuarantineMessage: https://docs.microsoft.com/en-us/powershell/module/exchange/get-quarantinemessage
- Export-Csv Cmdlet: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/export-csv
#>

Write-Host ""
Write-Host "Step 5: Get data" -Foregroundcolor White -BackgroundColor Blue

Write-Host ""
Write-Host "Downloading Quarantine message records"

# Initialize variables for quarantine data retrieval
$totalPages = 10
$pageSize = 1000

# Initialize progress tracking
$progress = 0

Show-ProgressBar -PercentComplete $progress

# Retrieve quarantine data in paginated format for efficient processing
for ($page = 1; $page -le $totalPages; $page++) {
    $isFirstPage = $page -eq 1

    # Retrieve quarantine messages for current page
    $messages = Get-QuarantineMessage -EntityType Email -Page $page -PageSize $pageSize
    
    # Configure CSV export parameters
    $exportParams = @{
        Path              = $filePathQuarantine
        NoTypeInformation = $true
    }

    # Export data - create new file for first page, append for subsequent pages
    if ($isFirstPage) {
        $messages | Export-Csv @exportParams
    }
    else {
        $messages | Export-Csv @exportParams -Append
    }

    # Update progress indicator
    $progress = [int](($page / $totalPages) * 100)
    Show-ProgressBar -PercentComplete $progress
}

Write-Host ""
Write-Host "Quarantine Download complete. Please re-run the script if any errors presented."
Write-Host ""

<#
Tenant Allow/Block List Retrieval
---------------------------------
Downloads the Tenant Allow/Block List for sender entries.

Purpose:
- Exports current sender allow/block list configuration
- Provides visibility into email filtering policies
- Enables policy analysis and compliance reporting

List Types Available:
- Sender: Email addresses and domains (used in this script)
- Url: Web URLs and domains
- FileHash: File hash values
- File: File types and extensions

Documentation:
- Get-TenantAllowBlockListItems: https://docs.microsoft.com/en-us/powershell/module/exchange/get-tenantallowblocklistitems
- Tenant Allow/Block Lists: https://docs.microsoft.com/en-us/microsoft-365/security/office-365-security/tenant-allow-block-list
#>

Write-Host "Downloading Tenant Allow/Block List"
Write-Host ""

Get-TenantAllowBlockListItems -ListType Sender | Export-Csv -Path $filePathTABL -NoTypeInformation

#-------------------------------------------------------------------------------------------
# STEP 6: Script Completion and Clean Exit
#-------------------------------------------------------------------------------------------

<#
Script Completion Process
------------------------
This section provides user feedback and ensures clean script termination.

Features:
- Success confirmation with clear visual indication
- File location reporting for easy access to generated data
- User-controlled script termination
- Proper process cleanup for different PowerShell hosts

Process Termination Methods:
- ConsoleHost: Uses Stop-Process for immediate termination
- Other hosts: Uses exit command for graceful shutdown

Documentation:
- Write-Host Cmdlet: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/write-host
- Read-Host Cmdlet: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/read-host
- Stop-Process Cmdlet: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/stop-process
- $host Variable: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_automatic_variables#host
#>

Write-Host "Download complete" -Foregroundcolor White -BackgroundColor Green

Write-Host ""
Write-Host "Report data available at:"
Write-Host $filePathQuarantine
Write-Host $filePathTABL

Write-Host ""
Write-Host "Press the Enter key to end and exit the script..."
Read-Host

# Clean script termination based on PowerShell host environment
if ($host.Name -eq 'ConsoleHost') {
    Stop-Process -Id $PID
}
else {
    exit
}