<#
Get E365 Quarantine Data to CSV for analysis

Created: 10/03/2025
Created by: Michael HARRIS

Updated: 14/04/2025 (In progress)
Updated by: Michael Harris

Purpose:
- Connect to E365
- Download last 10k of quarantine records
- Save to CSV in V:\Exchange-Blocking

Update history:

4/06/2025
- Converted to PowerShell 7 for FIDO2 Webauthn redirection support.
- Adjusted end script metho
14/04/2024
- If existing CSV file(s) in folder, first update file name(s) with creation date as suffix, and move them to archives folder.
- Cease adding creation date to file names created in script as a result.
- Designed to make running of Excel book easier, and not need to update the file name(s) in the Power Query each time.
10/03/2025
- Initial release.
#>

<# Core and initial variables #>

# Define the directory path and the target folder
$directoryPath = "V:\Exchange-Blocking\"
$targetFolder = "V:\Exchange-Blocking\Archived"

# Define the resulting files for the current day's data
$filePathQuarantine = "V:\Exchange-Blocking\quarantine.csv"
$filePathTABL = "V:\Exchange-Blocking\tabl.csv"

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
# STEP 1: Run me in Powershell 7
#-------------------------------------------------------------------------------------------

Write-Host ""
Write-Host "Step 1: Test if running in PowerShell 7" -Foregroundcolor White -BackgroundColor Blue

Write-Host ""

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
    } else {
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
Test for required modules, install if needed, then import the module.
#>

foreach ($moduleName in $moduleNames) {
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        Write-Host "Module '$moduleName' not found. Installing..." -ForegroundColor Yellow
        Install-Module -Name $moduleName -Scope CurrentUser -Force
    } else {
        Write-Host "Module '$moduleName' is already installed." -ForegroundColor Green
    }

    # Import the module
    Write-Host "Importing Module '$moduleName'." -ForegroundColor Green
    Import-Module -Name $moduleName -Force
}

#-------------------------------------------------------------------------------------------
# STEP 3: Clean up old files
#-------------------------------------------------------------------------------------------


Write-Host ""
Write-Host "Step 3: Archive old .CSV files" -Foregroundcolor White -BackgroundColor Blue

Write-Host "Rename previous CSV files in folder, and move to \Archived"
Write-Host ""

####
# Locate, rename and move files to \Archived folder
####

# Get all CSV files in the directory
$csvFiles = Get-ChildItem -Path $directoryPath -Filter *.csv

if ($csvFiles.Count -gt 0) {
    foreach ($file in $csvFiles) {
        # Get the file's creation date
        $creationDate = $file.CreationTime

        # Format the date as you prefer, e.g., YYYYMMDD
        $dateSuffix = $creationDate.ToString("yyyyMMdd")

        # Get the file's name without extension
        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($file)
        $fileExtension = $file.Extension

        # Create the new file name with the date suffix
        $newFileName = "$fileName-$dateSuffix$fileExtension"
        $newFilePath = [System.IO.Path]::Combine($directoryPath, $newFileName)

        # Rename the file
        Rename-Item -Path $file.FullName -NewName $newFileName

        # Move the file to the target folder
        Move-Item -Path $newFilePath -Destination $targetFolder

        Write-Output "File renamed to: $newFileName and moved to $targetFolder"
    }
    Write-Output "All CSV files have been renamed and moved."
} else {
    Write-Output "No CSV files found in the directory."
}


#-------------------------------------------------------------------------------------------
# STEP 4: Authenticate
#-------------------------------------------------------------------------------------------

Write-Host ""
Write-Host "Step 4: Authentication" -Foregroundcolor White -BackgroundColor Blue

Write-Host ""
Write-Host "When prompted - please authenticate with your permitted account, to connect to Exchange." -ForegroundColor Yellow

Write-Host ""

<# Connect to E365 #>

Connect-ExchangeOnline

#-------------------------------------------------------------------------------------------
# STEP 5: Get Data
#-------------------------------------------------------------------------------------------

<# Let's make it fancy, with a progress bar #>
function Show-ProgressBar {
    param (
        [int]$PercentComplete
    )

    $barLength = 50
    $filledLength = [Math]::Round($barLength * $PercentComplete / 100)
    $filled = '█' * $filledLength
    $empty = '-' * ($barLength - $filledLength)
    $bar = "`r[$filled$empty] $PercentComplete%"

    Write-Host -NoNewline $bar
}

<# Download reports, and append to same file after first page #>

Write-Host ""
Write-Host "Step 5: Get data" -Foregroundcolor White -BackgroundColor Blue

Write-Host ""
Write-Host "Downloading Quarantine message records"
# Initialize variables for this section
$totalPages = 10
$pageSize = 1000

# Initialize progress
$progress = 0

Show-ProgressBar -PercentComplete $progress

# Let's make this a for-each loop for some cleaner code

for ($page = 1; $page -le $totalPages; $page++) {
    $isFirstPage = $page -eq 1

    # Retrieve and export messages
    $messages = Get-QuarantineMessage -EntityType Email -Page $page -PageSize $pageSize
    $exportParams = @{
        Path = $filePathQuarantine
        NoTypeInformation = $true
    }

    if ($isFirstPage) {
        $messages | Export-Csv @exportParams
    } else {
        $messages | Export-Csv @exportParams -Append
    }

    # Update progress
    $progress = [int](($page / $totalPages) * 100)
    Show-ProgressBar -PercentComplete $progress
}

Write-Host ""
Write-Host "Querantine Download complete. Please re-run the script if any errors presented."
Write-Host ""

Write-Host "Downloading Tenant Allow/Block List"
Write-Host ""

Get-TenantAllowBlockListItems -ListType Sender | Export-Csv -Path $filePathTABL -NoTypeInformation

#-------------------------------------------------------------------------------------------
# STEP 6: Script complete
#-------------------------------------------------------------------------------------------

Write-Host "Download complete" -Foregroundcolor White -BackgroundColor Green

Write-Host ""
Write-Host "Report data available at:"
Write-Host $filePathQuarantine
Write-Host $filePathTABL

Write-Host ""
Write-Host "Press the Enter key to end and exit the script..."
Read-Host

if ($host.Name -eq 'ConsoleHost') {
    Stop-Process -Id $PID
} else {
    exit
}
