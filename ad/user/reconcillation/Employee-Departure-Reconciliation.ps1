# Employee Departure Reconcilation
# v1.01
# Created by Michael Harris - ICT Analyst
# Last updated 9/0/2025, by Michael Harris
# See Changelog for update history
#
#---------------------------------------------------------------------
# Purpose: Take CSV file provided by People Services, validate if person is still active, deactivate if yes, then alert ICT, also check for any assets still assigned to them
#---------------------------------------------------------------------
#
# Changelog
# - 9/05/2025: Initial creation
# - 22/05/2025: Add file lockout checks, to ensure CSV file is clear for additional writing before proceeding onto next step.

# Import Active Directory module
Import-Module ActiveDirectory


#-------------------------------------------------------------------------------------------
# STEP 0: Start script, get file names, check and select right file
#-------------------------------------------------------------------------------------------

Write-Host ""
Write-Host "Employee Departure Reconcillation (iChris to AD)" -Foregroundcolor White -BackgroundColor Blue
Write-Host ""

# Prompt for file name
Write-Host "Save the CSV to be used in V:\Employee-Departure-Check'." -ForegroundColor Yellow
Write-Host "Note: The file must contains columns 'detnumber'." -ForegroundColor Yellow
Write-Host ""

####
# Get file name
####

Write-Host "Locating relevant file in folder" -Foregroundcolor White -BackgroundColor Blue
Write-Host ""

# Define the directory and file type
$directory = "V:\Employee-Departure-Check"
$fileType = "*.csv"  # Change to the file type you need
$DestinationFolderInput = "V:\Employee-Departure-Check\Complete"
# Set location and file name for writing report
$DestinationFolderReport = "V:\Employee-Departure-Check\Reports"
$today = Get-Date -Format "yyyyMMdd"
$DestinationFolderReport = Join-Path $DestinationFolderReport "$today.csv"
# Define flags to generate specific text if accounts not disabled were found
$infoWrittenToCSV = $false

# Search for files, sort by last modified date (newest first)
$files = Get-ChildItem -Path $directory -Filter $fileType | Sort-Object LastWriteTime -Descending

# Check if any files were found
if ($files.Count -eq 0) {
    Write-Host "❌ Error - No CSV File found." -Foregroundcolor White -BackgroundColor Red
    Write-Host ""
    Write-Host "Check that the CSV file needed has been placed in the correct folder."
    Write-Host "Once placed in the folder, re-run the script."
    Write-Host ""
    Write-Host "Press the Enter key to end the script..."
    Read-Host
    exit
}

####
# Proceed based on how many files found
####

# If only one file is found, use it automatically
if ($files.Count -eq 1) {
    $selectedFile = $files[0]
    Write-Host "Only one file found: $($selectedFile.Name)"
} else {
    # Display files for selection
    Write-Host "Multiple files found. Select a file:"
    Write-Host ""
    for ($i = 0; $i -lt $files.Count; $i++) {
        Write-Host "[$i] $($files[$i].Name) - Last Modified: $($files[$i].LastWriteTime)"
    }

    # Prompt user for selection from menu if multiple files found
    do {
        Write-Host ""
        $selection = Read-Host "Enter the number of the file you want to select"
    } while ($selection -notmatch '^\d+$' -or [int]$selection -lt 0 -or [int]$selection -ge $files.Count)

    # Get the selected file
    $selectedFile = $files[[int]$selection]
    Write-Host ""
    Write-Host "You selected: $($selectedFile.Name)"
}


####
# Update csvPath variable for selected file
####

# Specify file location
$csvPath = $selectedFile.FullName

# Safety checl - verify the chosen file exists
if (-not (Test-Path -Path $csvPath)) {
    Write-Host "❌ Error - Missing CSV File." -Foregroundcolor White -BackgroundColor Red
    Write-Host "File '$fileName' not found in V:\Employee-Departure-Check. Please check and try again." -ForegroundColor Red
    Write-Host ""
    Write-Host "Press the Enter key to end the script..."
    Read-Host
    exit
}

Write-Host ""
Write-Host "Processing file" -Foregroundcolor White -BackgroundColor Blue
Write-Host ""

# Import the CSV file
#$csvData = Import-Csv -Path $csvPath

    try {
        $csvData = Import-Csv -Path $csvPath -ErrorAction Stop
    }
    catch {
        Write-Host "❌ Error - Failed to import CSV file at path '$csvPath'." -Foregroundcolor White -BackgroundColor Red
        Write-Host "Error details:"
        Write-Host "$($_.Exception.Message)"
        Write-Host ""
        Write-Host "Press Enter to end the script..."
        Read-Host
        exit
    }


#-------------------------------------------------------------------------------------------
# STEP 1: Process file, and check for active accounts of any departed staff
#-------------------------------------------------------------------------------------------

# Iterate through each row in the CSV
foreach ($row in $csvData) {
    # Get the Employee ID and Email from the CSV
    $employeeIDPS = $row.detnumber
    $employeeEmail = $row.detemailad
    $employeeDepartureDate = $row.detterdate
    $employeeDivision = $row.pdtorg2cd.trn
    $employeeDepartment = $row.pdtorg3cd.trn
    $employeeTeam = $row.pdtorg4cd.trn

    # Find the AD user based on the email address
    $user = Get-ADUser -Filter {EmployeeID -eq $employeeIDPS} -Properties EmployeeID

    # Step 1: Check if user is active (enabled)
    if ($user -and $user.Enabled) {
        # Step 2a: Disable the user account
        Disable-ADAccount -Identity $user

        # Step 2b: Retrieve required fields
        $employeeIDAD = $user.EmployeeID
        $userPrincipalName = $user.UserPrincipalName
        $displayName = $user.DisplayName

        # Get Manager's name
        $managerDN = $user.Manager
        $managerName = if ($managerDN) {
            (Get-ADUser -Identity $managerDN).Name
        } else {
            "No Manager Listed in AD"
        }

        # Step 2c: Check for Report file and create if not exists
        if (!(Test-Path $DestinationFolderReport)) {
            "EmployeeIDPS,EmployeeIDAD,UserPrincipalName,DisplayName,ManagerName,employeeDepartureDate,employeeDivision,employeeDepartment,employeeTeam" | Out-File -FilePath $DestinationFolderReport -Encoding UTF8
        }

        # Step 2d: Append user info to CSV
        $output = "$employeeIDPS,$employeeIDAD,$userPrincipalName,$displayName,$managerName,$employeeDepartureDate,$employeeDivision,$employeeDepartment,$employeeTeam"
        Add-Content -Path $DestinationFolderReport -Value $output

        # Step 2e: Set flag
        $infoWrittenToCSV = $true

        <# Check file lock removed before moving on #>
        # Wait-ForFileUnlock -FilePath $DestinationFolderReport -TimeoutSeconds 10

    } else {
        # Write-Host "User not found or already disabled."
        # Redundant to display this information
    }
}

#-------------------------------------------------------------------------------------------
# STEP 4: Alert based on results
#-------------------------------------------------------------------------------------------

if ($infoWrittenToCSV -eq $true) {
    Write-Host ""
    Write-Host "❌ ACTION REQUIRED" -Foregroundcolor White -BackgroundColor Red
    Write-Host ""
    Write-Host "At least one employee in the latest departure file from People Services"
    Write-Host "is yet to have their account disabled, due to an Employee Departure form"
    Write-Host "likely not yet being received by the ICT Service Desk."
    Write-Host ""
    Write-Host "These accounts have been disabled pending completion of relevant actions."
    Write-Host ""    
    Write-Host "Instructions for relevant actions can be found in the ICT User Management process document on the ICT SharePoint site."
    Write-Host ""
    Write-Host "Without delay - Please do all of the following steps"
    Write-Host ""
    Write-Host "1) Open the Reports file, located at: "$DestinationFolderReport
    Write-Host "2) Check Intune to see if there are any assets still assigned to the employee"
    Write-Host "3) Check Help Desk for a ticket to see if a departure form has been provided, and"
    Write-Host "3a) Process the departure form if not already actions, or"
    Write-Host "3b) If the ticket states departure was processed, check all actions completed and follow up with who processed it"
    Write-Host "4) If departure checklist ticket not received, or assets remain outstanding"
    Write-Host "4a) Open a new Help Desk ticket, with the user's Manager as the client, CC'ing People Services"
    Write-Host "4b) Send a ticket asking for the Departure Checklist and/or Outstanding assets."
    Write-Host ""
    Write-Host "Press the Enter key to finish the script..."
    Read-Host
}

if ($infoWrittenToCSV -eq $false) {
    Write-Host ""
    Write-Host "NO ACTION REQUIRED" -Foregroundcolor Black -BackgroundColor Green
    Write-Host ""
    Write-Host "No offboarded employees were found, by Employee ID, whose accounts have not been disabled."
    Write-Host ""
    Write-Host "Press the Enter key to finish the script..."
    Read-Host
}

#-------------------------------------------------------------------------------------------
# STEP 4: Cleanup
#-------------------------------------------------------------------------------------------

Move-Item -Path $csvPath -Destination $DestinationFolderInput -Force

Write-Host ""
Write-Host "Post-report cleanup" -Foregroundcolor Black -BackgroundColor Green
Write-Host ""

Write-Host "Report file used moved into" $DestinationFolderInput -ForegroundColor Green
