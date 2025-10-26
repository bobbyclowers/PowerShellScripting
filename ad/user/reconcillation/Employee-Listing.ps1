# Employee listing
# v1.3
# Created by 
# Last updated 20/05/2025, by Michael Harris
# See Changelog for update history
#
#---------------------------------------------------------------------
# Purpose: Take CSV file provided by People Services, and update Employee ID's and Manager relationships in Active Directory
#---------------------------------------------------------------------
#
# TODO:
# - Find a way to test that CSV file being written to is free, before proceeding, to prevent errors from writing to file.
#
# Changelog
# - 1/04/2025: Add options to pre-populate file name ised if only a single CSV present, and allow selection if multiple CSV's present; move used file to Completed folder.
# - 20/05/2025: Add error trapping to ignore blank EmployeeID's, 2nd heading row; Log output of issues for follow up to file with current day's date; add user alerts for further actions based on results.
# - 22/05/2025: Add method for checking that report and csv files are unlocked before proceeding with next action.

# Import Active Directory module
Import-Module ActiveDirectory


#-------------------------------------------------------------------------------------------
# STEP 0: Get file name for script
#-------------------------------------------------------------------------------------------

Write-Host ""
Write-Host "Employee Listing (iChris to AD)" -Foregroundcolor White -BackgroundColor Blue
Write-Host ""

# Prompt for file name
Write-Host "Save the CSV to be used in 'V:\Employee-Listing'." -ForegroundColor Yellow
Write-Host "Note: The file must contains columns 'detnumber' and 'detcurman'." -ForegroundColor Yellow
Write-Host ""

####
# Get file name
####

Write-Host "Locating relevant file in folder" -Foregroundcolor White -BackgroundColor Blue
Write-Host ""

# Define the directory and file type
$directory = "V:\Employee-Listing\"
$fileType = "*.csv"  # Change to the file type you need
$DestinationFolder = "V:\Employee-Listing\Complete"
# Set location and file name for writing report
$DestinationFolderReport = "V:\Employee-Listing\Reports"
$today = Get-Date -Format "yyyyMMdd"
$DestinationFolderReport = Join-Path $DestinationFolderReport "$today.csv"

# Search for files, sort by last modified date (newest first)
$files = Get-ChildItem -Path $directory -Filter $fileType | Sort-Object LastWriteTime -Descending

# Check if any files were found
if ($files.Count -eq 0) {
    Write-Host "No files found."
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

    # Prompt user for selection
    do {
        Write-Host ""
        $selection = Read-Host "Enter the number of the file you want to select"
    } while ($selection -notmatch '^\d+$' -or [int]$selection -lt 0 -or [int]$selection -ge $files.Count)

    # Get the selected file
    $selectedFile = $files[[int]$selection]
    Write-Host ""
    Write-Host "You selected: $($selectedFile.Name)"
}

# Output the selected file path
# $selectedFile.FullName

# Manually enter file name
# Write-Host "Please enter the name of the CSV file (e.g./ staff.csv)." -ForegroundColor Cyan
# No need to add a colon after the Read-Host text, as this is automatically added
# $fileName = Read-Host "Enter file name"

####
# Update csvPath variable for selected file
####

# Specify file location
$csvPath = $selectedFile.FullName

# Check if the file exists
if (-not (Test-Path -Path $csvPath)) {
    Write-Host "Error: File '$fileName' not found in V:\Employee-Listing. Please check and try again." -ForegroundColor Red
    exit
}

Write-Host ""
Write-Host "Processing file" -Foregroundcolor White -BackgroundColor Blue
Write-Host ""

# Import the CSV file
$csvData = Import-Csv -Path $csvPath

# Initialize an array to store skipped users
$skippedUsers = @()

#-------------------------------------------------------------------------------------------
# STEP 1: Add employee ID for all staff members from email address
#-------------------------------------------------------------------------------------------

# Iterate through each row in the CSV
foreach ($row in $csvData) {
    # Get the Employee ID and Email from the CSV
    $employeeID = $row.detnumber
    $email = $row.detemailad
    # Get other variables relevant for CSV report
    $lastName = $row.detsurname
    $preferredName = $row.detprefnm
    $firstName = $row.detg1name1

    # If the row is the second header row, skip it
    if ($employeeID -eq "Staff Member") {
        Write-Host "Skipping row: EmployeeID is empty."
        continue  # Skip this iteration
    }

    # If the Employuee ID is blank, skip it
    if ([string]::IsNullOrWhiteSpace($employeeID)) {
        Write-Host "Skipping row: EmployeeID is empty."
        continue  # Skip this iteration
    }

    # Find the AD user based on the email address
    $user = Get-ADUser -Filter {EmailAddress -eq $email} -Properties EmployeeID

    if ($user) {
        # If the user exists, set the EmployeeID attribute
        Set-ADUser -Identity $user -EmployeeID $employeeID
        Write-Host "Updated EmployeeID for $email to $employeeID"
    } else {
        # Display result
        Write-Host "No AD user found for $email"

        # Step 2c: Check for Report file and create if not exists, with appropriate column headings
        if (!(Test-Path $DestinationFolderReport)) {
            "EmployeeIDPS,Email,LastName,PreferredName,FirstName,ActionReason" | Out-File -FilePath $DestinationFolderReport -Encoding UTF8
        }

        # Step 2d: Append info of user not found to CSV
        $output = "$employeeID,$email,$lastName,$preferredName,$firstName,Not Found in AD"
        Add-Content -Path $DestinationFolderReport -Value $output

        # Step 2e: Set flag
        $infoWrittenToCSV = $true

        <# Check file lock removed before moving on
        # Wait-ForFileUnlock -FilePath $DestinationFolderReport -TimeoutSeconds 10
                
        try {
                $stream = [System.IO.File]::Open($DestinationFolderReport, 'Open', 'Write', 'None')
                $stream.Close()
                #Write-Output "File is unlocked and ready to be written to."
                return $true
            } catch {
                #Write-Output "File is locked or not accessible for writing."
                return $false
            }
        #>

    }
}

# Show alert to user based on outcome

if ($infoWrittenToCSV = $true) {
    Write-Host ""
    Write-Host "❌ Error - Unable to update employee ID for some users" -Foregroundcolor White -BackgroundColor Red
    Write-Host ""
    Write-Host "View the CSV report file located in:"
    Write-Host $DestinationFolderReport
    Write-Host ""
    Write-Host "To see the list of users whose AD account may not have the correct EmployeeID recorded/updated, and take appropriate action."
    Write-Host ""
    Write-Host "Press Enter to proceed with the script, and update Manager relationships..."
    $infoWrittenToCSV = $false
    Read-Host
    Write-Host ""
} else {
    Write-Host ""
    Write-Host "Employee ID update complete" -Foregroundcolor Black -BackgroundColor Green
    Write-Host ""
    Write-Host "All users found, and Employee IDs updated"
    Write-Host ""
    Write-Host "Proceeding to update Manager relationships..."
    Write-Host ""

}

#-------------------------------------------------------------------------------------------
# STEP 2: Update manager field for all employees
#-------------------------------------------------------------------------------------------

# Filter out rows with missing required fields
$processedData = $csvData | Where-Object { $_.detnumber -and $_.detcurman }

foreach ($entry in $processedData) {
    $employeeID = $entry.detnumber
    $managerID = $entry.detcurman

    # If the row is the second header row, skip it
    if ($employeeID -eq "Staff Member") {
        Write-Host "Skipping row: EmployeeID is empty."
        continue  # Skip this iteration
    }

    # If the Employuee ID is blank, skip it
    if ([string]::IsNullOrWhiteSpace($employeeID)) {
        Write-Host "Skipping row: EmployeeID is empty."
        continue  # Skip this iteration
    }

    # Look up the user in Active Directory using EmployeeID
    $user = Get-ADUser -Filter {EmployeeID -eq $employeeID} -Properties EmployeeID, Manager -ErrorAction SilentlyContinue

    if ($user) {
        # Look up the manager in Active Directory using EmployeeID
        $manager = Get-ADUser -Filter {EmployeeID -eq $managerID} -Properties DistinguishedName -ErrorAction SilentlyContinue

        if ($manager) {
            try {
                # Update the user's manager field in AD
                Set-ADUser -Identity $user.SamAccountName -Manager $manager.DistinguishedName

                Write-Host "Successfully updated manager for user '$($user.SamAccountName)' to manager '$($manager.SamAccountName)'" -ForegroundColor Green
            } catch {
                Write-Warning "Failed to update manager for user '$($user.SamAccountName)': $_"
                # Step 2c: Check for Report file and create if not exists, with appropriate column headings
                    if (!(Test-Path $DestinationFolderReport)) {
                        "EmployeeIDPS,Email,LastName,PreferredName,FirstName,ActionReason" | Out-File -FilePath $DestinationFolderReport -Encoding UTF8
                    }

                # Step 2d: Append info of user not found to CSV
                    $output = "$employeeID,$email,$lastName,$preferredName,$firstName,Failed to update Manager for this user"
                    Add-Content -Path $DestinationFolderReport -Value $output

                    # Step 2e: Set flag
                    $infoWrittenToCSV = $true

                    <# Check file lock removed before moving on #>
                    # Wait-ForFileUnlock -FilePath $DestinationFolderReport -TimeoutSeconds 10

            }
        } else {
            Write-Warning "Manager with EmployeeID '$managerID' not found in Active Directory. Skipping user '$($user.SamAccountName)'."
            # Step 2c: Check for Report file and create if not exists, with appropriate column headings
                    if (!(Test-Path $DestinationFolderReport)) {
                        "EmployeeIDPS,Email,LastName,PreferredName,FirstName,ActionReason" | Out-File -FilePath $DestinationFolderReport -Encoding UTF8
                    }

                # Step 2d: Append info of user not found to CSV
                    $output = "$employeeID,$email,$lastName,$preferredName,$firstName,Manager with EmployeeID '$managerID' not found in Active Directory"
                    Add-Content -Path $DestinationFolderReport -Value $output

                    # Step 2e: Set flag
                    $infoWrittenToCSV = $true

                    <# Check file lock removed before moving on #>
                    # Wait-ForFileUnlock -FilePath $DestinationFolderReport -TimeoutSeconds 10
        }
    } else {
        Write-Warning "User with EmployeeID '$employeeID' not found in Active Directory. Skipping."
        # Step 2c: Check for Report file and create if not exists, with appropriate column headings
            if (!(Test-Path $DestinationFolderReport)) {
                "EmployeeIDPS,Email,LastName,PreferredName,FirstName,ActionReason" | Out-File -FilePath $DestinationFolderReport -Encoding UTF8
            }

        # Step 2d: Append info of user not found to CSV
            $output = "$employeeID,$email,$lastName,$preferredName,$firstName,User with EmployeeID '$employeeID' not found in Active Directory"
            Add-Content -Path $DestinationFolderReport -Value $output

            # Step 2e: Set flag
            $infoWrittenToCSV = $true

            <# Check file lock removed before moving on
            Wait-ForFileUnlock -FilePath $DestinationFolderReport -TimeoutSeconds 10
            #>
    }
}


# Show alert to user based on outcome

if ($infoWrittenToCSV = $true) {
    Write-Host ""
    Write-Host "❌ Error - Unable to update Manager for some users" -Foregroundcolor White -BackgroundColor Red
    Write-Host ""
    Write-Host "View the CSV report file located in:"
    Write-Host $DestinationFolderReport
    Write-Host ""
    Write-Host "To view the list of users whose AD account may not have the correct EmployeeID recorded, and take appropriate action."
    Write-Host ""
    Write-Host "Press Enter to proceed to cleanup..."
    $infoWrittenToCSV = $false
    Read-Host
    Write-Host ""
} else {
    Write-Host ""
    Write-Host "Manager to Employee complete" -Foregroundcolor Black -BackgroundColor Green
    Write-Host ""
    Write-Host "All users found, and Managers updated"
    Write-Host ""
    Write-Host "Proceeding to cleanup..."
    Write-Host ""

}

#-------------------------------------------------------------------------------------------
# STEP 3: Cleanup
#-------------------------------------------------------------------------------------------

<# Check file lock removed before moving on #>
# Wait-ForFileUnlock -FilePath $csvPath -TimeoutSeconds 10

Move-Item -Path $csvPath -Destination $DestinationFolder -Force

Write-Host "File used moved into" $DestinationFolder -ForegroundColor Green
