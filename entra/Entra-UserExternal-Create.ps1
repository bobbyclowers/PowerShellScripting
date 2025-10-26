<#
User Creation External

Version 2.0

Created: 7/05/2025
Created by: Michael HARRIS

Updated: 20/06/2026
Updated by: Michael Harris

Purpose:
- Gather information needed for creating an External user account on Entra AD
- Gather an expiry date, or set one if not provided, and write this to an available field for the purpose of determining next expiry
- Provide a summary of the actions taken for pasting into the ticket body.
- Permit automated management of external accounts, including manager/sponsor emails when nearing expiry, and at expiry, for good security practice

TODO:
- Update other relevant property fields of account - ticket for request
- Capture additional fields (Ticket for request), Code to record, and Render in ICT SD summary text

Update history:

20/06/2025
- Added functions to test for, and import, the PowerShell 7 Test Module
- Now has multi-user with CSV, and single user mode
- Multi-user now provides visual feedback with a progress bar
- Sponsor name coded in invite
- Started to code some other properties into user information ($jobTitle, $companyName, $expiryDate)
- Some code errors still present at Lines 181 and 201, which are overcome by the prompt for the sponsor display name

4/06/2025
- Now working, with upgrade to support PowerShell 7, providing Webauthn redirect for authenticating accounts with FIDO2 keys
- Will run in PowerShell 7 regardless of how executed
- Simplified required module listing, testing for presence and installing if not
- Added $invitePurpose to capture justification for invite, and inclusion in invite email.

7/05/2025
- Initial release.

#>

# Core and initial variables
$moduleNames = @("Microsoft.Graph")

#-------------------------------------------------------------------------------------------
# Common functions
#-------------------------------------------------------------------------------------------

function Resolve-Sponsor {
    param (
        [Parameter(Mandatory)]
        [string]$InitialSponsorName
    )

    $sponsorName = $InitialSponsorName
    $sponsor = $null

    do {
        $sponsorMatches = Get-MgUser -Filter "displayName eq '$sponsorName'" -ConsistencyLevel eventual

        if ($sponsorMatches.Count -eq 0) {
            Write-Host "❌ Sponsor not found. Please re-enter the sponsor's display name and spelling." -ForegroundColor Yellow
            $sponsorName = Read-Host "Re-enter Sponsor Name (Display Name)"
        }
        elseif ($sponsorMatches.Count -eq 1) {
            $sponsor = $sponsorMatches[0]
        }
        else {
            Write-Host "`nMultiple users found for '$sponsorName':" -ForegroundColor Yellow
            for ($i = 0; $i -lt $sponsorMatches.Count; $i++) {
                $match = $sponsorMatches[$i]
                Write-Host "[$i] $($match.DisplayName) - $($match.UserPrincipalName)"
            }

            do {
                $selection = Read-Host "Enter the number of the correct sponsor"
                $validSelection = $selection -match '^\d+$' -and [int]$selection -ge 0 -and [int]$selection -lt $sponsorMatches.Count
                if (-not $validSelection) {
                    Write-Host "Invalid selection. Enter a number between 0 and $($sponsorMatches.Count - 1)." -ForegroundColor Red
                }
            } while (-not $validSelection)

            $sponsor = $sponsorMatches[$selection]
        }
    } while (-not $sponsor)

    return $sponsor
}

function New-ExternalUserInvite {
    param (
        [Parameter(Mandatory)][string]$FirstName,
        [Parameter(Mandatory)][string]$LastName,
        [Parameter(Mandatory)][string]$CompanyName,
        [Parameter(Mandatory)][string]$EmailAddress,
        [Parameter(Mandatory)][string]$SponsorName,
        [Parameter(Mandatory)][string]$InvitePurpose
    )

    $displayName = "$FirstName $LastName ($CompanyName)"
    $messageBody = "A request has been made by $SponsorName for you to be granted access to Company Name resources for the purpose of $InvitePurpose. Please follow the instructions in this invite to accept and be granted this access. If you have any questions, please reach out to your account sponsor, $SponsorName."

    $inviteParams = @{
        InvitedUserDisplayName  = $displayName
        InvitedUserEmailAddress = $EmailAddress
        InviteRedirectUrl       = "https://myapps.microsoft.com"
        SendInvitationMessage   = $true
        InvitedUserMessageInfo  = @{
            CustomizedMessageBody = $messageBody
        }
    }

    return New-MgInvitation @inviteParams
}

function Import-RequirePwsh7Module {
    $modulePath = "V:\Scripts\Saved Scripts\modules\Require-Pwsh7.ps1"

    Write-Host "$($PSStyle.Foreground.White)$($PSStyle.Background.Blue)Testing for, and loading, required scripts$($PSStyle.Reset)`n"

    Set-ExecutionPolicy RemoteSigned -Scope CurrentUser

    # Check if the file exists
    if (-not (Test-Path -Path $modulePath)) {
        Write-Host "❌ The required module file was not found at: $modulePath" -ForegroundColor Red
        return $false
        Write-Host "Press the Enter key to finish the script..."
        Read-Host
        exit
    }

    # Unblock the file if it's blocked (Zone.Identifier)
    try {
        if (Get-Item $modulePath -Stream 'Zone.Identifier' -ErrorAction SilentlyContinue) {
            Write-Host "⚠️  The script is blocked (from another source). Attempting to unblock..." -ForegroundColor Yellow
            Unblock-File -Path $modulePath
            Write-Host "✅ File unblocked." -ForegroundColor Green
        }
    }
    catch {
        Write-Host "⚠️  Could not determine if file is blocked. Continuing..." -ForegroundColor Yellow
    }

    # Check if execution policy will allow sourcing it
    $policy = Get-ExecutionPolicy -Scope CurrentUser
    if ($policy -in @('Restricted', 'AllSigned')) {
        Write-Host "⚠️  Current execution policy is '$policy' and may prevent script execution." -ForegroundColor Yellow
        Write-Host "You may need to re-run: Set-ExecutionPolicy RemoteSigned -Scope CurrentUser"
        return $false
        Write-Host "Press the Enter key to finish the script..."
        Read-Host
        exit
    }

    # Dot-source the script
    try {
        . "$modulePath"
        Write-Host "✅ Require-Pwsh7 module successfully imported from:`n   $modulePath" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "❌ Failed to import the module: $_" -ForegroundColor Red
        return $false
        Write-Host "Press the Enter key to finish the script..."
        Read-Host
        exit
    }
}



#-------------------------------------------------------------------------------------------
# STEP 0: Start script
#-------------------------------------------------------------------------------------------

Write-Host ""
Write-Host "ENTRA - CREATE EXTERNAL USER" -Foregroundcolor White -BackgroundColor Blue

#-------------------------------------------------------------------------------------------
# STEP 1: Run me in Powershell 7
#-------------------------------------------------------------------------------------------

Write-Host ""
Write-Host "Step 1: Test if running in PowerShell 7" -Foregroundcolor White -BackgroundColor Blue

Write-Host ""

# Check if running in PowerShell 7 (Core)
if (-not (Import-RequirePwsh7Module)) {
    Write-Host "Exiting because the required module could not be loaded." -ForegroundColor Red
    exit 1
}

#-------------------------------------------------------------------------------------------
# STEP 2: Test for required modules
#-------------------------------------------------------------------------------------------

Write-Host ""
Write-Host "Step 3: Testing for required modules" -Foregroundcolor White -BackgroundColor Blue

Write-Host ""

# Test for required modules, and install if needed.
foreach ($moduleName in $moduleNames) {
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        Write-Host "Module '$moduleName' not found. Installing..." -ForegroundColor Yellow
        Install-Module -Name $moduleName -Scope CurrentUser -Force
    }
    else {
        Write-Host "Module '$moduleName' is already installed." -ForegroundColor Green
    }
}

#-------------------------------------------------------------------------------------------
# STEP 3: Authenticate
#-------------------------------------------------------------------------------------------

Write-Host ""
Write-Host "Step 2: Authentication" -Foregroundcolor White -BackgroundColor Blue

Write-Host ""
Write-Host "When prompted - please authenticate with your permitted account, to connect to Entra." -ForegroundColor Yellow

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.Read.All", "User.Invite.All" -NoWelcome

#-------------------------------------------------------------------------------------------
# STEP 4: Get user details
#-------------------------------------------------------------------------------------------

# STEP 4A: Render menu for user creation mode
Write-Host "\nStep 4: Choose user creation mode" -ForegroundColor White -BackgroundColor Blue
Write-Host ""
Write-Host "[1] Create Single User"
Write-Host "[2] Create Multiple Users via CSV"

$choice = Read-Host "Select an option [1 or 2]"

#-------------------------------------------------------------------------------------------
# STEP 4B: Multi-user mode
#-------------------------------------------------------------------------------------------

if ($choice -eq "2") {
    # MULTIPLE USERS MODE
    $csvPath = "V:\Scripts\Saved Scripts\csv-input\user-creation-external.csv"

    Write-Host ""
    Write-Host "Checking if CSV file is present and available"
    Write-Host ""

    # 4B: Validate file exists and handle lock retry
    if (!(Test-Path $csvPath)) {
        Write-Host "❌ Error - CSV file not found at $csvPath" -ForegroundColor Red
        Write-Host "Press the Enter key to finish the script..."
        Read-Host
        exit
    }

    $fileIsLocked = $true
    while ($fileIsLocked) {
        try {
            $fileStream = [System.IO.File]::Open($csvPath, 'Open', 'Read', 'None')
            $fileStream.Close()
            $fileIsLocked = $false
        }
        catch {
            Write-Host "❌ Error - CSV file is currently locked or in use." -ForegroundColor Red
            $retry = Read-Host "Press Enter to retry or type 'exit' to quit"
            if ($retry -eq 'exit') { exit }
        }
    }

    # 2b: Import and filter records
    Write-Host ""
    Write-Host "Importing CSV file"
    Write-Host ""
    
    $csvData = Import-Csv -Path $csvPath
    $pendingUsers = $csvData | Where-Object { $_.creationStatus -ne "Complete and Invited" }

    if ($pendingUsers.Count -eq 0) {
        Write-Host "❌ Error - No new users to create (all users marked as Complete and Invited)." -ForegroundColor Red
        Write-Host "Press the Enter key to finish the script..."
        Read-Host
        exit
    }

    $createdBy = (Get-MgUser -UserId ((Get-MgContext).Account)).UserPrincipalName

    # 2c: Create users

    Write-Host ""
    Write-Host "Creating users - please wait"
    Write-Host ""

    $total = $pendingUsers.Count
    $i = 0

    foreach ($user in $pendingUsers) {
        try {
            # Setup the progress bar for visual progress and activity
            $i++
            $percentComplete = [math]::Round(($i / $total) * 100)

            # Update the progress bar when loop commences on next user
            Write-Progress -Activity "Inviting users" -Status "Processing $i of $total ($($user.emailAddress))" -PercentComplete $percentComplete

            # Get user information from pendingUsers, and populate values for user in script
            $firstName = $user.firstName
            $lastName = $user.lastName
            $jobTitle = $user.jobTitle
            $companyName = $user.companyName
            $emailAddress = $user.emailAddress
            $invitePurpose = $user.invitePurpose
            $sponsorName = $user.sponsor
            $expiryDate = $user.expiryDate

            # Get and validate expirydate or default to 90 days
            do {
                $expiryInput = $expiryDate
                if ([string]::IsNullOrWhiteSpace($expiryInput)) {
                    $expiryDate = (Get-Date).AddDays(90).Date.AddHours(23).AddMinutes(59)
                    $validExpiry = $true
                }
                else {
                    try {
                        $expiryDate = [datetime]::ParseExact($expiryInput, 'dd/MM/yyyy', $null).AddHours(23).AddMinutes(59)
                        $validExpiry = $true
                    }
                    catch {
                        Write-Host "Invalid expiry date format. Please use dd/MM/yyyy." -ForegroundColor Red
                        $validExpiry = $false
                    }
                }
            } while (-not $validExpiry)

            # Resolve sponsor
            # Note: There is some errors here that need to be addressed
            $sponsor = Resolve-Sponsor -InitialSponsorName $user.sponsor

            # Setup the contents of the Entra invite to be sent
            $invite = New-ExternalUserInvite `
                -FirstName $firstName `
                -LastName $lastName `
                -CompanyName $companyName `
                -EmailAddress $emailAddress `
                -SponsorName $sponsor.DisplayName `
                -InvitePurpose $invitePurpose

            # Update record with details, and assign Manager to user
            Set-MgUser -UserId $invite.InvitedUser.Id `
                -GivenName $firstName `
                -Surname $lastName `
                -JobTitle $jobTitle `
                -CompanyName $companyName `
                -EmployeeHireDate $expiryDate `
                -EmployeeType "External" `
                -Manager $sponsor `
                -Skills @($invitePurpose) `
                -AgeGroup "Adult" `
                -ConsentProvidedForMinor "NotRequired" `
                -UsageLocation "AU"

            # Assign the Sponsor as the Manager for this user
            # Note: There is some errors here that need to be addressed
            if ($sponsor) {
                Update-MgUser -UserId $invite.InvitedUser.Id -Manager @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$sponsor" }
            }

            $user.expiryDate = $expiryDate
            $user.creationStatus = "Complete and Invited"
            $user.creationDate = (Get-Date -Format "dd/MM/yyyy")
            $user.createdBy = $createdBy

            Write-Host "User $emailAddress invited successfully." -ForegroundColor Green
        }
        catch {
            # Document failure reason to CSV, and who ran the script at the time this user's creation failed
            $user.creationStatus = "Failed - $_"
            $user.createdBy = $createdBy

            # Render failure reason
            Write-Host "❌ Failed to invite user $($user.emailAddress): $_" -ForegroundColor Red
        }
    }

    # Clear progress bar
    Write-Progress -Activity "Inviting users" -Completed

    $csvData | Export-Csv -Path $csvPath -NoTypeInformation

    # Display Summary
    Write-Host ""
    Write-Host "Bulk user creation complete" -Foregroundcolor Green -BackgroundColor Black

    Write-Host ""
    Write-Host "Review the CSV file for confirmation of all actions"
    Write-Host "\n"

    Write-Host "Press the Enter key to finish the script..."
    Read-Host

    exit
}


#-------------------------------------------------------------------------------------------
# STEP 4B: Single user mode
#-------------------------------------------------------------------------------------------

# Get information about user being created
Write-Host "\nStep 4a: Enter details of user to be created" -ForegroundColor White -BackgroundColor Blue
$firstName = Read-Host "Enter First Name"
$lastName = Read-Host "Enter Last Name"
$jobTitle = Read-Host "Enter Job Title"
$companyName = Read-Host "Enter Company Name"
$emailAddress = Read-Host "Enter Email Address"
$invitePurpose = Read-Host "Enter details of the purpose for this External account being created (Note: This is shown to the user, do not add any full stops at the end of this text)"
$sponsorName = Read-Host "Enter Sponsor Name (Display Name)"

# Get and validate expirydate or default to 90 days

do {
    $expiryInput = Read-Host "Enter Expiry Date (dd/mm/yyyy) or leave blank for default 90 day expiry"
    if ([string]::IsNullOrWhiteSpace($expiryInput)) {
        $expiryDate = (Get-Date).AddDays(90).Date.AddHours(23).AddMinutes(59)
        $validExpiry = $true
    }
    else {
        try {
            $expiryDate = [datetime]::ParseExact($expiryInput, 'dd/MM/yyyy', $null).AddHours(23).AddMinutes(59)
            $validExpiry = $true
        }
        catch {
            Write-Host "Invalid expiry date format. Please use dd/MM/yyyy." -ForegroundColor Red
            $validExpiry = $false
        }
    }
} while (-not $validExpiry)


# Resolve sponsor
$sponsor = Resolve-Sponsor -InitialSponsorName $sponsorName

Write-Host ""
Write-Host "Valid sponsor found, continuing"
Write-Host ""

# Compose invite
$invite = New-ExternalUserInvite `
    -FirstName $firstName `
    -LastName $lastName `
    -CompanyName $companyName `
    -EmailAddress $emailAddress `
    -SponsorName $sponsor.DisplayName `
    -InvitePurpose $invitePurpose


# Update record with details, and assign Manager to user
Set-MgUser -UserId $invite.InvitedUser.Id `
    -GivenName $firstName `
    -Surname $lastName `
    -JobTitle $jobTitle `
    -CompanyName $companyName `
    -EmployeeHireDate $expiryDate `
    -EmployeeType "External" `
    -Manager $sponsor `
    -Skills @($invitePurpose) `
    -AgeGroup "Adult" `
    -ConsentProvidedForMinor "NotRequired" `
    -UsageLocation "AU"

# STEP 8: Summary Output
$summaryOutput = @"
Display Name: $firstName $lastName ($companyName)
Name: $firstName $lastName
Email: $emailAddress
Job Title: $jobTitle
Company: $companyName
Sponsorship purpose: $invitePurpose
Sponsor: $sponsorName
Expiry Date: $expiryDate
"@

# Copy contents of $summary to clipboard
$summary | Set-Clipboard

# Display Summary
Write-Host ""
Write-Host "Step 8: Summary" -Foregroundcolor White -BackgroundColor Blue

Write-Host ""
Write-Host "User has been created, and invite sent."
Write-Host "Information below has been copied to the clipboard. Please paste this information into your Service Desk ticket, for confirmation of steps completed:"
Write-Host "\n"
Write-Host $summaryOutput

Write-Host "Press the Enter key to finish the script..."
Read-Host
exit

# Documentation references:
# - File stream check: https://learn.microsoft.com/en-us/dotnet/api/system.io.file.open
# - New-MgInvitation: https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.users/new-mginvitation
# - Set-MgUser (which is an alias of Update-MgUser): https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.users/Update-MgUser
# - Get-MgUser: https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.users/get-mguser
