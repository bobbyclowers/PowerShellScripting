# User Creation Script
# v1.9
# Created by 
# Last updated 6/06/2025, by Michael Harris
# See Changelog for update history
#
#---------------------------------------------------------------------
# Purpose: This script automates the creation of user accounts in Active Directory (AD) with a GUI interface.
# The script collects user information, validates the input, generates a random password, creates the AD user, 
# and configures additional resources like mailbox, groups, employee ID, and org structure.
#---------------------------------------------------------------------
#
# Changelog
# - 20/06/2025: Start work on bulk import
# - 4/03/2025: Adjust List of departments to choose from, replace & in Children and Young People with the word and, to deal with flow-on effects to dynamic group memberships.
# - 29/01/2024: Add missing department name - Homelessness Support.
# - 27/03/2025: Clear Base User and Clear All User buttons added, to permit ease of reuse when creating multiple users - either with the same information or different information.
# - 21/05/2025: Add tests for required and unnecessary base groups, and adding/removing these from user account where needed.
# - 6/06/2025: Write-Host added to show that copy of groups from mirror has been performed, Improvements to code commenting for visibility.

# Set temporary environment path and load necessary modules
$env:tmp = "C:\Temp"  # Set temporary folder for module redirection

# Variables for use of Excel file on SharePoint
$sharePointFilePath = "/sites/SiteName/Shared Documents/Provisioning.xlsx"
$localExcelPath = "C:\Temp\Provisioning.xlsx"

#-------------------------------------------------------------------------------------------
# Common functions
#-------------------------------------------------------------------------------------------

# Function to validate required modules are installed, available, and address if not
function Ensure-Module {
    param(
        [string]$ModuleName,
        [string]$InstallName = $ModuleName
    )
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "Installing module '$InstallName'..."
        Install-Module -Name $InstallName -Force -Scope CurrentUser -AllowClobber
    }
    Import-Module $ModuleName -ErrorAction Stop
}

# Function for creating new users
function New-User {
    param (
        [string]$FirstName,
        [string]$LastName,
        [string]$EmployeeID,
        [string]$ManagerName,
        [string]$MirrorUserRaw,
        [string]$JobTitle,
        [string]$Department
    )

    $FullName = "$FirstName $LastName"
    $LogonName = "$FirstName.$LastName"
    $EmailAddress = "$LogonName@domain.org.au"
    $MirrorUser = $MirrorUserRaw -replace ' ', '.'
    $Manager = $ManagerName -replace ' ', '.'
    $Password = Generate-Password | ConvertTo-SecureString -AsPlainText -Force
    $PasswordPlainText = [System.Net.NetworkCredential]::new('', $Password).Password
    $Path = "OU=User Accounts,OU=Accounts,OU=ADOUName,DC=Domain,DC=local"

    $UserParams = @{
        Name                  = $FullName
        DisplayName           = $FullName
        GivenName             = $FirstName
        SurName               = $LastName
        AccountPassword       = $Password
        Enabled               = $true
        ChangePasswordAtLogon = $false
        Path                  = $Path
        UserPrincipalName     = $EmailAddress
        Title                 = $JobTitle
        Description           = $JobTitle
        Department            = $Department
        Manager               = $Manager
        SamAccountName        = $LogonName
        EmailAddress          = $EmailAddress
        employeeID            = $EmployeeID
        company               = "CompanyName"
    }

    try {
        New-ADUser @UserParams

        if ($MirrorUserRaw -ne "") {
            Get-ADUser -Identity $MirrorUser -Properties MemberOf | Select-Object -ExpandProperty MemberOf | Add-ADGroupMember -Members $LogonName
        }

        Set-ADUser $LogonName -HomeDrive "U:" -HomeDirectory "\\servername\users\$LogonName"
        New-Item -Path "\\servername\users\$LogonName" -ItemType Directory -Force

        $RoutingAddress = "$LogonName@domain.mail.onmicrosoft.com"
        try {
            Enable-RemoteMailbox -Identity $FullName -RemoteRoutingAddress $RoutingAddress -ErrorAction Stop
        }
        catch {
            Write-Warning "Remote mailbox creation failed for $FullName but continuing. Error: $_"
        }

        return @{ LogonName = $LogonName; Email = $EmailAddress; Password = $PasswordPlainText }
    }
    catch {
        Write-Warning "Error creating user $FullName: $_"
        return $null
    }
}

#-------------------------------------------------------------------------------------------
# Test for required modules functions
#-------------------------------------------------------------------------------------------

# Active Directory module: https://learn.microsoft.com/powershell/module/activedirectory/
Ensure-Module -ModuleName "ActiveDirectory"

# Exchange Online Management: https://learn.microsoft.com/powershell/exchange/connect-to-exchange-online-powershell
Ensure-Module -ModuleName "ExchangeOnlineManagement"

<# Skip as not yet required until bulk import is built, ready, and tested
# ImportExcel module: https://github.com/dfinke/ImportExcel
Ensure-Module -ModuleName "ImportExcel"

# PnP PowerShell for SharePoint: https://pnp.github.io/powershell/
Ensure-Module -ModuleName "PnP.PowerShell" -InstallName "PnP.PowerShell"
#endregion

#region SHAREPOINT CONNECTION AND EXCEL DOWNLOAD
# Connect to SharePoint site interactively
# Docs: https://pnp.github.io/powershell/cmdlets/Connect-PnPOnline.html
Connect-PnPOnline -Url "https://domain.sharepoint.com/sites/SiteName" -Interactive

# Download Excel from SharePoint
# Docs: https://pnp.github.io/powershell/cmdlets/Get-PnPFile.html
Get-PnPFile -Url $sharePointFilePath -Path "C:\Temp" -FileName "Provisioning.xlsx" -AsFile -Force
#endregion
#>

#-------------------------------------------------------------------------------------------
# Test to ensure Exchange OnPrem session is working, and fallback to other means if needed
#-------------------------------------------------------------------------------------------


# Try to get a mailbox
try {
    $tst = get-mailbox firstname.lastname@domainname.com.au -ErrorAction silentlycontinue
}
# If unable to import exchange session
catch {
    $s = new-pssession -ConfigurationName microsoft.exchange -connectionuri http://HQEXMGMT01.ADDomainName.local/powershell
    Import-PSSession $S -allowclobber
}

# Load Windows Forms assembly to create the GUI for the user creation
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

#
# Create a Windows Form for the user interface
#

$form = New-Object Windows.Forms.Form
$form.Text = "User Creation"
$form.Size = New-Object Drawing.Size(500, 600)

# Create form elements (labels, textboxes, combo boxes)
# Label and textbox for First Name input
$labelFirstName = New-Object Windows.Forms.Label
$labelFirstName.Text = "First Name:"
$labelFirstName.Location = New-Object Drawing.Point(20, 20)

$textBoxFirstName = New-Object Windows.Forms.TextBox
$textBoxFirstName.Location = New-Object Drawing.Point(200, 20)
$textBoxFirstName.Size = New-Object Drawing.Size(200, 20)

# Label and textbox for Last Name input
$labelLastName = New-Object Windows.Forms.Label
$labelLastName.Text = "Last Name:"
$labelLastName.Location = New-Object Drawing.Point(20, 60)

$textBoxLastName = New-Object Windows.Forms.TextBox
$textBoxLastName.Location = New-Object Drawing.Point(200, 60)
$textBoxLastName.Size = New-Object Drawing.Size(200, 20)

# Label and textbox for Employee ID input
$labelEmployeeID = New-Object Windows.Forms.Label
$labelEmployeeID.Text = "Employee ID:"
$labelEmployeeID.Location = New-Object Drawing.Point(20, 100)
$labelEmployeeID.Size = New-Object Drawing.Size(180, 20)

$textBoxEmployeeID = New-Object Windows.Forms.TextBox
$textBoxEmployeeID.Location = New-Object Drawing.Point(200, 100)
$textBoxEmployeeID.Size = New-Object Drawing.Size(200, 20)

# Label and textbox for Manager input (Full Name)
$labelManager = New-Object Windows.Forms.Label
$labelManager.Text = "Manager Name:"
$labelManager.Location = New-Object Drawing.Point(20, 140)
$labelManager.Size = New-Object Drawing.Size(180, 20)

$textBoxManager = New-Object Windows.Forms.TextBox
$textBoxManager.Location = New-Object Drawing.Point(200, 140)
$textBoxManager.Size = New-Object Drawing.Size(200, 20)

# Label and textbox for Mirror User input
$labelMirrorUser = New-Object Windows.Forms.Label
$labelMirrorUser.Text = "Mirror User:"
$labelMirrorUser.Location = New-Object Drawing.Point(20, 180)

$textBoxMirrorUser = New-Object Windows.Forms.TextBox
$textBoxMirrorUser.Location = New-Object Drawing.Point(200, 180)
$textBoxMirrorUser.Size = New-Object Drawing.Size(200, 20)

# Label and textbox for Job Title input
$labelJobTitle = New-Object Windows.Forms.Label
$labelJobTitle.Text = "Job Title:"
$labelJobTitle.Location = New-Object Drawing.Point(20, 220)

$textBoxJobTitle = New-Object Windows.Forms.TextBox
$textBoxJobTitle.Location = New-Object Drawing.Point(200, 220)
$textBoxJobTitle.Size = New-Object Drawing.Size(200, 20)

# Dropdown for Department selection (No manual input allowed)
$labelDepartment = New-Object Windows.Forms.Label
$labelDepartment.Text = "Department:"
$labelDepartment.Location = New-Object Drawing.Point(20, 260)

$comboBoxDepartment = New-Object Windows.Forms.ComboBox
$comboBoxDepartment.Location = New-Object Drawing.Point(200, 260)
$comboBoxDepartment.Size = New-Object Drawing.Size(200, 20)
$comboBoxDepartment.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList  # This ensures no manual input

# List of departments to choose from
$comboBoxDepartment.Items.Add("Administration")
$comboBoxDepartment.Items.Add("Assets & Facilities")
$comboBoxDepartment.Items.Add("Children and Young People")
$comboBoxDepartment.Items.Add("Community Housing")
$comboBoxDepartment.Items.Add("Digital & Change Management")
$comboBoxDepartment.Items.Add("Financial Services")
$comboBoxDepartment.Items.Add("Financial Wellbeing")
$comboBoxDepartment.Items.Add("Growth & Engagement")
$comboBoxDepartment.Items.Add("Homelessness Support")
$comboBoxDepartment.Items.Add("Individualised Services")
$comboBoxDepartment.Items.Add("Office of the CEO")
$comboBoxDepartment.Items.Add("People Services")
$comboBoxDepartment.Items.Add("Service Governance")
$comboBoxDepartment.Items.Add("Strengthening Families")
$comboBoxDepartment.Items.Add("Transitional Housing Support")

# Create a button for creating the user
$buttonCreateUser = New-Object Windows.Forms.Button
$buttonCreateUser.Text = "Create User"
$buttonCreateUser.Location = New-Object Drawing.Point(200, 300)

# Create a button for clearing selected fields
$buttonClearBase = New-Object Windows.Forms.Button
$buttonClearBase.Text = "Clear Base User Details"
$buttonClearBase.Location = New-Object Drawing.Point(200, 350)
$buttonClearBase.Size = New-Object System.Drawing.Size(150, 25)

# Create a button for clearing all fields
$buttonClearAll = New-Object Windows.Forms.Button
$buttonClearAll.Text = "Clear All Fields"
$buttonClearAll.Location = New-Object Drawing.Point(200, 400)
$buttonClearAll.Size = New-Object System.Drawing.Size(150, 25)

# Create a bulk create button
$buttonBulkCreate = New-Object System.Windows.Forms.Button
$buttonBulkCreate.Text = "Bulk Create Users from New Hire Sheet"
$buttonBulkCreate.Location = New-Object Drawing.Point(80, 450)
$buttonBulkCreate.Size = New-Object Drawing.Size(320, 30)
$form.Controls.Add($buttonBulkCreate)

# Create a progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(40, 500)
$progressBar.Size = New-Object System.Drawing.Size(400, 20)
$progressBar.Minimum = 0
$progressBar.Value = 0
$progressBar.Step = 1
$form.Controls.Add($progressBar)

#
# Generate a Random Password Function
#

# Dot-source the password generator script to load the function
. "V:\Scripts\Saved Scripts\Password-Generator-Silent.ps1"

# Call the function to generate a password
$password = Generate-Password | ConvertTo-SecureString -AsPlainText -Force

#
# Windows forms event handlers
#

#
# Event Handler for "Clear Base User Details" Button
#

$buttonClearBase.Add_Click({
        # Actions to go here
        $textBoxFirstName.Text = ''
        $textBoxLastName.Text = ''
        $textBoxEmployeeID.Text = ''
        $textBoxFirstName.Focus()
    })

#
# Event Handler for "Clear All User Details" Button
#

$buttonClearAll.Add_Click({
        # Actions to go here
        $textBoxFirstName.Text = ''
        $textBoxLastName.Text = ''
        $textBoxEmployeeID.Text = ''
        $textBoxManager.Text = ''
        $textBoxMirrorUser.Text = ''
        $textBoxJobTitle.Text = ''
        $comboBoxDepartment.Text = ""
        $textBoxFirstName.Focus()

    })


#
# User creation process
# Event Handler for "Create User" Button
#

$buttonCreateUser.Add_Click({
        # Retrieve input values from the form
        $firstname = $textBoxFirstName.Text
        $lastname = $textBoxLastName.Text
        $employeeid = $textBoxEmployeeID.Text
        $manager = $textBoxManager.Text
        $mirroruserinput = $textBoxMirrorUser.Text
        $jobtitle = $textBoxJobTitle.Text
        $department = $comboBoxDepartment.SelectedItem

        #
        # Validate user input to ensure all fields are filled
        #

        if ($firstname -eq "" -or $lastname -eq "" -or $employeeid -eq "" -or $manager -eq "" -or $mirroruserinput -eq "" -or $jobtitle -eq "" -or $department -eq $null) {
            [System.Windows.Forms.MessageBox]::Show("Please fill in all fields.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        #
        # Proceed with user creation process
        #

        try {
            # Define full name and other derived fields for AD creation
            $fullname = $firstname + " " + $lastname
            $logonname = $firstname + "." + $lastname
            $emailaddress = $logonname + "@domainname.com.au"
            $mirroruser = ($mirroruserinput.replace(" ", "."))

            # Format the manager name to logon name
            $manager = ($manager.replace(" ", "."))

            # Define the AD path for user creation
            $path = "OU=User Accounts,OU=Accounts,OU=ADOUName,DC=ADDCName,DC=local"

            # Set up user creation parameters for AD
            $fields = @{
                "Name"                  = $fullname
                "DisplayName"           = $fullname
                "GivenName"             = $firstname
                "SurName"               = $lastname
                "AccountPassword"       = $password
                "Enabled"               = $true
                "ChangePasswordAtLogon" = $false
                "Path"                  = $path
                "UserPrincipalName"     = $emailaddress
                "Title"                 = $jobtitle
                "Description"           = $jobtitle
                "Department"            = $department
                "Manager"               = $manager
                "SamAccountName"        = $logonname
                "EmailAddress"          = $emailaddress
                "employeeID"            = $employeeid
                "company"               = "Company Name"
            }

            Import-Module ExchangeOnlineManagement
            # Create the AD user account
            New-ADUser @fields
            Write-Host "AD Account for $fullname has been created."

            #
            # Copy group memberships from the mirror user
            #

            get-ADuser -identity $mirroruser -properties memberof | select-object memberof -expandproperty memberof | Add-AdGroupMember -Members $logonname
            Write-Host "Group memberships of $mirroruser have been copied to $fullname."

            #
            # Set up user folder and home directory in file share
            #

            Set-ADUser $logonname -HomeDrive "U:" -HomeDirectory "\\hqfile01\users\$logonname"
            New-Item -Path "\\hqfile01\users\$logonname" -ItemType Directory -Force

            #
            # Enable the mailbox for the user
            #

            Write-Host "Enabling online mailbox for $logonname..."
            $routingaddress = "$logonname@ucwest.mail.onmicrosoft.com"
            Enable-RemoteMailbox -Identity $fullname -RemoteRoutingAddress $routingaddress
            Write-Host "Online mailbox has been enabled for $fullname"
            Write-Host "User creation completed."
            remove-pssession -session $S

            #
            # Validate user is a member of all required base groups, and remove all non-default signatures
            #

            Write-Host "Checking base group memberships for $fullname"

            #
            # Define user and list of groups to add the user to
            #

            $userSamAccountName = $logonname
            $groupNamesAdd = @("DL All Staff", "UCW-Wireless", "UCW-AllStaff", "Signature - Default", "KnowBe4 - Membership", "U Drive Removal")

            #
            # Get the user object
            #

            $user = Get-ADUser -Identity $userSamAccountName

            # Test user for membership of all groupNames, and add if missing
            foreach ($groupName in $groupNamesAdd) {
                try {
                    # Get the group object
                    $group = Get-ADGroup -Identity $groupName

                    # Check if user is a member of the group
                    $isMember = Get-ADGroupMember -Identity $group -Recursive |
                    Where-Object { $_.DistinguishedName -eq $user.DistinguishedName }

                    if (-not $isMember) {
                        # Add the user if not a member
                        Add-ADGroupMember -Identity $group -Members $user
                        Write-Host "$userSamAccountName added to group '$groupName'."
                    }
                    else {
                        Write-Host "$userSamAccountName is already a member of '$groupName'."
                    }

                }
                catch {
                    Write-Warning "Error processing addition to group '$groupName': $_"
                }
            }

            #
            # Test user for member of non-Standard signature groups, and remove if present
            #

            # Define the user and the list of groups to remove them from
            $userSamAccountName = $logonname
            $groupNamesRemove = @("Signature - PuP", "Signature - EVP")

            # Get the user object
            $user = Get-ADUser -Identity $userSamAccountName

            foreach ($groupName in $groupNamesRemove) {
                try {
                    # Get the group object
                    $group = Get-ADGroup -Identity $groupName

                    # Check if user is a member of the group
                    $isMember = Get-ADGroupMember -Identity $group -Recursive |
                    Where-Object { $_.DistinguishedName -eq $user.DistinguishedName }

                    if ($isMember) {
                        # Remove user from the group
                        Remove-ADGroupMember -Identity $group -Members $user -Confirm:$false
                        Write-Host "$userSamAccountName removed from group '$groupName'."
                    }
                    else {
                        Write-Host "$userSamAccountName is not a member of '$groupName'. No action taken."
                    }

                }
                catch {
                    Write-Warning "Error processing removal of group '$groupName': $_"
                }
            }


            #
            # Convert SecureString password to plain text for display purposes
            #

            $passwordPlainText = [System.Net.NetworkCredential]::new('', $password).Password

            # Display the password in the first pop-up
            [System.Windows.Forms.MessageBox]::Show("User successfully created!`nUsername: $logonname`nEmail address: $emailaddress`nPassword: $passwordPlainText`nManager: $manager", "User Created", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

            # Copy the password to clipboard
            [System.Windows.Forms.Clipboard]::SetText($passwordPlainText)

            # Display a second pop-up stating the password has been copied
            [System.Windows.Forms.MessageBox]::Show("Password has been copied to clipboard.", "Password Copied", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

        }
        catch {
            # Error handling for any issues during user creation
            [System.Windows.Forms.MessageBox]::Show("An error occurred: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }

        $buttonClearBase.Focus()
    })

#
# Add all form elements to the form
#

$form.Controls.Add($labelFirstName)
$form.Controls.Add($textBoxFirstName)
$form.Controls.Add($labelLastName)
$form.Controls.Add($textBoxLastName)
$form.Controls.Add($labelEmployeeID)
$form.Controls.Add($textBoxEmployeeID)
$form.Controls.Add($labelManager)
$form.Controls.Add($textBoxManager)
$form.Controls.Add($labelMirrorUser)
$form.Controls.Add($textBoxMirrorUser)
$form.Controls.Add($labelJobTitle)
$form.Controls.Add($textBoxJobTitle)
$form.Controls.Add($labelDepartment)
$form.Controls.Add($comboBoxDepartment)
$form.Controls.Add($buttonCreateUser)
$form.Controls.Add($buttonClearBase)
$form.Controls.Add($buttonClearAll)

#
# Show the form
#

$form.ShowDialog()

#
# Dispose of the form when the process is done
#

$form.Dispose()