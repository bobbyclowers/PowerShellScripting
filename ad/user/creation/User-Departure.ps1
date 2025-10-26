# User Departure Script
# v1.1
# Created by 
# Last updated 22/05/2025, by Michael Harris
# See Changelog for update history
#
#---------------------------------------------------------------------
# Purpose: This script automates the depature of user accounts in Active Directory (AD) with a GUI interface.
# The script collects required information, disables the account, moves into the correct OU, removes user from
# DL memberships, updates On Prem Exchange, executes delta sync, and any other actions to offboard.
#---------------------------------------------------------------------
#
# Changelog
# - 22/05/2025: Bulk processing enhancements - Clear All button added, to permit ease of reuse when departing multiple users;
#    set focus to Clear All button after shared mailbox dialogue is actioned, to enable next logical step if script is to be
#    run again; and finally return focus to Full Name text box when Clear All is pressed, to enable rapid typing or pasting
#    of next user to be offboarded.

# Load Windows Forms assembly to create the GUI for the user creation
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create a Windows Form for the user interface
$form = New-Object Windows.Forms.Form
$form.Text = "User Management"
$form.Size = New-Object Drawing.Size(650, 300)  # Increased width for better spacing
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::WhiteSmoke  # Set a neutral background color

# Create form elements (labels, textboxes, combo boxes)
# Create the label for the full name input
$labelFullName = New-Object Windows.Forms.Label
$labelFullName.Location = New-Object Drawing.Point(20, 30)
$labelFullName.Size = New-Object Drawing.Size(160, 30)
$labelFullName.Text = "Full Name or Logon Name:"
$labelFullName.Font = New-Object Drawing.Font("Segoe UI", 10)
$labelFullName.TextAlign = 'MiddleLeft'

# Create the textbox for full name input
$textBoxFullName = New-Object Windows.Forms.TextBox
$textBoxFullName.Location = New-Object Drawing.Point(190, 30)
$textBoxFullName.Size = New-Object Drawing.Size(420, 30)  # Adjusted width for better alignment

# Create the "Depart User" button
$buttonDepart = New-Object Windows.Forms.Button
$buttonDepart.Location = New-Object Drawing.Point(20, 80)
$buttonDepart.Text = "Depart User"
$buttonDepart.Font = New-Object Drawing.Font("Segoe UI", 12)
$buttonDepart.BackColor = [System.Drawing.Color]::LightSteelBlue
$buttonDepart.ForeColor = [System.Drawing.Color]::Black
$buttonDepart.Size = New-Object Drawing.Size(250, 40)
$buttonDepart.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat

# Create a button for clearing all fields
$buttonClearAll = New-Object Windows.Forms.Button
$buttonClearAll.Text = "Clear Form"
$buttonClearAll.Location = New-Object Drawing.Point(20, 140)
$buttonClearAll.Size = New-Object System.Drawing.Size(250, 40)
$buttonClearAll.Font = New-Object Drawing.Font("Segoe UI", 12)
$buttonClearAll.BackColor = [System.Drawing.Color]::LightSteelBlue
$buttonClearAll.ForeColor = [System.Drawing.Color]::Black
$buttonClearAll.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat

# Add controls to the form
$form.Controls.Add($labelFullName)
$form.Controls.Add($textBoxFullName)
$form.Controls.Add($buttonDepart)
$form.Controls.Add($buttonClearAll)
$form.Controls.Add($buttonInactiveMailbox)

# Function to show a custom message box
function Show-CustomMessageBox {
    param (
        [string]$message,
        [string]$title
    )

    $customForm = New-Object Windows.Forms.Form
    $customForm.Text = $title
    $customForm.Size = New-Object Drawing.Size(300, 150)
    $customForm.StartPosition = "CenterScreen"
    $customForm.BackColor = [System.Drawing.Color]::LightGray  # Set a background color for the message box

    $label = New-Object Windows.Forms.Label
    $label.Location = New-Object Drawing.Point(10, 10)
    $label.Size = New-Object Drawing.Size(270, 60)
    $label.Text = $message
    $label.Font = New-Object Drawing.Font("Segoe UI", 10)
    $label.AutoSize = $true
    $label.MaximumSize = New-Object Drawing.Size(270, 60)
    $label.TextAlign = 'MiddleCenter'

    $okButton = New-Object Windows.Forms.Button
    $okButton.Location = New-Object Drawing.Point(100, 80)
    $okButton.Size = New-Object Drawing.Size(100, 30)
    $okButton.Text = "OK"
    $okButton.Font = New-Object Drawing.Font("Segoe UI", 10)
    $okButton.BackColor = [System.Drawing.Color]::DodgerBlue
    $okButton.ForeColor = [System.Drawing.Color]::White
    $okButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $okButton.Add_Click({ $customForm.Close() })

    $customForm.Controls.Add($label)
    $customForm.Controls.Add($okButton)
    $customForm.ShowDialog() | Out-Null
}

#----------DEPARTURE----------
# Function to run the user departure script
function Run-UserDepartureScript {
    # Check and import necessary modules
    #Check-And-ImportModule -moduleName "ExchangeOnlineManagement" -installCommand "Install-Module -Name ExchangeOnlineManagement -Scope #CurrentUser -Force"
    
    # Connect to Exchange Online if not already connected
    if (-not (Get-PSSession | Where-Object { $_.ComputerName -like "*exchange*" })) {
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://HQEXMGMT01.ADDomainName.local/powershell
        Import-PSSession $Session -AllowClobber
    }

    $fullname = $textBoxFullName.Text
    $logonname = $fullname.Replace(" ", ".")
    
    # Test if the user can be found
    $user = Get-ADUser -Identity $logonname -ErrorAction SilentlyContinue
    if ($null -eq $user) {
        Show-CustomMessageBox "User $logonname cannot be found in the domain." "User Not Found"
        return
    }

    # Reset password to default
    Write-Output "Resetting the password to default..."
    Set-ADAccountPassword -Identity $logonname -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "Password#1234" -Force)
    Write-Output "Password reset completed."

    # Disable the AD Account
    Write-Output "Disabling the AD Account..."
    Disable-ADAccount -Identity $logonname
    Write-Output "Account disabled."

    # Hide mailbox from GAL & convert to shared mailbox
    $emailaddress = "$logonname@domainname.com.au"
    Write-Output "Hiding the mailbox from the Global Address List (GAL) and converting to shared mailbox..."
    try {
        $mbx = Get-RemoteMailbox -Identity $emailaddress
        Set-RemoteMailbox -Identity $mbx.Identity -HiddenFromAddressListsEnabled $true -Type Shared
        Write-Output "Mailbox converted to shared and hidden from GAL."
    }
    catch {
        Show-CustomMessageBox "Failed to process the mailbox. Please ensure the Exchange Online module is installed and the mailbox exists." "Mailbox Processing Error"
        return
    }


    # Move OU to disabled
    Write-Output "Moving the account to 'Disabled account' OU..."
    Get-ADUser -Identity $logonname | Move-ADObject -TargetPath "OU=Disabled Accounts,OU=Accounts,OU=ADOUName,DC=ADDCName,DC=local"

    # Remove defined group memberships
    Write-Output "Removing DLs and license AD groups..."
    $groups = (Get-ADUser -Identity $logonname -Properties MemberOf).MemberOf | Where-Object { 
        $_ -like "CN=DL *" -or 
        $_ -eq "CN=e3_e5s_microsoft,OU=Security Groups,OU=Groups,OU=ADOUName,DC=ADDCName,DC=local" -or 
        $_ -eq "CN=e1_office,OU=Security Groups,OU=Groups,OU=ADOUName,DC=ADDCName,DC=local" -or
        $_ -eq "CN=signature - default,OU=Security Groups,OU=Groups,OU=ADOUName,DC=ADDCName,DC=local" -or
        $_ -eq "CN=Security Training - Membership,OU=Security Groups,OU=Groups,OU=ADOUName,DC=ADDCName,DC=local"
    }

    foreach ($group in $groups) {
        Remove-ADGroupMember -Identity $group -Members $logonname -Confirm:$false
    }

    Add-ADGroupMember -Identity "e1_office" -Members $logonname
    Write-Output "Group memberships updated."

    # Run a delta sync
    Write-Output "Running a delta sync..."
    Invoke-Command -ComputerName "ADDCServer" -ScriptBlock { Import-Module ADSync; Start-AdSyncSyncCycle -PolicyType Delta }
    Write-Output "Delta sync completed."

    Show-CustomMessageBox "The user departure script has successfully completed!" "Script Completed"
    Show-CustomMessageBox "Please manually check mailbox has been converted to a shared mailbox for data retention purposes" "Script Completed"

    //TODO: Code to output text for inclusion in a ITSM ticket to document what was done, when, by whom, etc. 

    # Remove the session when done
    $session = Get-PSSession | Where-Object { $_.ComputerName -like "*exchange*" }
    if ($session) {
        Remove-PSSession -Session $session
    }

    # Move focus to Clear Form button, in case user needs to run the form again
    $buttonClearAll.Focus()
}

# Event handlers for buttons
$buttonDepart.Add_Click({ Run-UserDepartureScript })

$buttonClearAll.Add_Click({
        # Actions to go here
        $textBoxFullName.Text = ''
        $textBoxFullName.Focus()
    })

# Show the form
$form.ShowDialog() | Out-Null