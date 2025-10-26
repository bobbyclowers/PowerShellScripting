<#
User Profile Copy

Wersion 1.3

Created: 30/10/2024
Created by: Michael HARRIS

Updated: 21/05/2025
Updated by: Michael Harris

Purpose:
- Gather from and to user for copying permissions
- Give options to reset password, and remove existing memberships
- Give options to move OU if user is outside of active users

Update history:

12/02/2025
- Sanitise from and to search input, to account for single quotation marks (') breaking the query.

29/01/2025
- Check for missing or incorrect information on To user (Employee ID, Manager Name, Department, etc), and prompt to capture or correct

21/05/2025
- Add tests for required and unnecessary base groups, and adding/removing these from user account where needed (function configureDefaultGroups)
#>

<# Core and initial variables #>

$scriptComplete = "0"

<# Load Core Assemblies #>

# Load the Windows Forms assembly
Add-Type -AssemblyName System.Windows.Forms

<# Step 1: Select and test for user to copy from #>

function findFromUser {
    # Get name input for user to copy FROM
    Write-Host ""
    Write-Host "Step 1: From user (Person copying from)" -Foregroundcolor White -BackgroundColor Blue
    Write-Host ""
    $userNameFromInput = Read-Host "Please enter the name of the user you are looking to match`n(Either Login, User Principal, Email address, or Display Name)"
    Write-Host ""
    Write-Host "You are searching for:" $userNameFromInput
    Write-Host ""

    # Sanitise inputs for single quotes, to prevent breaking script
    $userNameFrom = $userNameFromInput.Replace("'", "''")

    # Match against SamAccountName

    $userNameFromNull = Get-ADUser -Filter "SamAccountName -eq '$userNameFrom'"

    if ($null -eq $userNameFromNull) {
        # write-host "SamAccountName '$userNameFrom' does not yet exist in active directory"
    }
    else {
        # Match found on samAccountName
        Write-Host "Match found on SamAccountName"
        # Get details and write to variables
        $SamAccountNameFrom = Get-ADUser -Filter "SamAccountName -eq '$userNameFrom'" | Select-Object -ExpandProperty samAccountName
        $orgUnitLocationFrom = Get-ADUser -Filter "SamAccountName -eq '$userNameFrom'" | Select-Object -ExpandProperty DistinguishedName
        findToUser
    }

    # Match against UserPrincipalName

    $userNameFromNull = Get-ADUser -Filter "UserPrincipalName -eq '$userNameFrom'"

    if ($null -eq $userNameFromNull) {
        # write-host "UserPrincipalName '$userNameFrom' does not yet exist in active directory"
    }
    else {
        # Match found on UserPrincipalName
        Write-Host "Match found on UserPrincipalName"
        # Get details and write to variables
        $SamAccountNameFrom = Get-ADUser -Filter "UserPrincipalName -eq '$userNameFrom'" | Select-Object -ExpandProperty samAccountName
        $orgUnitLocationFrom = Get-ADUser -Filter "UserPrincipalName -eq '$userNameFrom'" | Select-Object -ExpandProperty DistinguishedName
        # Skip remaining, and jump ahead to specific point
        findToUser
    }

    # Match against EmailAddress

    $userNameFromNull = Get-ADUser -Filter "EmailAddress -eq '$userNameFrom'"

    if ($null -eq $userNameFromNull) {
        # write-host "EmailAddress '$userNameFrom' does not yet exist in active directory"
    }
    else {
        # Match found on EmailAddress
        Write-Host "Match found on EmailAddress"
        # Get details and write to variables
        $SamAccountNameFrom = Get-ADUser -Filter "EmailAddress -eq '$userNameFrom'" | Select-Object -ExpandProperty samAccountName
        $orgUnitLocationFrom = Get-ADUser -Filter "EmailAddress -eq '$userNameFrom'" | Select-Object -ExpandProperty DistinguishedName
        # Skip remaining, and jump ahead to specific point
        findToUser
    }

    # Match against Name

    $userNameFromNull = Get-ADUser -Filter "Name -eq '$userNameFrom'"

    if ($null -eq $userNameFromNull) {
        # write-host "Display Name '$userNameFrom' does not yet exist in active directory"
        noFromMatch
    }
    else {
        # Match found on DisplayName
        Write-Host "Match found on DisplayName"
        # Get details and write to variables
        $SamAccountNameFrom = Get-ADUser -Filter "Name -eq '$userNameFrom'" | Select-Object -ExpandProperty samAccountName
        $orgUnitLocationFrom = Get-ADUser -Filter "Name -eq '$userNameFrom'" | Select-Object -ExpandProperty DistinguishedName
        # Skip remaining, and jump ahead to specific point
        findToUser
    }


}

<# No match for FROM user - Alert and return to findFromUser #>

function noFromMatch {
    # Show a message box
    [System.Windows.Forms.MessageBox]::Show("There is no match for the FROM user you entered. Please check the name, and re-enter when prompted", "PROBLEM: NO MATCH FOR FROM USER", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    findFromUser
}


<# Step 2: Select and test for user to copy TO #>

function findToUser {
    # Get name input for user to copy TO
    Write-Host ""
    Write-Host "Step 2: To user (Person copying to)" -Foregroundcolor White -BackgroundColor Blue
    Write-Host ""
    $userNameToInput = Read-Host "Please enter the name of the user you are looking to match`n(Either Login, User Principal, Email address, or Display Name)"
    Write-Host ""
    Write-Host "You are searching for:" $userNameToInput
    Write-Host ""

    # Sanitise inputs for single quotes, to prevent breaking script
    $userNameTo = $userNameToInput.Replace("'", "''")

    # Match against SamAccountName

    $userNameToNull = Get-ADUser -Filter "SamAccountName -eq '$userNameTo'"

    if ($null -eq $userNameToNull) {
        # write-host "SamAccountName '$userNameTo' does not yet exist in active directory"
    }
    else {
        # Match found on samAccountName
        Write-Host "Match found on samAccountName"
        # Get details and write to variables
        $SamAccountNameTo = Get-ADUser -Filter "SamAccountName -eq '$userNameTo'" | Select-Object -ExpandProperty samAccountName
        $orgUnitLocationTo = Get-ADUser -Filter "SamAccountName -eq '$userNameTo'" | Select-Object -ExpandProperty DistinguishedName
        performActions
    }

    # Match against UserPrincipalName

    $userNameToNull = Get-ADUser -Filter "UserPrincipalName -eq '$userNameTo'"

    if ($null -eq $userNameToNull) {
        # write-host "UserPrincipalName '$userNameTo' does not yet exist in active directory"
    }
    else {
        # Match found on UserPrincipalName
        Write-Host "Match found on UserPrincipalName"
        # Get details and write to variables
        $SamAccountNameTo = Get-ADUser -Filter "UserPrincipalName -eq '$userNameTo'" | Select-Object -ExpandProperty samAccountName
        $orgUnitLocationTo = Get-ADUser -Filter "UserPrincipalName -eq '$userNameTo'" | Select-Object -ExpandProperty DistinguishedName
        # Skip remaining, and jump ahead to specific point
        performActions
    }

    # Match against EmailAddress

    $userNameToNull = Get-ADUser -Filter "EmailAddress -eq '$userNameTo'"

    if ($null -eq $userNameToNull) {
        # write-host "EmailAddress '$userNameTo' does not yet exist in active directory"
    }
    else {
        # Match found on EmailAddress
        Write-Host "Match found on EmailAddress"
        # Get details and write to variables
        $SamAccountNameTo = Get-ADUser -Filter "EmailAddress -eq '$userNameTo'" | Select-Object -ExpandProperty samAccountName
        $orgUnitLocationTo = Get-ADUser -Filter "EmailAddress -eq '$userNameTo'" | Select-Object -ExpandProperty DistinguishedName
        # Skip remaining, and jump ahead to specific point
        performActions
    }

    # Match against Name

    $userNameToNull = Get-ADUser -Filter "Name -eq '$userNameTo'"

    if ($null -eq $userNameToNull) {
        # write-host "Name '$userNameTo' does not yet exist in active directory"
        noToMatch
    }
    else {
        # Match found on DisplayName
        Write-Host "Match found on DisplayName"
        # Get details and write to variables
        $SamAccountNameTo = Get-ADUser -Filter "Name -eq '$userNameTo'" | Select-Object -ExpandProperty samAccountName
        $orgUnitLocationTo = Get-ADUser -Filter "Name -eq '$userNameTo'" | Select-Object -ExpandProperty DistinguishedName
        # Skip remaining, and jump ahead to specific point
        performActions
    }

    <#
Write-Host ""
Write-Host "NO MATCH FOR TO USER" -Foregroundcolor White -BackgroundColor Red
Write-Host "There is no match for the TO user you entered."
Write-Host "Please check the name, and re-run this script"
#>

}

<# No match for TO user - Alert and exit script #>

function noToMatch {
    # If no match for To user
    # Show a message box
    [System.Windows.Forms.MessageBox]::Show("There is no match for the TO user you entered. Please check the name, and re-run this script", "PROBLEM: NO MATCH FOR TO USER", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    findToUser
}

<# Menu: Users found, display information and provide menu options #>

function performActions {

    Write-Host ""
    Write-Host "From and to users both found, details as follows" -Foregroundcolor White -BackgroundColor Green
    Write-Host "" 
    Write-Host "From user:" $userNameFrom" ("$SamAccountNameFrom", "$orgUnitLocationFrom")"
    Write-Host "To user:" $userNameTo" ("$SamAccountNameTo", "$orgUnitLocationTo")"
    Write-host ""
    Write-host "Proceeding with script..."
    Write-host "

What do you want to do?
----------------------------

1. Copy Roles
   (With option to clear existing roles on TO account)

2. Re-enable account and copy roles
   (i.e. Departed employee; including Place in correct OU, then Copy Roles, with options to reset password, and clear existing roles on TO account)

3. Exit and end script
   (i.e. inactive user was provided, and will be reaching out for an Active user to clone)

"

    $number = Read-Host "Which task do you want to perform?"
    $output = @()
    switch ($number) {

        1 {
            Write-host ""
            # Go to function to make changes
            removeExistingMemberships

        }

        2 {
            Write-host ""
            # Go to function to make changes
            enableUserAccount
            
        }

        3 {
            Write-host ""
            # Go to function
            endScript
        }
        Default { Write-Host "No matches found , Enter Options 1, 2 or 3" -ForeGround "red" }

    }

}

<# Step 3: Apply changes to the TO account #>
function enableUserAccount {
    # Connect to Exchange Online if not already connected
    Write-Host ""
    Write-Host "Step 3: Enable TO user account if not already enabled" -ForegroundColor White -BackgroundColor Blue
    Write-Host ""

    $user = Get-ADUser -Identity $SamAccountNameTo -Properties Enabled

    if ($user.Enabled) {
        Write-Host "The account for $SamAccountNameTo is already enabled."
        # Go to next function
        moveOU
    }
    else {
        # Enable the AD account
        Enable-ADAccount -Identity $SamAccountNameTo
        Write-Host "The AD account for $SamAccountNameTo is now enabled."
    
        # Enable Mailbox
        $routingaddress = "$logonname@ucwest.mail.onmicrosoft.com"
    
        Write-Output "Enabling online mailbox for $logonname..."
        try {
            Enable-RemoteMailbox -Identity $SamAccountNameTo -RemoteRoutingAddress $routingaddress
            Write-Output "Cloud mailbox successfully enabled with routing address: $routingaddress"
        }
        catch {
            Write-Output "Failed to enable cloud mailbox. Please check Exchange Online connectivity."
        }
    
        # Retrieve user attributes
        $userDetails = Get-ADUser -Identity $SamAccountNameTo -Properties EmployeeID, Manager, Title, Department
    
        # Check and prompt if missing
        if (-not $userDetails.EmployeeID) {
            $newEmpID = Read-Host "Employee ID is missing. Please enter Employee ID"
            Set-ADUser -Identity $SamAccountNameTo -EmployeeID $newEmpID
        }
    
        if (-not $userDetails.Manager) {
            $newManager = Read-Host "Manager is missing. Please enter Manager's username (SAMAccountName)"
            # Enhancement opportunity - check if entered manager name exists on search, identify SAMaccountname, and populate accordingly
            Set-ADUser -Identity $SamAccountNameTo -Manager $newManager
        }
    
        if (-not $userDetails.Title) {
            $newTitle = Read-Host "Position Title is missing. Please enter Position Title"
            Set-ADUser -Identity $SamAccountNameTo -Title $newTitle
        }
    
        if (-not $userDetails.Department) {
            $newDepartment = Read-Host "Department is missing. Please enter Department/Area"
            Set-ADUser -Identity $SamAccountNameTo -Department $newDepartment
        }
    
        # Go to next function
        moveOU
    }
}


function moveOU {

    <#
2) Move the TO user into the active user OU
2a) Find which OU the user is currently in
#>

    $currentOU = $orgUnitLocationTo
    $targetOU = "OU=User Accounts,OU=Accounts,OU=ADOUName,DC=ADDCName,DC=local"
    $userName = Get-ADUser -Identity $SamAccountNameTo | Select-Object -ExpandProperty Name
    $userTargetOU = "CN=$userName,$targetOU"

    Write-Host ""
    Write-Host "Step 4: Move TO user into correct OU if not already there" -Foregroundcolor White -BackgroundColor Blue
    Write-Host ""
    Write-Host "Current TO OU:"
    Write-Host $currentOu
    Write-Host ""
    Write-Host "Intended TO Target OU:"
    Write-Host $userTargetOU

    if ($user) {
        # Test that the TO user has a current OU
        if ($currentOU) {
            # Go to next function
            compareAndMoveOU
        }
        else {
            # Error somewhere in the script, alert and exit
            Write-Host ""
            Write-Host "TO USER NOT IN AN OU" -Foregroundcolor White -BackgroundColor Red
            Write-Host "The user does not appear to be in an OU."
            Write-Host "Please rectify in AD after completing this script"
            Write-Host ""
            Write-Host "Press any key to continue..."
            [void][System.Console]::ReadKey($true)
            # Go to next function
            changePassword
        }
    }
    else {
        # Error somewhere in the script, alert and exit
        # This should be trapped earlier, but trapping again for safety
        Write-Host "User $SamAccountNameTo not found."
        #Break
        # Go to next function
        changePassword
    }
}

function compareAndMoveOU {

    <# 2b) Compare the current OU to the target OU, and move the account if not a match between current and target #>

    if ($currentOU -eq $userTargetOU) {
        # In the right OU, skip ahead
        Write-Host ""
        Write-Host "User is is already in the correct OU, skipping ahead"
        # Go to next function
        changePassword
    }
    elseif ($currentOU -ne $userTargetOU) {
        # In the wrong OU, let's move it
        Write-Host ""
        Write-Host $SamAccountNameTo "is in wrong OU, and will now be moved"
        # Get details and write to variables
        Write-Host ""
        Move-ADObject -Identity $orgUnitLocationTo -TargetPath $targetOU
        Write-Host "User $SamAccountNameTo moved to $targetOU successfully."
        # Go to next function
        changePassword
    }
}

function changePassword {

    <#
3) Change the password to default password
3a) Safety prompt for this? i.e. yes/no
3b) Force password change at next login if prompt activated
#>

    Write-Host ""
    Write-Host "Step 5: Offer option to change password to company default, and force change at next login" -Foregroundcolor White -BackgroundColor Blue
    Write-Host ""

    # Ask the user a yes/no question
    # Loop until a valid answer is given
    do {
        $promptChangePassword = Read-Host "Do you want to change the user's password to the company default, and force a password change at next login? (yes/no/y/n)"
        if ($promptChangePassword -eq "yes" -or $promptChangePassword -eq "y") {
            # Place the action to perform if the user answers 'yes'
            Write-Host ""
            Write-Host "User password WILL be changed to default, and flag set to require password change at next login"
            # Place the action to perform if the user answers 'yes'
            Set-ADAccountPassword -Identity $SamAccountNameTo -NewPassword (ConvertTo-SecureString -AsPlainText "Password#1234" -Force)
            Set-ADUser -Identity $SamAccountNameTo -ChangePasswordAtLogon $true
            # Move to next function
            removeExistingMemberships
        }
        elseif ($promptChangePassword -eq "no" -or $promptChangePassword -eq "n") {
            # Place the action to perform if the user answers 'no'
            "User password WILL NOT be changed"
            # Move to next function
            removeExistingMemberships
        }
        else {
            Write-Output "Please enter a valid response (yes or no)."
        }
    } until ($promptChangePassword -eq "yes" -or $promptChangePassword -eq "no" -or $promptChangePassword -eq "n" -or $promptChangePassword -eq "y")
    break
    removeExistingMemberships
}

function removeExistingMemberships {

    Write-Host ""
    Write-Host "Step 6: Offer option to remove all existing memberships from TO account" -Foregroundcolor White -BackgroundColor Blue
    Write-Host ""

    # Ask the user a yes/no question
    # Loop until a valid answer is given
    do {
        $promptRemoveExistingMemberships = Read-Host "Do you want to remove all existing group memberships from the TO user's account? (yes/no/y/n)"
        if ($promptRemoveExistingMemberships -eq "yes" -or $promptRemoveExistingMemberships -eq "y") {
            Write-Host ""
            Write-Host "Existing memberships WILL be removed"
            Write-Host ""
            # Place the action to perform if the user answers 'yes'
            $groups = Get-ADPrincipalGroupMembership -Identity $SamAccountNameTo
            foreach ($group in $groups) {
                if ($group.Name -ne "Domain Users") {
                    Remove-ADGroupMember -Identity $group -Members $SamAccountNameTo -Confirm:$false
                    Write-Output "Removed $SamAccountNameTo from group $($group.Name)"
                }                        
            }
            Write-Host ""
            Write-Host "Press any key to continue..."
            #[void][System.Console]::Read()
            # Go to next function
            CopyMemberships
        }
        elseif ($promptRemoveExistingMemberships -eq "no" -or $promptRemoveExistingMemberships -eq "n") {
            Write-Host ""
            Write-Output "Existing memberships WILL NOT be removed"
            # Place the action to perform if the user answers 'no'
            # Go to next function
            CopyMemberships
        }
        else {
            Write-Output "Please enter a valid response (yes or no)."
        }
    } until ($promptRemoveExistingMemberships -eq "yes" -or $promptRemoveExistingMemberships -eq "no" -or $promptChangePassword -eq "n" -or $promptChangePassword -eq "y")
}


function CopyMemberships {

    <#
Get memberships of FROM account, and apply to the TO account
#>
    Write-Host ""
    Write-Host "Step 7: Copy memberships for TO account" -Foregroundcolor White -BackgroundColor Blue
    Write-Host ""
    get-ADuser -identity $SamAccountNameFrom -properties memberof | select-object memberof -expandproperty memberof | Add-AdGroupMember -Members $SamAccountNameTo
    Write-Host "Group memberships from" $SamAccountNameFrom "have been copied to" $SamAccountNameTo
    Write-Host ""
    configureDefaultGroups

}

function configureDefaultGroups {

    <#
Step 8: Check and add user to required default groups, and remove from others present that aren't needed
#>

    Write-Host ""
    Write-Host "Step 8: Ensure correct baseline group membership for user" -ForegroundColor White -BackgroundColor Blue
    #
    # Validate user is a member of all required base groups, and remove all non-default signatures
    #

    Write-Host ""
    Write-Host "Checking base group memberships for $SamAccountNameTo"

    # Define user and list of groups to add the user to
    $userSamAccountName = $SamAccountNameTo
    $groupNamesAdd = @("All Staff")

    # Get the user object
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


    # Test user for member of non-Standard signature groups, and remove if present

    Write-Host ""
    Write-Host "Removing non-default signature"

    # Define the user and the list of groups to remove them from
    $userSamAccountName = $SamAccountNameTo
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

    # Go to next function
    showComplete

}

<# Step 9: Advise script and steps #>
function showComplete {
    # Set Variable Complete
    $scriptComplete = "1"
    Write-Host ""
    Write-Host "Script Complete" -Foregroundcolor White -BackgroundColor Green
    Write-Host ""
    Write-Host "Please stop to review the script output if needed"
    Write-Host ""
    Write-Host "Press any key to proceed..."
    Read-Host
    # Show a message box
    # [System.Windows.Forms.MessageBox]::Show("Requested actions performed", "Script Complete")
    endScript

}

function endScript {
    exit 1
}

# Execute first function, which runs script and calls all later functions
# findFromUser

<# Write header for script #>
Write-Host ""
Write-Host "COPY USER SCRIPT" -Foregroundcolor White -BackgroundColor Blue
Write-Host ""
Write-Host "Target OU for users:"$targetOu
Write-Host ""
Write-Host "Step 0: Connecting to On Prem AD" -Foregroundcolor White -BackgroundColor Blue
Write-Host ""

<# Step 0: Connect to On Prem AD instance #>

#Try to get a mailbox
try {
    $tst = get-mailbox firstname.lastname@domainname.com.au -ErrorAction silentlycontinue
}

#If unable to import exchange session
catch {
    $s = new-pssession -ConfigurationName microsoft.exchange -connectionuri http://HQEXMGMT01.ADDomainName.local/powershell
    Import-PSSession $S -allowclobber

}

<# Start the script #>

if ($scriptComplete -eq "1") {
    #Script Complete
    exit 1
}
else {
    # Call first function
    findFromUser
}
