<#
Convert E365 User Mailbox to Shared Mailbox

Created: 25/03/2025
Created by: Michael HARRIS

Updated: -
Updated by: -

Purpose:
- Confirm user exists
- Update mailbox to shared mailbox

Update history:

10/03/2025
- Initial release.
#>

<# Core and initial variables #>

$scriptComplete = "0"

<# Key functions #>

<# Connect to E365 #>

function showHeader {

    Write-Host ""
    Write-Host "E365 - Convert user mailbox to shared mailbox" -Foregroundcolor White -BackgroundColor Blue
    Write-Host ""
    Write-Host "Step 0: Connecting to On Prem AD/Exchange environment" -Foregroundcolor White -BackgroundColor Blue

    <# Step 0: Connect to On Prem AD/Exchange instance #>

    findFromUser

}

<# Find the user #>

function findFromUser {
    # Get name input for user to copy FROM
    Write-Host ""
    Write-Host "Step 1: Which user (Person to convert to shared mailbox)" -Foregroundcolor White -BackgroundColor Blue
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
        connectExchange
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
        connectExchange
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
        connectExchange
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
        connectExchange
    }


}

<# No match for FROM user - Alert and return to findFromUser #>

function noFromMatch {
    # Show a message box
    [System.Windows.Forms.MessageBox]::Show("There is no match for the FROM user you entered. Please check the name, and re-enter when prompted", "PROBLEM: NO MATCH FOR FROM USER", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    findFromUser
}

<# Connect to E365 #>

function connectExchange {

    <# Step 2a: Connect to On Prem AD/Exchange instance #>

    # Connect to Exchange Online if not already connected
    if (-not (Get-PSSession | Where-Object { $_.ComputerName -like "*exchange*" })) {
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://HQEXMGMT01.ADDomainName.local/powershell
        Import-PSSession $Session -AllowClobber
    }

    Set-RemoteMailbox -Identity $SamAccountNameFrom -HiddenFromAddressListsEnabled $true

    <# Step 2b: Connect to Cloud Exchange instance #>

    Write-Host ""
    Write-Host "Step 2: E365 Modern Auth" -Foregroundcolor White -BackgroundColor Blue
    Write-Host "Connect to E365 with your admin (X suffix) account to perform this action"

    Connect-ExchangeOnline

    convertToShared

}

<# Convert the box to a shared mailbox #>

function convertToShared {

    Write-Host ""
    Write-Host "Step 3: Convert the mailbox" -Foregroundcolor White -BackgroundColor Blue
    Write-Host ""

    # Set-Mailbox $SamAccountNameFrom -Type Shared
    Set-Mailbox -Identity $SamAccountNameFrom -Type Shared

    showComplete

}

<# Advise script and steps complete #>

function showComplete {

    # Set Variable Complete
    $scriptComplete = "1"
    # Show a message box
    [System.Windows.Forms.MessageBox]::Show("Requested actions performed", "Script Complete")
    endScript

}

function endScript {
    exit 1
}

<# Start the script #>

if ($scriptComplete -eq "1") {
    #Script Complete
    exit 1
}
else {
    # Call first function
    showHeader
}
