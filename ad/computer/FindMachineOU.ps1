<#
Find machine in AD

Created: 4/11/2024
Created by: Michael HARRIS
Updated: X
Updated by: X

Purpose:
- Ask user for DNS name of machine
- Return information about that DNS name

Useful when unable to rename a machine due to naming conflict
#>

<# Core and initial variables #>

$scriptComplete = "0"

<# Load Core Assemblies #>

# Load the Windows Forms assembly
Add-Type -AssemblyName System.Windows.Forms

<# Step 1: Ask for machine name #>

function askMachine {

    Write-Host ""
    $machineName = Read-Host "Enter the machine name (whole or part of) being searched for`n(i.e. LAP524, WKS059)"
    Write-Host ""

    searchMachine

}

<# Step 2: Get and render machine information #>

function searchMachine {

    Write-Host "You entered" $machineName
    Write-Host ""
    Write-Host "Results (if any) are as follows:"
    Get-ADComputer -Filter "DNSHostName -like '$machineName*'" -Properties DisplayName, DistinguishedName | Format-Table DisplayName, Name, DistinguishedName -AutoSize
    showComplete

}

<# Step 3: Advise script complete #>

function showComplete {

    # Set Variable Complete
    $scriptComplete = "1"
    # Show message to end script
    Write-Host ""
    Write-Host "Press the Enter key to finish and exit the script..."
    Read-Host
    <# End Script #>

}

# Execute first function, which runs script and calls all later functions

<# Write header for script #>

Write-Host ""
Write-Host "LOCATE MACHINE" -Foregroundcolor White -BackgroundColor Blue
Write-Host ""

<# Step 0: Connect to On Prem AD instance #>

#Try to get a mailbox
try {

    $tst = get-mailbox firstname.lastname@domainname.com.au -ErrorAction silentlycontinue

}

#If unable to import exchange session
catch {

    $s = new-pssession -ConfigurationName microsoft.exchange -connectionuri http://ADSERVER.ADDomainName.local/powershell

    Import-PSSession $S -allowclobber

}


# Call first function to run script if not in a complete state

if ($scriptComplete -eq "1") {
    #Script Complete
    $scriptComplete = "0"
}
else {
    # Call first function
    askMachine
}
