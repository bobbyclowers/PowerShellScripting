<#
User Management External

Created:7/05/2025
Created by: Michael HARRIS

Updated: -
Updated by: -

Purpose:
- Gather details on all active Guest accounts
- Write details to CSV

Update history:

7/05/2025
- Initial release.

To-Do:
- Check for accounts within 14, and 7 days of expiry
- Notify sponsors by email of expiry and requirement to extend
- Disable account if expiry has now passed
#>
<#

Purpose:
Assist in the auditing of all guest accounts from CSV output, for future activities.


Notes:
1) employeeHireDate field will serve the purpose of a GuestDisabledDate, due to the lack of a feature in Entra AD to expire guest accounts automatically.

On creation of account, creator will be required to specify a hire date no later than 90 days from the current date.
14 days before this date, notification will be triggered to the sponsor.
If the sponsor replies to the ticket, advising of an extension, then the employeeHireDate should be extended for an additional 90 days.

If the sponsor does not reply, or an extension isn't processed on time, a further script will disable the account at midnight on the specified date.

If plausible, this would be configured in the automation account within PowerAutomate, but if not possible, run as a powershell script.

ChatGPT question used as part of construction (results partially incorrect due to not understanding the -Property flag required to get variable results with Microsoft.Graph queries:
Using powershell, how do I:

1. run Get-MgUser to get the to get the properties of Id, DisplayName, UserPrincipalName, Mail, createdDateTime, externalUserState, employeeHireDate for all external users, then
2. run Get-MgUserSponsor against the UserID from the previous step, to get the properties of X for the sponsor, then
3. run Get-MgUser again, to get the Display name of the UserID returned from Get-MgUserSponsor, then
4. Display all the results from the multiple Get-MgUser commands before writing it to a CSV, then
5. For any external user in the first step, where the property employeeHireDate is within the next 14 days, send an email - using SSL to Microsoft Exchange server with a specified user account, user account password, and mail server - to the sponsor advising them of the need to request extension of the account, then
6. For any external user in the first step, where the property employeeHireDate has already passed, disabled the account.
#>

# Install module if not installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}

# Script header

Write-Host ""
Write-Host "AUDIT EXTERNAL USER SCRIPT" -Foregroundcolor White -BackgroundColor Blue
Write-Host ""
Write-Host "Step 1: Connect to Entra AD via Microsoft Graph" -Foregroundcolor White -BackgroundColor Blue
Write-Host "When prompted, authenticate with the details of your IT Basic account"
Write-Host ""

# Connect to Microsoft Graph
# Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All", "EntitlementManagement.Read.All", "Mail.Send" -NoWelcome
Connect-MgGraph -ClientID 105184c7-ec2e-4202-93ec-068732391744 -TenantID 93a0a622-8ba7-41b3-a17f-1fbcc96fbff9 -Scopes "User.ReadWrite.All", "Directory.Read.All", "User.Invite.All", "EntitlementManagement.Read.All", "Mail.Send", "Mail.Send.Shared" -NoWelcome

Write-Host "Step 2: Locating active guest accounts, and sponsor details" -Foregroundcolor White -BackgroundColor Blue
Write-Host ""

# Prepare results collection
$results = @()

# Step 1: Get all enabled guest users (external accounts)
$guestUsers = Get-MgUser -Filter "userType eq 'Guest' and accountEnabled eq true" -All -Property Id, DisplayName, UserPrincipalName, Mail, CreatedDateTime, ExternalUserState, EmployeeHireDate, CompanyName | Select-Object Id, DisplayName, UserPrincipalName, Mail, CreatedDateTime, ExternalUserState, EmployeeHireDate, CompanyName | Sort-Object ExternalUserState, CompanyName, DisplayName

# Exchange Server SMTP configuration
<#
$smtpServer = "smtp.yourdomain.com"
$smtpPort = 587  # Use 587 for TLS/SSL email transmission
$smtpUser = "automation@domainname.com.au"
$smtpPassword = "your-email-password"
$emailSenderAddress = "it.accounts@domainname.com.au"
$emailCCAddress = "it.helpdesk@domainname.com.au"
$emailReplyToAddress = "it.helpdesk@domainname.com.au"
$emailPriority = "High"
#>

# Step 2–3: Loop through each user to find their sponsor and display name
foreach ($guest in $guestUsers) {
    try {
        # Get the sponsor(s) for the guest user
        $sponsors = Get-MgUserSponsor -UserId $guest.Id

        foreach ($sponsor in $sponsors) {
            # Step 3: Get sponsor details (DisplayName)
            $sponsorDetails = Get-MgUser -UserId $sponsor.Id

            # Step 4: Collect results
            $results += [PSCustomObject]@{
                GuestDisplayName    = $guest.DisplayName
                GuestCompanyName    = $guest.CompanyName
                GuestUPN            = $guest.UserPrincipalName
                GuestUserId         = $guest.Id
                GuestMail           = $guest.Mail
                GuestCreatedDate    = $guest.CreatedDateTime
                GuestExternalState  = $guest.ExternalUserState
                GuestDisableDate    = $guest.EmployeeHireDate
                SponsorDisplayName  = $sponsorDetails.DisplayName
                SponsorUPN          = $sponsorDetails.UserPrincipalName
                SponsorUserId       = $sponsorDetails.Id
            }

            # Step 5: Check the employeeHireDate and take actions accordingly
            <#
            $hireDate = [datetime]::Parse($guest.EmployeeHireDate)
            #>

            # If the hire date is within the next 14 days, send an email to the sponsor
            # step currently disabled pending mail configuration
            <# if ($hireDate -lt (Get-Date).AddDays(14) -and $hireDate -gt (Get-Date)) {
                # Send email to sponsor requesting account extension
                $emailBody = "Dear $($sponsorDetails.DisplayName),`n`nAs the sponsor for the external account of $($guest.DisplayName) from $($guest.companyName), you are being advised this account will expire in the next 14 days.`n`nIf this person has a continued need for access, you must reply to this email to request a further 90 day extension.`n`nFailure to reply to or action this email will result in the account being automatically disabled after $($guest.EmployeeHireDate).`n`nYour prompt action on this notification is greatly appreciated.`n`nKind regards,"
               
                # Send the email using SMTP
                # Look at using Power Automate to capture this Outbound email once sent, and post a notification to Teams about same; including a relevant teams message to the sponsoring manager
                Send-MailMessage -From $emailSenderAddress `
                                 -To $sponsorDetails.UserPrincipalName `
                                 -ReplyTo $emailReplyToAddress `
                                 -CC $emailCCAddress `
                                 -Subject "Action required: Request renewal of external account for $($guest.DisplayName) from $($guest.companyName) no later than $($guest.EmployeeHireDate)" `
                                 -Priority $emailPriority `
                                 -Body $emailBody `
                                 -SmtpServer $smtpServer `
                                 -Port $smtpPort `
                                 -UseSsl $true `
                                 -Credential (New-Object System.Management.Automation.PSCredential($smtpUser, (ConvertTo-SecureString $smtpPassword -AsPlainText -Force))) `
                                 -DeliveryNotificationOption OnFailure
            }
            #>

            # Consider adding an additional final reminder email here when the hire date is less than 7 days
            # Add if statement and email code for same here with appropriate adjustments, stealing from Step 5
            # Look at using Power Automate to capture this Outbound email once sent, and post a notification to Teams about same; including a relevant teams message to the sponsoring manager

            # Step 6: Disable the account if the hire date has already passed
            <# if ($hireDate -lt (Get-Date)) {
                # Disable the account
                Set-MgUser -UserId $guest.Id -AccountEnabled $false
                # Notify the sponsor again
                # Add email code for same here with appropriate adjustments, stealing from Step 5
                # Look at using Power Automate to capture this Outbound email once sent, and post a notification to Teams about same.
            }
            #>
        }
    }
    catch {
        Write-Warning "Error processing guest user '$($guest.DisplayName)': $_"
    }
}

Write-Host "Step 3: Results" -Foregroundcolor White -BackgroundColor Blue
Write-Host ""

# Step 7: Display results to screen
$results | Format-Table -AutoSize

Write-Host "Step 4: Write to CSV" -Foregroundcolor White -BackgroundColor Blue

# Note: Add in date for file, for record keeping and analysis purposes
# Step 8: Export results to CSV
$results | Export-Csv -Path "V:\Guest-Account-Audit\GuestUsers_Active.csv" -NoTypeInformation
Write-Host "File can be found at V:\Guest-Account-Audit\GuestUsers_Active.csv"

Write-Host ""
Write-Host "Step 5: Disconnect from Microsoft Graph" -Foregroundcolor White -BackgroundColor Blue
# Disconnect from session
Disconnect-MgGraph
Write-Host "Disconnected"