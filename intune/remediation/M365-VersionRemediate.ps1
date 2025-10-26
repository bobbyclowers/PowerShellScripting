# -------------------------------------------------------------
# Remediation Script - M365 Monthly Enterprise Channel Update (User-Safe)
# -------------------------------------------------------------
# Ensures Office C2R update runs non-disruptively, with user warning if Office apps are open.
# References:
# - https://learn.microsoft.com/en-us/deployoffice/update-options-for-office365-proplus
# - https://learn.microsoft.com/en-us/officeupdates/monthly-enterprise-channel
# - https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/expand
# - https://learn.microsoft.com/en-us/deployoffice/office-deployment-tool-configuration-options#clientversiontoreport

$logFile = Join-Path $env:TEMP "M365RemediateLog-$(Get-Date -Format yyyyMMdd-HHmmss).log"
$fallbackMinimumBuildVersion = [version]"18730.20220"
$maxReminders = 2
$reminderIntervalSeconds = 120  # Adjust this value as needed (in seconds)

function Write-Log {
    param([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $full = "[$timestamp] $message"
    Add-Content -Path $logFile -Value $full
    Write-Host $full
}

function Get-InstalledBuildVersionOnly {
    try {
        $reg = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
        $version = (Get-ItemProperty -Path $reg -ErrorAction Stop).ClientVersionToReport
        Write-Log "Installed Office version: $version"
        $parts = $version -split '\.'
        if ($parts.Count -ge 4) {
            return [version]"$($parts[2]).$($parts[3])"
        }
    }
    catch { Write-CustomMessage "WARNING: Ignored exception in version remediate step: $_" }
    return $null
}

function Get-LatestMonthlyEnterpriseBuildVersion {
    try {
        $url = 'https://officecdn.microsoft.com/pr/wsus/ofl.cab'
        $cabPath = Join-Path $env:TEMP 'ofl.cab'
        $extractedXml = Join-Path $env:TEMP 'ofl.xml'
        Invoke-WebRequest -Uri $url -OutFile $cabPath -UseBasicParsing
        expand $cabPath -F:* $env:TEMP | Out-Null

        if (Test-Path $extractedXml) {
            [xml]$xml = Get-Content -Path $extractedXml
            $latest = $xml.OfficeClientEdition.DownloadUrls.DownloadUrl |
            Where-Object { $_.Branch -eq 'MonthlyEnterprise' } |
            Sort-Object Revision -Descending |
            Select-Object -First 1
            if ($latest.Revision) {
                return [version]$latest.Revision
            }
        }
    }
    catch { Write-CustomMessage "WARNING: Ignored exception in version remediate step: $_" }
    return $fallbackMinimumBuildVersion
}

function Detect-RunningOfficeApps {
    $apps = @("WINWORD", "EXCEL", "OUTLOOK", "POWERPNT", "ONENOTE", "MSACCESS", "VISIO", "MSPUB", "LYNC")
    return Get-Process | Where-Object { $apps -contains $_.Name } | Select-Object -Unique
}

function Prompt-UserToCloseOffice {
    $running = Detect-RunningOfficeApps
    if (-not $running) { return $true }

    Write-Log "Prompting user to close Office apps: $($running.Name -join ', ')"
    foreach ($proc in $running) {
        try {
            if ($proc.MainWindowHandle -ne 0) {
                $proc.CloseMainWindow() | Out-Null
            }
        }
        catch { Write-CustomMessage "WARNING: Ignored exception in version remediate nested step: $_" }
    }

    Start-Sleep -Seconds 10
    return (-not (Detect-RunningOfficeApps))
}

function Show-HighPriorityToast {
    param (
        [string]$Title,
        [string]$Body,
        [int]$RemindIntervalSec
    )

    $escapedTitle = [System.Security.SecurityElement]::Escape($Title)
    $escapedBody = [System.Security.SecurityElement]::Escape($Body)

    $remindMinutes = [math]::Round($RemindIntervalSec / 60)
    $remindLabel = "Remind Me in $remindMinutes Minute" + ($(if ($remindMinutes -ne 1) { "s" } else { "" }))

    $template = @"
<toast scenario="reminder" launch="m365update">
  <visual>
    <binding template='ToastGeneric'>
      <text>$escapedTitle</text>
      <text>$escapedBody</text>
    </binding>
  </visual>
  <actions>
    <action content="Save & Close Now" arguments="saveclose" activationType="foreground"/>
    <action content="$remindLabel" arguments="remindlater" activationType="background"/>
  </actions>
</toast>
"@

    try {
        [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
        $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
        $xml.LoadXml($template)
        $toast = [Windows.UI.Notifications.ToastNotification]::new($xml)
        $notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("M365Update")
        $notifier.Show($toast)
        Write-Log "Toast shown: $Title with delay of $remindMinutes minutes"
    }
    catch {
        Write-Log "Toast failed: $_"
    }
}

function Start-OfficeUpdate {
    param([version]$TargetVersion)

    $c2r = Get-ChildItem -Path "${env:ProgramFiles(x86)}\Microsoft Office\root" -Recurse -Filter OfficeC2RClient.exe -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $c2r -or -not $TargetVersion) {
        Write-Log "OfficeC2RClient not found or no target version."
        return $false
    }

    $fullVersion = "16.0.$($TargetVersion.ToString())"
    Write-Log "Running update to $fullVersion with forceappshutdown=false"
    & $c2r.FullName /update user updatetoversion=$fullVersion forceappshutdown=false displaylevel=true
    return $true
}

# -------------------------------
# Main Execution
# -------------------------------
Write-Log "--- M365 Remediation Start ---"

$installed = Get-InstalledBuildVersionOnly
$latest = Get-LatestMonthlyEnterpriseBuildVersion
$needsUpdate = (-not $installed -or $installed -lt $latest)

if ($needsUpdate) {
    Write-Log "Update required: Installed=$installed, Latest=$latest"

    $running = Detect-RunningOfficeApps
    $reminderCount = 0

    while ($running -and $reminderCount -lt $maxReminders) {
        Show-HighPriorityToast -Title "Microsoft 365 Update Needed" `
            -Body "Please save your work in, and close, $($running.Name -join ', ') to avoid interruptions." `
            -RemindIntervalSec $reminderIntervalSeconds

        Write-Log "Waiting $reminderIntervalSeconds seconds before re-check (Reminder #$($reminderCount + 1))"
        Start-Sleep -Seconds $reminderIntervalSeconds

        if (Prompt-UserToCloseOffice) {
            Write-Log "User closed Office apps. Proceeding with update."
            break
        }

        $reminderCount++
        $running = Detect-RunningOfficeApps
    }

    if ($reminderCount -ge $maxReminders) {
        Write-Log "Max reminders ($maxReminders) reached. Proceeding with update regardless."
    }

    if (Start-OfficeUpdate -TargetVersion $latest) {
        Write-Log "Update initiated successfully."
    }
    else {
        Write-Log "Failed to initiate update."
    }
}
else {
    Write-Log "No update required."
}

Write-Log "--- M365 Remediation Complete ---"
exit 0
