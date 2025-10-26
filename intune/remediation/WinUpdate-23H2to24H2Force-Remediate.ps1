<#
.SYNOPSIS
Installs Windows 11 24H2 enablement package for targeted devices and schedules reboot with user notifications.

#>

# Bypass execution policy for this session
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

Start-Transcript -Path "$env:ProgramData\Intune-Win11-24H2-Enablement.log" -Append

# --- Begin targeting logic ---

$dayOfWeek = (Get-Date).DayOfWeek.value__  # Sunday=0 ... Saturday=6

if ($dayOfWeek -eq 0 -or $dayOfWeek -eq 6) {
    Write-Host "Weekend detected (Saturday or Sunday). Skipping remediation."
    exit 0
}

$allowedDigitsMap = @{
    1 = @(0,1)  # Monday
    2 = @(2,3)  # Tuesday
    3 = @(4,5)  # Wednesday
    4 = @(6,7)  # Thursday
    5 = @(8,9)  # Friday
}

$hostname = $env:COMPUTERNAME
$lastChar = $hostname.Substring($hostname.Length - 1, 1)

if ($lastChar -match '^[A-Za-z]$') {
    if ($dayOfWeek -ne 5) {
        Write-Host "Hostname ends with letter '$lastChar' and today is not Friday. Skipping remediation."
        exit 0
    }
    Write-Host "Hostname ends with letter '$lastChar' and today is Friday. Proceeding..."
}
elseif ($lastChar -match '^\d$') {
    $lastDigit = [int]$lastChar
    if (-not $allowedDigitsMap.ContainsKey($dayOfWeek)) {
        Write-Host "No allowed digits configured for day $dayOfWeek. Skipping remediation."
        exit 0
    }
    if (-not ($allowedDigitsMap[$dayOfWeek] -contains $lastDigit)) {
        Write-Host "Hostname ends with digit $lastDigit, which is not allowed for day $dayOfWeek. Skipping remediation."
        exit 0
    }
    Write-Host "Hostname ends with digit $lastDigit allowed for today (day $dayOfWeek). Proceeding..."
}
else {
    Write-Host "Hostname last char '$lastChar' is neither digit nor letter. Skipping remediation."
    exit 0
}

# --- End targeting logic ---

$WorkHoursStart = 9
$WorkHoursEnd = 17
$ToastScriptPath = "$env:ProgramData\Win11-24H2-RebootWarning.ps1"

function Create-RebootScheduledTask {
    param (
        [string]$scriptPath,
        [int]$workStart,
        [int]$workEnd
    )

    $taskName = "Win11-24H2-RebootWarning"

    $scriptContent = @"
Start-Sleep -Seconds 10

function Show-ToastNotification {
    param (
        [string]\$Title,
        [string]\$Message
    )
    \$toastXml = @"
<toast>
  <visual>
    <binding template='ToastGeneric'>
      <text>\$Title</text>
      <text>\$Message</text>
    </binding>
  </visual>
</toast>
"@

    \$xml = New-Object Windows.Data.Xml.Dom.XmlDocument
    \$xml.LoadXml(\$toastXml)

    \$toast = [Windows.UI.Notifications.ToastNotification]::new(\$xml)
    \$notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("Windows Update")
    \$notifier.Show(\$toast)
}

\$now = Get-Date
\$currentHour = \$now.Hour

if (\$currentHour -ge $workStart -and \$currentHour -lt $workEnd) {
    \$rebootTime = \$now.Date.AddHours($workEnd).AddMinutes(15)
} else {
    \$rebootTime = \$now.AddMinutes(5)
}

\$timeToReboot = \$rebootTime - \$now

if (\$timeToReboot.TotalMinutes -gt 15) {
    Start-Sleep -Seconds ([math]::Round((\$timeToReboot.TotalMinutes - 15) * 60))
}

\$toastMessage15 = @"
**IMPORTANT**  
Your device will restart in 15 minutes to complete a mandatory Windows update.  
Please save all work immediately. Restart cannot be postponed.
"@
Show-ToastNotification -Title "Restart Required - Windows Update" -Message \$toastMessage15

\$now = Get-Date
\$wait2min = (\$rebootTime - \$now).TotalMinutes - 2
if (\$wait2min -gt 0) {
    Start-Sleep -Seconds ([math]::Round(\$wait2min * 60))
}

\$toastMessage2 = @"
**FINAL WARNING**  
Your device will restart in 2 minutes to complete a mandatory Windows update.  
Please save all work now. Restart cannot be postponed.
"@
Show-ToastNotification -Title "Final Restart Warning - Windows Update" -Message \$toastMessage2

Restart-Computer -Force
"@

    $scriptContent | Out-File -FilePath $scriptPath -Encoding UTF8 -Force

    $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""
    $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1)
    $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -RunLevel Highest

    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Force
    Write-Host "Scheduled task '$taskName' created to handle reboot notifications and restart."
}

try {
    Write-Host "Starting remediation for Windows 11 24H2 enablement package..."

    $osInfo = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    $displayVersion = $osInfo.DisplayVersion
    $currentBuild = [int]$osInfo.CurrentBuild

    if ($displayVersion -ne "23H2" -or $currentBuild -ne 22631) {
        Write-Host "Device is not running Windows 11 23H2. No remediation required."
        exit 0
    }

    $arch = (Get-CimInstance Win32_OperatingSystem).OSArchitecture
    Write-Host "Detected architecture: $arch"

    $kbUrls = @{
        "64-bit" = "https://download.windowsupdate.com/d/msdownload/update/software/updt/2024/06/windows11.0-kb5039302-x64_abcdef1234567890.msu"
        "ARM64"  = "https://download.windowsupdate.com/d/msdownload/update/software/updt/2024/06/windows11.0-kb5039302-arm64_abcdef1234567890.msu"
    }

    $archKey = if ($arch -like "*64*") { "64-bit" } elseif ($arch -like "*ARM64*") { "ARM64" } else { "" }

    if ([string]::IsNullOrEmpty($archKey) -or -not $kbUrls.ContainsKey($archKey)) {
        Write-Error "Unsupported architecture detected: $arch"
        exit 2
    }

    $kbUrl = $kbUrls[$archKey]
    $tempFile = Join-Path $env:TEMP ("KB5039302-$archKey.msu")

    Write-Host "Downloading enablement package from $kbUrl ..."
    Invoke-WebRequest -Uri $kbUrl -OutFile $tempFile -UseBasicParsing -ErrorAction Stop
    Write-Host "Download complete: $tempFile"

    Write-Host "Installing enablement package silently..."
    $installProcess = Start-Process -FilePath "wusa.exe" -ArgumentList "`"$tempFile` /quiet /norestart" -Wait -PassThru

    if ($installProcess.ExitCode -eq 0 -or $installProcess.ExitCode -eq 3010) {
        Write-Host "Enablement package installed successfully."

        Create-RebootScheduledTask -scriptPath $ToastScriptPath -workStart $WorkHoursStart -workEnd $WorkHoursEnd

        exit 0
    }
    else {
        Write-Error "Enablement package installation failed with exit code $($installProcess.ExitCode). Skipping reboot."
        exit 2
    }
}
catch {
    Write-Error "Remediation failed with error: $_"
    exit 2
}
finally {
    Stop-Transcript
}
