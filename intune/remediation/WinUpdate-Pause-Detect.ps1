# Detection Script - Detect if Windows Update is paused
# PowerShell 5.1 and 7 Compatible

$log = @()
$paused = $false
$reason = @()

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
    Write-Host "Hostname ends with letter '$lastChar' and today is Friday. Checking compliance..."
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
    Write-Host "Hostname ends with digit $lastDigit allowed for today (day $dayOfWeek). Checking compliance..."
}
else {
    Write-Host "Hostname last char '$lastChar' is neither digit nor letter. Skipping remediation."
    exit 0
}

# --- End targeting logic ---

try {
    $wuKey = "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings"
    $policyKey = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
    $pauseUpdatesValue = Get-ItemProperty -Path $wuKey -Name PauseUpdates -ErrorAction SilentlyContinue
    $pauseStartTime = Get-ItemProperty -Path $wuKey -Name PauseStartTime -ErrorAction SilentlyContinue
    $policyPause = Get-ItemProperty -Path $policyKey -Name SetDisablePauseUXAccess -ErrorAction SilentlyContinue

    if ($pauseUpdatesValue.PauseUpdates -eq 1) {
        $paused = $true
        $reason += "PauseUpdates=1 in UX\Settings"
    }

    if ($pauseStartTime.PauseStartTime) {
        $paused = $true
        $reason += "PauseStartTime exists"
    }

    if ($policyPause.SetDisablePauseUXAccess -eq 0) {
        $paused = $true
        $reason += "Policy disables pause UX access"
    }

    if ($paused) {
        $log += "Windows Update appears to be paused."
        $log += "Reasons: $($reason -join '; ')"
        $log | ForEach-Object { Write-Output $_ }
        exit 1
    } else {
        Write-Output "Windows Update is not paused."
        exit 0
    }
}
catch {
    Write-Output "Detection script failed: $_"
    exit 1
}