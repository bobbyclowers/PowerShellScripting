<#
.SYNOPSIS
Detect if device is running Windows 11 23H2 and matches day/hostname targeting rules.

.EXIT CODES
0 = Compliant (no remediation needed)
1 = Not compliant (remediation required)
2 = Error or unsupported architecture/version
#>

# Bypass execution policy for this session
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

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
    $osInfo = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    $displayVersion = $osInfo.DisplayVersion
    $currentBuild = [int]$osInfo.CurrentBuild

    if ($displayVersion -eq "24H2") {
        Write-Host "Windows 11 24H2 detected - compliant."
        exit 0
    }
    elseif ($displayVersion -eq "23H2") {
        Write-Host "Windows 11 23H2 detected - remediation needed."
        exit 1
    }
    else {
        Write-Host "Unsupported Windows version: $displayVersion. Skipping remediation."
        exit 0
    }
}
catch {
    Write-Host "Error detecting OS version: $_"
    exit 2
}
