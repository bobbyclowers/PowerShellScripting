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

$CurrentWin10 = [Version]"10.0.19045"
$CurrentWin11 = [Version]"10.0.22631"

$GetOS = Get-ComputerInfo -property OsVersion
$OSversion = [Version]$GetOS.OsVersion

if  ($OSversion -match [Version]"10.0.1")
    {
    if  ($OSversion -lt $CurrentWin10)
        {
        Write-Output "OS version currently on $OSversion"
        exit 1
        }
    }

if  ($OSversion -match [Version]"10.0.2")
    {
    if  ($OSversion -lt $CurrentWin11)
        {
        Write-Output "OS version currently on $OSversion"
        exit 1
        }
    }

do  {
    try {
        $lastupdate = Get-HotFix | Sort-Object -Property InstalledOn | Select-Object -Last 1 -ExpandProperty InstalledOn
        $Date = Get-Date

        $diff = New-TimeSpan -Start $lastupdate -end $Date
        $days = $diff.Days
        }
    catch   {
            Write-Output "Attempting WMI repair"
            Start-Process "C:\Windows\System32\wbem\WMIADAP.exe" -ArgumentList "/f"
            Start-Sleep -Seconds 120
            }
    }
    until ($null -ne $days)

$Date = Get-Date

$diff = New-TimeSpan -Start $lastupdate -end $Date
$days = $diff.Days

if  ($days -ge 40 -or $null -eq $days)
    {
    Write-Output "Troubleshooting Updates - Last update was $days days ago"
    exit 1
    }
else{
    Write-Output "Windows Updates ran $days days ago"
    exit 0
    }