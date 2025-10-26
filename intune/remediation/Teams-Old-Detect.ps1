# Detection script for Microsoft Teams (Classic, New, Personal) – Intune
# Exit Codes:
# 0 = Compliant
# 1 = Non-compliant (remediation required)
# 2 = Off-schedule
# 3 = User in Teams call/meeting
# 4 = Device name ends in symbol — unable to determine target group

# Path to the log file (you can customize this)
$logFile = "$env:TEMP\TeamsDetection.log"

function Write-Log {
    param([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $full = "[$timestamp] $message"
    Add-Content -Path $logFile -Value $full
    Write-Host $full
}

$ExitCode = 0
$teamsFound = @()

function Add-TeamsResult {
    param (
        [string]$Path,
        [string]$Type
    )
    if (Test-Path $Path) {
        try { $version = (Get-Item $Path).VersionInfo.ProductVersion } catch { $version = "Unknown" }
        $teamsFound += [PSCustomObject]@{
            Type    = $Type
            Path    = $Path
            Version = $version
        }
        Write-Log "Detected: $Type"
        Write-Log "Path: $Path"
        Write-Log "Version: $version"
    }
}

# 0. Validate device naming pattern
$hostname = $env:COMPUTERNAME
if ($hostname[-1] -notmatch '[A-Za-z0-9]') {
    Write-Log "Device name ends in non-alphanumeric character. Manual remediation required."
    exit 4
}

# 1. Check remediation day
$today = (Get-Date).DayOfWeek
$lastChar = $hostname[-1]
$allowedDigitsMap = @{
    1 = @(0,1)  # Monday
    2 = @(2,3)  # Tuesday
    3 = @(4,5)  # Wednesday
    4 = @(6,7)  # Thursday
    5 = @(8,9)  # Friday
}

$dayNumber = [int](Get-Date).DayOfWeek
$runToday = $false

if ($lastChar -match '\d') {
    $digit = [int]$lastChar
    if ($allowedDigitsMap.ContainsKey($dayNumber)) {
        if ($allowedDigitsMap[$dayNumber] -contains $digit) {
            $runToday = $true
        }
    }
} elseif ($lastChar -match '[A-Za-z]' -and $today -eq 'Friday') {
    $runToday = $true
}

if (-not $runToday) {
    Write-Log "Not scheduled remediation day for this device."
    exit 2
}

# 2. Check if Teams is running or user is in a call
$runningTeams = Get-Process -Name Teams -ErrorAction SilentlyContinue
$inActiveCall = $false

if ($runningTeams) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder text, int count);
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
}
"@
    foreach ($proc in $runningTeams) {
        if ($proc.MainWindowHandle -ne 0) {
            $sb = New-Object System.Text.StringBuilder 256
            [void][Win32]::GetWindowText($proc.MainWindowHandle, $sb, $sb.Capacity)
            $title = $sb.ToString()
            if ($title -match '(call|meeting|sharing|audio|video)' -and [Win32]::IsWindowVisible($proc.MainWindowHandle)) {
                $inActiveCall = $true
                Write-Log "Teams call/meeting detected: '$title'"
                break
            }
        }
    }
}

if ($inActiveCall) {
    Write-Log "User appears to be in a Teams call or meeting. Skipping remediation."
    exit 3
}

# 3. Detect Classic Teams
Add-TeamsResult -Path "$env:LOCALAPPDATA\Microsoft\Teams\Update.exe" -Type "Classic Teams (Per-User)"
Add-TeamsResult -Path "C:\Program Files\Microsoft\Teams\current\Teams.exe" -Type "Classic Teams (Machine-Wide)"
Add-TeamsResult -Path "C:\Program Files (x86)\Teams Installer\Teams.exe" -Type "Classic Teams (Machine-Wide Installer)"

# 4. Detect New Teams per-user
$teamsFolder = "$env:LOCALAPPDATA\Microsoft\Teams"
if (Test-Path $teamsFolder) {
    Get-ChildItem -Path $teamsFolder -Filter Teams.exe -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
        Add-TeamsResult -Path $_.FullName -Type "New Teams (Per-User)"
    }
}

# 5. Detect AppX/MSIX (Org Teams or Personal Teams)
$appxTeams = Get-AppxPackage -AllUsers | Where-Object { $_.Name -like "*Teams*" }
foreach ($pkg in $appxTeams) {
    $type = if ($pkg.Name -like "MicrosoftCorporationII.MSTeams") {
        "Teams (Personal - AppX)"
    } elseif ($pkg.Name -like "MSTeams*" -or $pkg.Name -like "*MicrosoftTeams*") {
        "Teams (Org - AppX/MSIX)"
    } else {
        "Teams (Unknown AppX): $($pkg.Name)"
    }

    $teamsFound += [PSCustomObject]@{
        Type    = $type
        Path    = $pkg.InstallLocation
        Version = $pkg.Version
    }
    Write-Log "Detected: $type"
    Write-Log "Path: $($pkg.InstallLocation)"
    Write-Log "Version: $($pkg.Version)"
}

# 6. Determine compliance
$nonCompliant = $teamsFound | Where-Object {
    $_.Type -notlike "*AppX/MSIX*" -and $_.Type -notlike "*Org*"
}

if ($nonCompliant) {
    Write-Log "Non-compliant Teams versions detected. Remediation required."
    exit 1
} else {
    Write-Log "Compliant — only AppX/MSIX Org Teams detected."
    exit 0
}
