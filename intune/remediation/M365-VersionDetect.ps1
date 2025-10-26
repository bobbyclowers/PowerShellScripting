# -----------------------------
# Detection Script - M365 Apps
# Ensures device is on the latest Monthly Enterprise version
# Exit Codes:
# 0 = Compliant
# 1 = Non-compliant (requires remediation)
# 2 = Off-schedule (intentionally skipped)
# 4 = Unidentifiable target group (symbol in device name)
# -----------------------------
# References:
# - https://learn.microsoft.com/en-us/deployoffice/office-deployment-tool-configuration-options#channel
# - https://learn.microsoft.com/en-us/officeupdates/monthly-enterprise-channel
# - https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/expand
# - https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.management/get-itemproperty
# - https://learn.microsoft.com/en-us/powershell/scripting/learn/deep-dives/everything-about-exit-codes

$logFile = Join-Path $env:TEMP "M365DetectLog-$(Get-Date -Format yyyyMMdd-HHmmss).log"

# Schedule map: used for spread scheduling by device name last digit
$allowedDigitsMap = @{
    1 = @(0,1)  # Monday
    2 = @(2,3)  # Tuesday
    3 = @(4,5)  # Wednesday
    4 = @(6,7)  # Thursday
    5 = @(8,9)  # Friday
}

$fallbackMinimumBuildVersion = [version]"18730.20220"  # Published June 11, 2025

function Write-Log {
    param([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $full = "[$timestamp] $message"
    Add-Content -Path $logFile -Value $full
    Write-Host $full
}

function Should-RunToday {
    $today = [int](Get-Date).DayOfWeek
    if ($today -eq 0 -or $today -eq 6) {
        Write-Log "Weekend detected. Skipping execution."
        exit 2
    }

    $hostname = $env:COMPUTERNAME
    $lastChar = $hostname[-1]

    if ($lastChar -match '[a-zA-Z]') {
        $run = ($today -eq 5)  # Friday only for letter suffix
        Write-Log "Hostname ends in letter. Today is Friday: $run"
        if (-not $run) { exit 2 }
        return
    }

    if ($lastChar -match '\d') {
        $digit = [int]$lastChar
        $shouldRun = $allowedDigitsMap[$today] -contains $digit
        Write-Log "Today = $today; Hostname digit = $digit; Run: $shouldRun"
        if (-not $shouldRun) { exit 2 }
        return
    }

    Write-Log "Hostname ends in symbol. Cannot determine compliance group."
    exit 4
}

function Get-InstalledBuildVersionOnly {
    try {
        $reg = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
        $version = (Get-ItemProperty -Path $reg -ErrorAction Stop).ClientVersionToReport
        # Learn: https://learn.microsoft.com/en-us/deployoffice/office-deployment-tool-configuration-options#clientversiontoreport
        Write-Log "Full installed Office version: $version"

        $parts = $version -split '\.'
        if ($parts.Count -ge 4) {
            $buildVer = "$($parts[2]).$($parts[3])"
            Write-Log "Parsed build version: $buildVer"
            return [version]$buildVer
        } else {
            Write-Log "Unexpected version format: $version"
            return $null
        }
    } catch {
        Write-Log "Office not installed or registry unreadable."
        return $null
    }
}

function Get-LatestMonthlyEnterpriseBuildVersion {
    try {
        $url = 'https://officecdn.microsoft.com/pr/wsus/ofl.cab'
        $cabPath = Join-Path $env:TEMP 'ofl.cab'
        $extractedXml = Join-Path $env:TEMP 'ofl.xml'

        # Learn: https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/expand
        Invoke-WebRequest -Uri $url -OutFile $cabPath -UseBasicParsing
        expand $cabPath -F:* $env:TEMP | Out-Null

        if (Test-Path -Path $extractedXml) {
            [xml]$xml = Get-Content -Path $extractedXml
            $latest = $xml.OfficeClientEdition.DownloadUrls.DownloadUrl |
                Where-Object { $_.Branch -eq 'MonthlyEnterprise' } |
                Sort-Object Revision -Descending |
                Select-Object -First 1

            if ($latest -and $latest.Revision) {
                Write-Log "Latest Monthly Enterprise version from ofl.cab: $($latest.Revision)"
                return [version]$latest.Revision
            }
        } else {
            Write-Log "ofl.xml missing. Trying fallback."
        }
    } catch {
        Write-Log "Failed ofl.cab check: $_"
    }

    try {
        $jsonUrl = 'https://learn.microsoft.com/en-us/officeupdates/monthly-enterprise-channel'
        $json = Invoke-WebRequest -Uri $jsonUrl -UseBasicParsing
        if ($json.Content -match 'Build ([0-9]+\.[0-9]+)') {
            $build = $matches[1]
            Write-Log "Parsed build from HTML fallback: $build"
            return [version]$build
        } else {
            Write-Log "Regex fallback match failed on HTML."
        }
    } catch {
        Write-Log "JSON fallback error: $_"
    }

    Write-Log "Defaulting to fallback version: $fallbackMinimumBuildVersion"
    return $null
}

# ---------------------
# MAIN EXECUTION
# ---------------------
Write-Log "--- Detection Script Start ---"
Should-RunToday

$installed = Get-InstalledBuildVersionOnly
$latest = Get-LatestMonthlyEnterpriseBuildVersion

if (-not $installed) {
    Write-Log "Installed version not found. Marking non-compliant."
    exit 1
}

if ($latest) {
    if ($installed -ge $latest) {
        Write-Log "Office is compliant. Installed: $installed | Latest: $latest"
        exit 0
    } else {
        Write-Log "Office is outdated. Installed: $installed | Latest: $latest"
        exit 1
    }
} else {
    if ($installed -lt $fallbackMinimumBuildVersion) {
        Write-Log "Installed $installed < fallback $fallbackMinimumBuildVersion — Non-compliant."
        exit 1
    } else {
        Write-Log "Installed $installed >= fallback $fallbackMinimumBuildVersion — Compliant."
        exit 0
    }
}
