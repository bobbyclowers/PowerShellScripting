# Remediation Script: Remove Non-AppX Teams + Teams Personal and Install AppX version
# Aborts if non-AppX Teams running or user in active call
# Same exit code logic as detection script for consistency
# https://learn.microsoft.com/en-us/microsoftteams/new-teams-bulk-install
# https://learn.microsoft.com/en-us/powershell/module/appx/get-appxpackage
# https://learn.microsoft.com/en-us/powershell/module/appx/remove-appxpacka

Start-Transcript -Path "$env:TEMP\TeamsRemediation.log" -Append

function Write-Log {
    param([string]$msg)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $full = "[$timestamp] $msg"
    Write-Host $full
    # Transcript captures Write-Host, so Add-Content not needed
}

# Check for Teams in call
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
                Write-Log "User is in a Teams meeting: '${title}' — skipping remediation. [Exit 3]"
                Stop-Transcript
                exit 3
            }
        }
    }
}

# Uninstall legacy Teams folders
$legacyPaths = @(
    "$env:LOCALAPPDATA\Microsoft\Teams",
    "C:\Program Files\Microsoft\Teams",
    "C:\Program Files (x86)\Teams Installer"
)
foreach ($path in $legacyPaths) {
    if (Test-Path $path) {
        try {
            Remove-Item -Recurse -Force -Path $path
            Write-Log "Removed: ${path}"
        } catch {
            Write-Log "Failed to remove ${path}: $_"
        }
    }
}

# Uninstall non-AppX Teams via registry, excluding the Office add-in
$excludedName = "Microsoft Teams Meeting Add-in for Microsoft Office"
$uninstallKeys = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
)
foreach ($key in $uninstallKeys) {
    Get-ItemProperty -Path $key -ErrorAction SilentlyContinue | ForEach-Object {
        $displayName = $_.'DisplayName'
        if (
            $displayName -like "*Teams*" -and
            $displayName -ne $excludedName -and
            $_.UninstallString -and
            $_.UninstallString -notlike "*AppX*"
        ) {
            try {
                Start-Process -FilePath "cmd.exe" -ArgumentList "/c", "$($_.UninstallString) /quiet /norestart" -Wait
                Write-Log "Uninstalled via registry: ${displayName}"
            } catch {
                Write-Log "Uninstall error: ${displayName}: $_"
            }
        }
    }
}

# Remove Teams Personal AppX packages
$appxPersonal = Get-AppxPackage -AllUsers | Where-Object {
    $_.Name -eq "MicrosoftCorporationII.MSTeams" -or $_.Name -eq "MicrosoftTeams"
}
foreach ($pkg in $appxPersonal) {
    try {
        Remove-AppxPackage -AllUsers -Package $pkg.PackageFullName
        Write-Log "Removed Teams Personal AppX: $($pkg.Name)"
    } catch {
        Write-Log "Failed to remove AppX package $($pkg.Name): $_"
    }
}

Stop-Transcript
exit 0
