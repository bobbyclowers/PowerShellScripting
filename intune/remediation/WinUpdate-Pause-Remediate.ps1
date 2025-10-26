# Remediation Script - Unpause and force Windows Update
# PowerShell 5.1 and 7 Compatible

$log = @()

function Log {
    param ($msg)
    $log += $msg
    Write-Output $msg
}

try {
    Log "Starting remediation script."

    # Remove Pause UX settings
    $wuKey = "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings"
    if (Test-Path $wuKey) {
        Remove-ItemProperty -Path $wuKey -Name PauseUpdates -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path $wuKey -Name PauseStartTime -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path $wuKey -Name PauseEndTime -ErrorAction SilentlyContinue
        Log "Removed PauseUpdate registry values from UX\\Settings"
    }

    # Check and remove policy key if applicable
    $policyKey = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
    if (Test-Path $policyKey) {
        Remove-ItemProperty -Path $policyKey -Name SetDisablePauseUXAccess -ErrorAction SilentlyContinue
        Log "Removed SetDisablePauseUXAccess policy value"
    }

    # Reset Windows Update components (minimal)
    net stop wuauserv /y | Out-Null
    net stop bits /y | Out-Null
    net start wuauserv | Out-Null
    net start bits | Out-Null
    Log "Restarted Windows Update services"

    # Re-register WUCOM components (optional)
    # Log "Re-registering WUCOM components..."
    # Start-Process -NoNewWindow -FilePath "powershell" -ArgumentList "-Command `"Get-WmiObject -Namespace root\cimv2 -Class Win32_Process`"" -Wait

    # Trigger update scan and install
    Log "Triggering Windows Update scan and install..."

    $updateSession = New-Object -ComObject Microsoft.Update.Session
    $updateSearcher = $updateSession.CreateUpdateSearcher()
    $searchResult = $updateSearcher.Search("IsInstalled=0 and Type='Software'")

    if ($searchResult.Updates.Count -gt 0) {
        $updatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
        foreach ($update in $searchResult.Updates) {
            $updatesToInstall.Add($update) | Out-Null
            Log "Queued update: $($update.Title)"
        }

        $installer = $updateSession.CreateUpdateInstaller()
        $installer.Updates = $updatesToInstall
        $result = $installer.Install()

        Log "Installation result: $($result.ResultCode)"
        Log "Reboot required: $($result.RebootRequired)"
    }
    else {
        Log "No updates found to install."
    }

    Log "Remediation completed."
    exit 0
}
catch {
    Log "Remediation script failed: $_"
    exit 1
}
