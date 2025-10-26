function Require-Pwsh7 {
    param (
        [string]$ScriptToRelaunch = $MyInvocation.MyCommand.Path
    )

    Write-Host "$($PSStyle.Foreground.White)$($PSStyle.Background.Blue)Test if script running in PowerShell 7$($PSStyle.Reset)`n"

    if ($PSVersionTable.PSEdition -eq 'Core' -and $PSVersionTable.PSVersion.Major -ge 7) {
        Write-Host "✅ PowerShell 7 is already in use. Continuing..." -ForegroundColor Green
        return
    }

    if ($PSVersionTable.PSEdition -ne 'Core' -or $PSVersionTable.PSVersion.Major -lt 7) {
        Write-Host "This script requires PowerShell 7. Relaunching in PowerShell 7..." -ForegroundColor Yellow

        if (Get-Command pwsh -ErrorAction SilentlyContinue) {
            Start-Process -FilePath "pwsh" -ArgumentList "-NoExit", "-File", "`"$ScriptToRelaunch`""
            exit
        } else {
            Write-Host ""
            Write-Host "❌ PowerShell 7 (pwsh.exe) is not installed or not found in PATH." -ForegroundColor Red
            Write-Host ""

            do {
                Write-Host "Choose an option:" -ForegroundColor Cyan
                Write-Host "[I] Install PowerShell 7.4.1 automatically (silent install)"
                Write-Host "[W] Open the official install website"
                Write-Host "[N] Do nothing and exit"
                $choice = Read-Host "Your choice (I/W/N)"
            } while ($choice -notmatch '^[IWNiwn]$')

            switch ($choice.ToUpper()) {
                "W" {
                    Start-Process "https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.5"
                    Write-Host "Opened install page in your browser." -ForegroundColor Green
                }
                "I" {
                    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
                        Write-Host "⚠️  This script must be run as Administrator to install PowerShell 7 silently." -ForegroundColor Yellow
                    } else {
                        try {
                            Write-Host "Downloading and installing PowerShell 7.4.1..." -ForegroundColor Cyan
                            $installerPath = "$env:TEMP\PowerShell-7-x64.msi"
                            $downloadUrl = "https://github.com/PowerShell/PowerShell/releases/latest/download/PowerShell-7.4.1-win-x64.msi"
                            Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath
                            Start-Process "msiexec.exe" -ArgumentList "/i `"$installerPath`" /qn /norestart" -Wait
                            Write-Host "`n✅ PowerShell 7 installed successfully. Relaunching..." -ForegroundColor Green
                            Start-Process -FilePath "pwsh" -ArgumentList "-NoExit", "-File", "`"$ScriptToRelaunch`""
                            exit
                        } catch {
                            Write-Host "❌ Failed to install PowerShell 7: $_" -ForegroundColor Red
                        }
                    }
                }
                "N" {
                    Write-Host "Exiting script. PowerShell 7 is required to continue." -ForegroundColor Yellow
                }
            }

            Write-Host ""
            Write-Host "Press Enter to exit..." -ForegroundColor Cyan
            Read-Host
            exit 1
        }
    }


}
