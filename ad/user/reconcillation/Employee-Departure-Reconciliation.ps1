<#
.SYNOPSIS
    Canonical Employee Departure Reconciliation script

.DESCRIPTION
    Processes the latest departure CSV placed in a configured folder, disables any active accounts
    matching the Employee ID provided in the CSV, writes a report, and records any failures
    that require manual follow-up.

.NOTES
    - Designed to be generic about organisational teams (uses neutral wording for service teams).
    - Manager lookup is performed by searching AD for the manager's EmployeeID supplied in the CSV.
    - The CSV is assumed to include a manager employee id column named 'managerEmployeeID' if manager linking is desired.

.FILECREATED
    2025-10-29

.FILELASTUPDATED
    2025-10-29
#>

## Header and configuration
[CmdletBinding()]
param()

# Configuration - adjust these paths if required by your environment
$SourceDirectory      = 'V:\Employee-Departure-Check'
$DestinationFolderComplete = 'V:\Employee-Departure-Check\Complete'
$DestinationFolderReports  = 'V:\Employee-Departure-Check\Reports'

# Logging
$logFile = Join-Path $env:TEMP 'Employee-Departure-Reconciliation.log'
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = 'INFO'
    )
    $ts = (Get-Date).ToString('s')
    $line = "[$ts] [$Level] $Message"
    Add-Content -Path $logFile -Value $line -ErrorAction SilentlyContinue
}

Write-Log "Script start"

try {
    # Ensure report directories exist
    foreach ($p in @($DestinationFolderReports, $DestinationFolderComplete)) {
        if (-not (Test-Path -Path $p)) {
            New-Item -Path $p -ItemType Directory -Force | Out-Null
            Write-Log "Created missing directory: $p"
        }
    }
}
catch {
    Write-Log "Failed to ensure destination directories exist: $($_.Exception.Message)" 'ERROR'
    throw
}

# Prepare report and failure files
$today = Get-Date -Format 'yyyyMMdd'
$reportFile = Join-Path $DestinationFolderReports "$today.csv"
$failureFile = Join-Path $DestinationFolderReports "$today.failures.csv"

# Header for report file if missing
if (-not (Test-Path -Path $reportFile)) {
    'EmployeeIDPS,EmployeeIDAD,UserPrincipalName,DisplayName,ManagerEmployeeID,ManagerName,DepartureDate,Division,Department,Team' | Out-File -FilePath $reportFile -Encoding UTF8
    Write-Log "Created report file with header: $reportFile"
}

# Header for failure file if missing
if (-not (Test-Path -Path $failureFile)) {
    'EmployeeIDPS,Reason,Details' | Out-File -FilePath $failureFile -Encoding UTF8
    Write-Log "Created failure file with header: $failureFile"
}

## Select latest file or prompt
$fileType = '*.csv'
$files = Get-ChildItem -Path $SourceDirectory -Filter $fileType -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending

if (-not $files -or $files.Count -eq 0) {
    Write-Log "No CSV files found in $SourceDirectory" 'ERROR'
    Write-Host "❌ Error - No CSV File found in $SourceDirectory" -ForegroundColor Red
    exit 1
}

if ($files.Count -eq 1) {
    $selectedFile = $files[0]
    Write-Host "Using file: $($selectedFile.Name)"
    Write-Log "Single file selected: $($selectedFile.FullName)"
}
else {
    Write-Host "Multiple files found. Select a file:" -ForegroundColor Yellow
    for ($i = 0; $i -lt $files.Count; $i++) {
        Write-Host "[$i] $($files[$i].Name) - Last Modified: $($files[$i].LastWriteTime)"
    }
    do {
        $selection = Read-Host 'Enter the number of the file you want to select (default 0)'
        if ($selection -eq '') { $selection = 0 }
    } while ($selection -notmatch '^[0-9]+$' -or [int]$selection -lt 0 -or [int]$selection -ge $files.Count)
    $selectedFile = $files[[int]$selection]
    Write-Log "User selected file: $($selectedFile.FullName)"
}

$csvPath = $selectedFile.FullName

## Import CSV
try {
    $csvData = Import-Csv -Path $csvPath -ErrorAction Stop
    Write-Log "Imported CSV $csvPath with $($csvData.Count) rows"
}
catch {
        Write-Log -Message ("Failed to import CSV {0}: {1}" -f $csvPath, $_.Exception.Message) -Level 'ERROR'
    Write-Host "❌ Error - Failed to import CSV file at path '$csvPath'." -ForegroundColor Red
    exit 1
}

$infoWrittenToCSV = $false

## Process rows
foreach ($row in $csvData) {
    # Primary employee identifier in CSV
    $employeeIDPS = $row.detnumber
        $employeeEmail = $row.detemailad
    $employeeDepartureDate = $row.detterdate
    $employeeDivision = $row.'pdtorg2cd.trn'
    $employeeDepartment = $row.'pdtorg3cd.trn'
    $employeeTeam = $row.'pdtorg4cd.trn'

    # ManagerEmployeeID is expected to be provided in the CSV if available
    $managerEmployeeID = $null
    if ($row.PSObject.Properties.Name -contains 'managerEmployeeID') { $managerEmployeeID = $row.managerEmployeeID }

    # Attempt to find the AD user by EmployeeID
    try {
        $user = Get-ADUser -Filter { EmployeeID -eq $employeeIDPS } -Properties EmployeeID,Enabled,UserPrincipalName,DisplayName | Select-Object -First 1
    }
    catch {
        Write-Log -Message ("Get-ADUser failed for EmployeeID '{0}': {1}" -f $employeeIDPS, $_.Exception.Message) -Level 'ERROR'
        Add-Content -Path $failureFile -Value ("{0},ADQueryFailed,{1}" -f $employeeIDPS, $_.Exception.Message)
        continue
    }

    if ($user -and $user.Enabled) {
        # Attempt disable with error handling
        try {
            Disable-ADAccount -Identity $user -ErrorAction Stop
            Write-Log "Disabled account for EmployeeID $employeeIDPS (UPN: $($user.UserPrincipalName))"
        }
        catch {
            Write-Log -Message ("Failed to disable account for EmployeeID {0}: {1}" -f $employeeIDPS, $_.Exception.Message) -Level 'ERROR'
            Add-Content -Path $failureFile -Value ("{0},DisableFailed,{1}" -f $employeeIDPS, $_.Exception.Message)
            # continue processing other rows
        }

        # Resolve manager name by searching AD for managerEmployeeID (if provided)
        $managerName = 'Not provided'
        if ($managerEmployeeID) {
            try {
                $manager = Get-ADUser -Filter { EmployeeID -eq $managerEmployeeID } -Properties Name | Select-Object -First 1
                if ($manager) { $managerName = $manager.Name }
                else {
                    $managerName = 'Manager not found'
                    Write-Log "Manager not found for managerEmployeeID '$managerEmployeeID' (row employee $employeeIDPS)" 'WARN'
                    Add-Content -Path $failureFile -Value "$employeeIDPS,ManagerLookupFailed,ManagerEmployeeID:$managerEmployeeID"
                }
            }
            catch {
                Write-Log -Message ("Manager lookup failed for managerEmployeeID '{0}': {1}" -f $managerEmployeeID, $_.Exception.Message) -Level 'ERROR'
                Add-Content -Path $failureFile -Value ("{0},ManagerLookupError,{1}" -f $employeeIDPS, $_.Exception.Message)
            }
        }

        # Append result to report (simple CSV line as per existing approach)
        $employeeIDAD = $user.EmployeeID
        $userPrincipalName = $user.UserPrincipalName
        $displayName = $user.DisplayName
        $output = "$employeeIDPS,$employeeIDAD,$userPrincipalName,$displayName,$managerEmployeeID,$managerName,$employeeDepartureDate,$employeeDivision,$employeeDepartment,$employeeTeam"
        Add-Content -Path $reportFile -Value $output
        $infoWrittenToCSV = $true
    }
    else {
        # Not found or already disabled - no action required, but for visibility add a note if user exists but disabled
        if ($user -and -not $user.Enabled) {
            Write-Log "User for EmployeeID $employeeIDPS already disabled; skipping"
        }
        else {
            Write-Log "No AD account found for EmployeeID $employeeIDPS" 'WARN'
            Add-Content -Path $failureFile -Value ("{0},UserNotFound,{1}" -f $employeeIDPS, $employeeEmail)
        }
    }
}

## Post-processing and reporting
if ($infoWrittenToCSV) {
    Write-Host "❌ ACTION REQUIRED - Some accounts were disabled and a report has been written to: $reportFile" -ForegroundColor Yellow
    Write-Log "Action required: report written to $reportFile"
}
else {
    Write-Host "✅ NO ACTION REQUIRED - No active accounts required disabling." -ForegroundColor Green
    Write-Log "No accounts required disabling"
}

if ((Get-Item -Path $failureFile -ErrorAction SilentlyContinue)) {
    $failCount = (Get-Content -Path $failureFile | Measure-Object -Line).Lines - 1 # subtract header
    if ($failCount -gt 0) {
        Write-Host "⚠️  $failCount items require manual follow-up. See: $failureFile" -ForegroundColor Yellow
        Write-Log "$failCount failure items recorded in $failureFile" 'WARN'
    }
}

# Move processed file to complete folder
try {
    Move-Item -Path $csvPath -Destination (Join-Path $DestinationFolderComplete $selectedFile.Name) -Force
    Write-Log "Moved processed file to $DestinationFolderComplete"
}
catch {
    Write-Log "Failed to move processed file: $($_.Exception.Message)" 'ERROR'
}

Write-Log "Script end"
