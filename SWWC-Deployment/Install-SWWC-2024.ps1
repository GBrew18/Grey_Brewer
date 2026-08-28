<#
.SYNOPSIS
    Installs Windchill Workgroup Manager (WGM) and SOLIDWORKS 2024 SP5.0 from
    pre-staged zip payloads on Windows 11 workstations.

.DESCRIPTION
    Deployment order (hard requirement):
        1. Prechecks (fail fast, clear messages, nothing modified on failure)
        2. Extract + install Windchill Workgroup Manager, wait, validate
        3. Extract + install SOLIDWORKS 2024, wait, validate
        4. Print summary table (also written to the log)

    If any step fails, the script stops safely at that step, reports why, and
    returns a distinct exit code. It never proceeds to SOLIDWORKS if WGM did
    not install and validate cleanly.

    Validation per product = registry uninstall entry present (name + version
    pattern) AND install directory exists on disk. Installer exit code and
    elapsed time are captured and reported for every install step.

    Designed to run as SYSTEM (e.g. deployed via SCCM/Intune/PDQ/psexec -s).

.PARAMETER CheckOnly
    Run all prechecks and report, but do not extract or install anything.

.PARAMETER ForceExtract
    Re-extract zips even if a previous extraction of the same zip is present.

.PARAMETER AllowNonSystem
    Testing escape hatch: permit running as an elevated administrator instead
    of SYSTEM. Production runs should NOT use this.

.EXAMPLE
    psexec -s powershell.exe -ExecutionPolicy Bypass -File C:\SWWC\Install-SWWC-2024.ps1

.EXAMPLE
    powershell.exe -File .\Install-SWWC-2024.ps1 -CheckOnly -AllowNonSystem

.NOTES
    Exit codes:
         0  = success (check log/summary for "reboot required" flag)
        10  = precheck failure
        20  = WGM install or validation failure
        30  = SOLIDWORKS install or validation failure
        99  = unexpected script error
#>

[CmdletBinding()]
param(
    [switch]$CheckOnly,
    [switch]$ForceExtract,
    [switch]$AllowNonSystem
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# =============================================================================
#  CONFIGURATION - review this block before first deployment
# =============================================================================

$StageRoot   = 'C:\SWWC'
$ExtractRoot = 'C:\SWWC\Extract'
$LogRoot     = 'C:\SWWC\Logs'

$WgmZip = 'C:\SWWC\MED-60707-CD-131_13-1-2-2_Windchill-Workgroup-Managers.zip'
$SwZip  = 'C:\SWWC\Solidworks_2024\SOLIDWORKS_2024_SP5.0.zip'

# Free space required on C: (GB). A dynamic estimate based on zip sizes is
# also computed; the larger of the two is enforced.
$MinimumFreeSpaceGB = 60

# --- Windchill Workgroup Manager -------------------------------------------
# The script auto-detects the WGM client MSI inside the extracted CD image
# (preferring a path that mentions SOLIDWORKS, since the CD ships workgroup
# managers for several CAD systems). If your CD layout differs, set
# $WgmInstallerOverride to the installer path RELATIVE to the extracted CD
# root and $WgmInstallerOverrideArgs to its silent arguments.
$WgmInstallerOverride     = ''      # e.g. 'solidworks\wgmclient.msi' or 'setup.exe'
$WgmInstallerOverrideArgs = ''      # e.g. '-silent -DINSTALL_DIR=C:\ptc\wgm'

# Extra MSI properties appended when installing via wgmclient.msi
# (e.g. 'INSTALLDIR="C:\ptc\wgm"'). Leave empty for vendor defaults.
$WgmMsiExtraProperties = ''

# Validation: registry display-name/version patterns and fallback directory.
# The registry InstallLocation is preferred when the installer records one.
$WgmDisplayNamePattern = '*Windchill Workgroup Manager*'
$WgmVersionPattern     = '13.1*'
$WgmFallbackInstallDir = 'C:\Program Files\PTC\Windchill Workgroup Manager'
$WgmInstallTimeoutMin  = 60

# --- SOLIDWORKS 2024 --------------------------------------------------------
# The zip is expected to contain a SOLIDWORKS ADMINISTRATIVE IMAGE
# (StartSWInstall.exe at its root or one level down). Admin images embed the
# serial number and options, which is the supported way to script SOLIDWORKS
# across a fleet. Raw download media is detected and rejected with guidance.
$SwDisplayNamePattern = 'SOLIDWORKS 2024*'
$SwVersionPattern     = '32.5*'          # 2024 SP5.0 = 32.5.x
$SwFallbackInstallDir = 'C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS'
$SwInstallTimeoutMin  = 240
$SldimStartTimeoutSec = 300              # how long to wait for sldIM.exe to appear

# --- Blocking processes (precheck fails while any of these run) -------------
$BlockingProcesses = @(
    # SOLIDWORKS
    'SLDWORKS', 'sldworks_fs', 'swVBAServer', 'sldProcMon',
    'SOLIDWORKSBackgroundDownloader', 'swShellFileLauncher', 'eDrawings',
    # PTC / Windchill / Creo (WGM upgrades fail if the client or CAD is open)
    'wgmclient', 'xtop', 'creoagent',
    # Installers already in flight
    'sldIM', 'setup', 'setup_win64'
)

# =============================================================================
#  LOGGING AND RESULT TRACKING
# =============================================================================

$script:LogFile = $null
$script:Results = New-Object System.Collections.ArrayList
$script:RebootRequired = $false

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR', 'STEP')][string]$Level = 'INFO'
    )
    $line = ('{0} [{1,-5}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message)
    if ($script:LogFile) {
        Add-Content -Path $script:LogFile -Value $line -ErrorAction SilentlyContinue
    }
    switch ($Level) {
        'OK'    { Write-Host $line -ForegroundColor Green }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'STEP'  { Write-Host ''; Write-Host $line -ForegroundColor Cyan }
        default { Write-Host $line }
    }
}

function Add-Result {
    param(
        [Parameter(Mandatory)][string]$Step,
        [Parameter(Mandatory)][string]$Status,   # Passed / Failed / Skipped / Installed / AlreadyInstalled
        [object]$ExitCode = '',
        [object]$ElapsedSec = '',
        [string]$Detail = ''
    )
    [void]$script:Results.Add([pscustomobject]@{
        Step       = $Step
        Status     = $Status
        ExitCode   = $ExitCode
        ElapsedSec = $ElapsedSec
        Detail     = $Detail
    })
}

function Write-Summary {
    param([Parameter(Mandatory)][int]$FinalExitCode)
    $banner = '=' * 78
    Write-Host ''
    Write-Host $banner -ForegroundColor Cyan
    Write-Host '  DEPLOYMENT SUMMARY' -ForegroundColor Cyan
    Write-Host $banner -ForegroundColor Cyan
    $table = $script:Results | Format-Table -AutoSize Step, Status, ExitCode, ElapsedSec, Detail | Out-String
    Write-Host $table
    if ($script:LogFile) { Add-Content -Path $script:LogFile -Value $table }
    if ($script:RebootRequired) {
        Write-Log 'One or more installers requested a REBOOT (exit code 3010/1641). Reboot this workstation before use.' 'WARN'
    }
    if ($FinalExitCode -eq 0) {
        Write-Log "RESULT: SUCCESS (exit code 0). Full log: $($script:LogFile)" 'OK'
    } else {
        Write-Log "RESULT: FAILED (exit code $FinalExitCode). Full log: $($script:LogFile)" 'ERROR'
    }
    Write-Host $banner -ForegroundColor Cyan
}

function Stop-Deployment {
    param(
        [Parameter(Mandatory)][string]$Message,
        [Parameter(Mandatory)][int]$ExitCode
    )
    Write-Log $Message 'ERROR'
    Write-Summary -FinalExitCode $ExitCode
    exit $ExitCode
}

# =============================================================================
#  PRECHECKS
# =============================================================================

function Test-Prechecks {
    Write-Log 'PRECHECKS' 'STEP'
    $failures = New-Object System.Collections.ArrayList

    # -- PowerShell version --------------------------------------------------
    $psv = $PSVersionTable.PSVersion
    if ($psv.Major -lt 5 -or ($psv.Major -eq 5 -and $psv.Minor -lt 1)) {
        [void]$failures.Add("PowerShell 5.1 or later required; running $psv.")
    } else {
        Write-Log "PowerShell version: $psv" 'OK'
    }

    # -- 64-bit process ------------------------------------------------------
    if (-not [Environment]::Is64BitProcess) {
        [void]$failures.Add('Running in a 32-bit PowerShell host. Launch the 64-bit powershell.exe (C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe) so 64-bit registry and installers behave correctly.')
    } else {
        Write-Log '64-bit PowerShell process: yes' 'OK'
    }

    # -- SYSTEM identity -----------------------------------------------------
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    if ($identity.IsSystem) {
        Write-Log "Running as: $($identity.Name) (SYSTEM)" 'OK'
    } elseif ($AllowNonSystem) {
        $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
        if ($principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
            Write-Log "Running as: $($identity.Name) - NOT SYSTEM, permitted by -AllowNonSystem (elevated admin confirmed)" 'WARN'
        } else {
            [void]$failures.Add("-AllowNonSystem was passed but the current session ($($identity.Name)) is not elevated. Run from an elevated prompt.")
        }
    } else {
        [void]$failures.Add("Not running as SYSTEM (current identity: $($identity.Name)). Deploy via your management tool or 'psexec -s'. For testing only, re-run with -AllowNonSystem from an elevated prompt.")
    }

    # -- Staged payload ------------------------------------------------------
    if (-not (Test-Path -LiteralPath $StageRoot -PathType Container)) {
        [void]$failures.Add("Staging folder not found: $StageRoot")
    } else {
        Write-Log "Staging folder present: $StageRoot" 'OK'
    }

    $zipBytesTotal = 0
    foreach ($zip in @($WgmZip, $SwZip)) {
        if (-not (Test-Path -LiteralPath $zip -PathType Leaf)) {
            [void]$failures.Add("Required payload missing: $zip")
            continue
        }
        $item = Get-Item -LiteralPath $zip
        if ($item.Length -lt 1MB) {
            [void]$failures.Add("Payload looks truncated/corrupt (size $($item.Length) bytes): $zip")
            continue
        }
        $zipBytesTotal += $item.Length
        Write-Log ("Payload present: {0} ({1:N1} GB)" -f $zip, ($item.Length / 1GB)) 'OK'
    }

    # -- Disk space ----------------------------------------------------------
    # Need room for extraction (~1x zip size) plus the installed footprint;
    # 2.5x total zip size + 10 GB headroom is a conservative estimate.
    $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'"
    $freeGB = [math]::Round($disk.FreeSpace / 1GB, 1)
    $dynamicNeedGB = [math]::Ceiling(($zipBytesTotal * 2.5) / 1GB) + 10
    $requiredGB = [math]::Max($MinimumFreeSpaceGB, $dynamicNeedGB)
    if ($freeGB -lt $requiredGB) {
        [void]$failures.Add("Insufficient disk space on C: - $freeGB GB free, $requiredGB GB required (configured minimum $MinimumFreeSpaceGB GB / payload-based estimate $dynamicNeedGB GB).")
    } else {
        Write-Log "Disk space on C: $freeGB GB free ($requiredGB GB required)" 'OK'
    }

    # -- Blocking processes --------------------------------------------------
    $running = @(Get-Process -Name $BlockingProcesses -ErrorAction SilentlyContinue)
    if ($running.Count -gt 0) {
        $names = ($running | Select-Object -ExpandProperty ProcessName -Unique) -join ', '
        [void]$failures.Add("Blocking process(es) running: $names. Close SOLIDWORKS / Creo / Workgroup Manager and any installers, then re-run. (This script deliberately does not kill processes.)")
    } else {
        Write-Log 'No blocking processes running' 'OK'
    }

    # -- Windows Installer already busy? -------------------------------------
    # The Global\_MSIExecute mutex is held only while an MSI is actively
    # installing; a lingering idle msiexec.exe process does not hold it.
    $msiBusy = $false
    try {
        $mutex = [System.Threading.Mutex]::OpenExisting('Global\_MSIExecute')
        $mutex.Dispose()
        $msiBusy = $true
    } catch [System.Threading.WaitHandleCannotBeOpenedException] {
        $msiBusy = $false
    } catch [System.UnauthorizedAccessException] {
        $msiBusy = $true
    }
    if ($msiBusy) {
        [void]$failures.Add('Another Windows Installer (MSI) operation is in progress. Wait for it to finish and re-run.')
    } else {
        Write-Log 'Windows Installer service is idle' 'OK'
    }

    # -- Pending reboot (warn only) ------------------------------------------
    $pendingReboot = $false
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $pendingReboot = $true }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $pendingReboot = $true }
    if ($pendingReboot) {
        Write-Log 'A Windows reboot is pending on this machine. Installation will proceed, but if an installer fails, reboot and retry first.' 'WARN'
    } else {
        Write-Log 'No pending reboot detected' 'OK'
    }

    # -- Verdict -------------------------------------------------------------
    if ($failures.Count -gt 0) {
        foreach ($f in $failures) { Write-Log "PRECHECK FAILED: $f" 'ERROR' }
        Add-Result -Step 'Prechecks' -Status 'Failed' -Detail ("{0} check(s) failed - see log" -f $failures.Count)
        Stop-Deployment -Message ("{0} precheck(s) failed. Nothing was installed or modified." -f $failures.Count) -ExitCode 10
    }
    Add-Result -Step 'Prechecks' -Status 'Passed'
    Write-Log 'All prechecks passed' 'OK'
}

# =============================================================================
#  HELPERS
# =============================================================================

function Expand-Payload {
    <# Extracts a zip to $ExtractRoot\<zip basename>. A marker file recording
       the zip's size + timestamp makes re-runs skip an identical extraction
       (override with -ForceExtract). Returns the extraction folder path. #>
    param([Parameter(Mandatory)][string]$ZipPath)

    $zipItem  = Get-Item -LiteralPath $ZipPath
    $destName = [System.IO.Path]::GetFileNameWithoutExtension($zipItem.Name)
    $dest     = Join-Path $ExtractRoot $destName
    $marker   = Join-Path $dest '.swwc_extract_ok'
    $stamp    = '{0}|{1:o}' -f $zipItem.Length, $zipItem.LastWriteTimeUtc

    if (-not $ForceExtract -and (Test-Path -LiteralPath $marker)) {
        $existing = Get-Content -LiteralPath $marker -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($existing -eq $stamp) {
            Write-Log "Reusing previous extraction of $($zipItem.Name) at $dest (zip unchanged). Use -ForceExtract to re-extract." 'INFO'
            return $dest
        }
    }

    if (Test-Path -LiteralPath $dest) {
        Write-Log "Removing stale extraction folder: $dest"
        Remove-Item -LiteralPath $dest -Recurse -Force
    }
    New-Item -ItemType Directory -Path $dest -Force | Out-Null

    Write-Log ("Extracting {0} ({1:N1} GB) to {2} - this can take several minutes..." -f $zipItem.Name, ($zipItem.Length / 1GB), $dest)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $dest)
    $sw.Stop()
    Set-Content -LiteralPath $marker -Value $stamp
    Write-Log ("Extraction complete in {0:N0}s" -f $sw.Elapsed.TotalSeconds) 'OK'
    return $dest
}

function Get-InstalledProduct {
    <# Scans 64-bit and 32-bit uninstall hives for a product whose DisplayName
       and DisplayVersion match the given wildcard patterns. #>
    param(
        [Parameter(Mandatory)][string]$DisplayNamePattern,
        [string]$VersionPattern = '*'
    )
    $hives = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    Get-ItemProperty -Path $hives -ErrorAction SilentlyContinue |
        Where-Object {
            ($_.PSObject.Properties['DisplayName'] -and $_.DisplayName -like $DisplayNamePattern) -and
            ((-not $_.PSObject.Properties['DisplayVersion']) -or $_.DisplayVersion -like $VersionPattern)
        } |
        Select-Object DisplayName, DisplayVersion, InstallLocation
}

function Test-ProductInstalled {
    <# Validates: (1) registry uninstall entry matches name+version patterns,
       (2) install directory exists (registry InstallLocation preferred,
       else the configured fallback path). Returns $true/$false and logs. #>
    param(
        [Parameter(Mandatory)][string]$ProductLabel,
        [Parameter(Mandatory)][string]$DisplayNamePattern,
        [Parameter(Mandatory)][string]$VersionPattern,
        [Parameter(Mandatory)][string]$FallbackInstallDir
    )

    $match = @(Get-InstalledProduct -DisplayNamePattern $DisplayNamePattern -VersionPattern $VersionPattern) | Select-Object -First 1
    if (-not $match) {
        # Distinguish "wrong version" from "not present" for clearer messages.
        $anyVersion = @(Get-InstalledProduct -DisplayNamePattern $DisplayNamePattern) | Select-Object -First 1
        if ($anyVersion) {
            Write-Log ("{0}: found '{1}' version '{2}' in registry, but it does not match expected version pattern '{3}'." -f $ProductLabel, $anyVersion.DisplayName, $anyVersion.DisplayVersion, $VersionPattern) 'ERROR'
        } else {
            Write-Log ("{0}: no registry uninstall entry matching '{1}' was found." -f $ProductLabel, $DisplayNamePattern) 'ERROR'
        }
        return $false
    }
    Write-Log ("{0}: registry entry OK - '{1}' version '{2}'" -f $ProductLabel, $match.DisplayName, $match.DisplayVersion) 'OK'

    $dir = $FallbackInstallDir
    if ($match.PSObject.Properties['InstallLocation'] -and $match.InstallLocation) {
        $dir = $match.InstallLocation.TrimEnd('\')
    }
    if (-not (Test-Path -LiteralPath $dir -PathType Container)) {
        Write-Log ("{0}: expected install directory does not exist: {1}" -f $ProductLabel, $dir) 'ERROR'
        return $false
    }
    Write-Log ("{0}: install directory OK - {1}" -f $ProductLabel, $dir) 'OK'
    return $true
}

function Invoke-TrackedInstaller {
    <# Runs an installer, waits for completion (with timeout), and returns a
       result object with ExitCode / ElapsedSec / Success. Exit codes 0, 3010
       and 1641 count as success; 3010/1641 set the global reboot flag. #>
    param(
        [Parameter(Mandatory)][string]$Label,
        [Parameter(Mandatory)][string]$FilePath,
        [string[]]$ArgumentList = @(),
        [Parameter(Mandatory)][int]$TimeoutMin
    )

    Write-Log ("Launching {0}: `"{1}`" {2}" -f $Label, $FilePath, ($ArgumentList -join ' '))
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    if ($ArgumentList.Count -gt 0) {
        $proc = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -PassThru
    } else {
        $proc = Start-Process -FilePath $FilePath -PassThru
    }
    $null = $proc.Handle  # cache the handle so ExitCode is readable after exit (PS 5.1 quirk)
    if (-not $proc.WaitForExit($TimeoutMin * 60 * 1000)) {
        $sw.Stop()
        Write-Log ("{0} did not finish within {1} minutes. Leaving the installer process running (PID {2}) - investigate before retrying." -f $Label, $TimeoutMin, $proc.Id) 'ERROR'
        return [pscustomobject]@{ Success = $false; ExitCode = 'timeout'; ElapsedSec = [math]::Round($sw.Elapsed.TotalSeconds) }
    }
    $sw.Stop()
    $code = $proc.ExitCode
    $elapsed = [math]::Round($sw.Elapsed.TotalSeconds)
    if ($code -eq 3010 -or $code -eq 1641) { $script:RebootRequired = $true }
    $ok = ($code -eq 0 -or $code -eq 3010 -or $code -eq 1641)
    if ($ok) {
        Write-Log ("{0} finished: exit code {1} in {2}s" -f $Label, $code, $elapsed) 'OK'
    } else {
        Write-Log ("{0} FAILED: exit code {1} after {2}s" -f $Label, $code, $elapsed) 'ERROR'
    }
    return [pscustomobject]@{ Success = $ok; ExitCode = $code; ElapsedSec = $elapsed }
}

# =============================================================================
#  STEP 1: WINDCHILL WORKGROUP MANAGER
# =============================================================================

function Install-Wgm {
    Write-Log 'STEP 1: WINDCHILL WORKGROUP MANAGER' 'STEP'

    if (Test-ProductInstalled -ProductLabel 'WGM (pre-scan)' -DisplayNamePattern $WgmDisplayNamePattern -VersionPattern $WgmVersionPattern -FallbackInstallDir $WgmFallbackInstallDir) {
        Write-Log 'Target WGM version already installed - skipping install.' 'OK'
        Add-Result -Step 'WGM install' -Status 'AlreadyInstalled'
        Add-Result -Step 'WGM validation' -Status 'Passed'
        return
    }

    $cdRoot = Expand-Payload -ZipPath $WgmZip

    # Locate the installer inside the extracted CD image.
    $installerPath = $null
    $installerArgs = @()
    $msiLog = Join-Path $LogRoot ('WGM_msi_{0}.log' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

    if ($WgmInstallerOverride) {
        $installerPath = Join-Path $cdRoot $WgmInstallerOverride
        if (-not (Test-Path -LiteralPath $installerPath -PathType Leaf)) {
            Add-Result -Step 'WGM install' -Status 'Failed' -Detail 'Configured override installer not found'
            Stop-Deployment -Message "Configured WGM installer not found: $installerPath" -ExitCode 20
        }
        if ($installerPath -like '*.msi') {
            $installerArgs = @('/i', "`"$installerPath`"", '/qn', '/norestart', '/l*v', "`"$msiLog`"")
            if ($WgmMsiExtraProperties) { $installerArgs += $WgmMsiExtraProperties }
            $installerPath = "$env:SystemRoot\System32\msiexec.exe"
        } elseif ($WgmInstallerOverrideArgs) {
            $installerArgs = $WgmInstallerOverrideArgs -split '\s+'
        }
    } else {
        # Prefer a wgmclient.msi under a SOLIDWORKS-named path, then any wgmclient.msi.
        $msis = @(Get-ChildItem -Path $cdRoot -Recurse -Filter '*.msi' -ErrorAction SilentlyContinue |
                  Where-Object { $_.Name -match 'wgm' })
        $chosen = $msis | Where-Object { $_.FullName -match 'solidworks|swx' } | Select-Object -First 1
        if (-not $chosen) { $chosen = $msis | Select-Object -First 1 }
        if (-not $chosen) {
            $found = @(Get-ChildItem -Path $cdRoot -Recurse -Include 'setup.exe', 'setup_win64.exe', '*.msi' -ErrorAction SilentlyContinue | Select-Object -First 10 -ExpandProperty FullName)
            Write-Log ("No WGM client MSI auto-detected. Installer candidates found on the CD:`n  " + ($found -join "`n  ")) 'ERROR'
            Add-Result -Step 'WGM install' -Status 'Failed' -Detail 'No installer auto-detected'
            Stop-Deployment -Message 'Could not auto-detect the WGM installer. Set $WgmInstallerOverride (and args if not an MSI) in the config block to the correct path from the list above.' -ExitCode 20
        }
        Write-Log "Auto-detected WGM installer: $($chosen.FullName)"
        $installerArgs = @('/i', "`"$($chosen.FullName)`"", '/qn', '/norestart', '/l*v', "`"$msiLog`"")
        if ($WgmMsiExtraProperties) { $installerArgs += $WgmMsiExtraProperties }
        $installerPath = "$env:SystemRoot\System32\msiexec.exe"
    }

    $result = Invoke-TrackedInstaller -Label 'WGM installer' -FilePath $installerPath -ArgumentList $installerArgs -TimeoutMin $WgmInstallTimeoutMin
    Add-Result -Step 'WGM install' -Status $(if ($result.Success) { 'Installed' } else { 'Failed' }) -ExitCode $result.ExitCode -ElapsedSec $result.ElapsedSec -Detail $(if (Test-Path $msiLog) { "MSI log: $msiLog" } else { '' })
    if (-not $result.Success) {
        Stop-Deployment -Message "WGM installer failed (exit code $($result.ExitCode)). SOLIDWORKS installation was NOT attempted. See $msiLog for details." -ExitCode 20
    }

    Write-Log 'Validating WGM installation...'
    if (Test-ProductInstalled -ProductLabel 'WGM' -DisplayNamePattern $WgmDisplayNamePattern -VersionPattern $WgmVersionPattern -FallbackInstallDir $WgmFallbackInstallDir) {
        Add-Result -Step 'WGM validation' -Status 'Passed'
        Write-Log 'WGM installed and validated.' 'OK'
    } else {
        Add-Result -Step 'WGM validation' -Status 'Failed'
        Stop-Deployment -Message 'WGM validation failed after an apparently successful install. SOLIDWORKS installation was NOT attempted. Check the version patterns in the config block against what the installer actually registered (see log above).' -ExitCode 20
    }
}

# =============================================================================
#  STEP 2: SOLIDWORKS 2024
# =============================================================================

function Install-SolidWorks {
    Write-Log 'STEP 2: SOLIDWORKS 2024 SP5.0' 'STEP'

    if (Test-ProductInstalled -ProductLabel 'SOLIDWORKS (pre-scan)' -DisplayNamePattern $SwDisplayNamePattern -VersionPattern $SwVersionPattern -FallbackInstallDir $SwFallbackInstallDir) {
        Write-Log 'Target SOLIDWORKS version already installed - skipping install.' 'OK'
        Add-Result -Step 'SOLIDWORKS install' -Status 'AlreadyInstalled'
        Add-Result -Step 'SOLIDWORKS validation' -Status 'Passed'
        return
    }

    $swRoot = Expand-Payload -ZipPath $SwZip

    # An administrative image is required for unattended install. Confirmed
    # layout (per pilot machine C:\SWWC\Solidworks_2024\SOLIDWORKS_2024_SP5.0):
    # StartSWInstall.exe and sldIM.exe both live under a \sldim\ subfolder.
    # Search unbounded (no -Depth limit) since some zips add an extra
    # top-level wrapper folder, which would otherwise push it out of reach.
    $startSw = Get-ChildItem -Path $swRoot -Filter 'StartSWInstall.exe' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $startSw) {
        $hasMedia = Test-Path (Join-Path $swRoot 'setup.exe')
        if (-not $hasMedia) {
            $sub = Get-ChildItem -Path $swRoot -Directory | Select-Object -First 1
            if ($sub) { $hasMedia = Test-Path (Join-Path $sub.FullName 'setup.exe') }
        }
        Add-Result -Step 'SOLIDWORKS install' -Status 'Failed' -Detail 'No admin image in zip'
        if ($hasMedia) {
            Stop-Deployment -Message 'The SOLIDWORKS zip contains raw installation media (setup.exe), not an administrative image. Unattended fleet installs require an admin image (it embeds the serial number and install options). On one machine, run the media''s sldIM Installation Manager, choose "Administrative image", configure options, then zip the resulting image folder (containing StartSWInstall.exe) and restage it as this payload.' -ExitCode 30
        } else {
            Stop-Deployment -Message "Could not find StartSWInstall.exe (admin image) or setup.exe (media) in the extracted SOLIDWORKS payload at $swRoot. Verify the zip contents." -ExitCode 30
        }
    }
    Write-Log "Found SOLIDWORKS admin image: $($startSw.FullName)"

    # StartSWInstall.exe /install /now kicks off sldIM.exe and returns; the
    # real installer is sldIM, so we wait for it to appear and then to exit.
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $launcher = Start-Process -FilePath $startSw.FullName -ArgumentList '/install', '/now' -PassThru -WorkingDirectory $startSw.DirectoryName
    $null = $launcher.Handle  # cache the handle so ExitCode is readable after exit (PS 5.1 quirk)
    $launcher.WaitForExit(120000) | Out-Null

    Write-Log 'Waiting for the SOLIDWORKS Installation Manager (sldIM.exe) to start...'
    $sldim = $null
    $deadline = (Get-Date).AddSeconds($SldimStartTimeoutSec)
    while ((Get-Date) -lt $deadline) {
        $sldim = Get-Process -Name 'sldIM' -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($sldim) { break }
        Start-Sleep -Seconds 5
    }

    $exitCode = 'n/a'
    if ($sldim) {
        try { $null = $sldim.Handle } catch { }  # cache handle so ExitCode is readable after exit
        Write-Log "sldIM.exe running (PID $($sldim.Id)). Waiting for install to complete (timeout $SwInstallTimeoutMin min) - a full install typically takes 20-60 minutes..."
        if (-not $sldim.WaitForExit($SwInstallTimeoutMin * 60 * 1000)) {
            $sw.Stop()
            Add-Result -Step 'SOLIDWORKS install' -Status 'Failed' -ExitCode 'timeout' -ElapsedSec ([math]::Round($sw.Elapsed.TotalSeconds))
            Stop-Deployment -Message "SOLIDWORKS install did not finish within $SwInstallTimeoutMin minutes. sldIM.exe (PID $($sldim.Id)) is still running - investigate before retrying." -ExitCode 30
        }
        try { $exitCode = $sldim.ExitCode } catch { $exitCode = 'unavailable' }
        # sldIM can respawn (elevation, phases) - keep waiting while any live.
        while ($true) {
            $again = Get-Process -Name 'sldIM' -ErrorAction SilentlyContinue | Select-Object -First 1
            if (-not $again) { break }
            try { $null = $again.Handle } catch { }
            Write-Log "A follow-on sldIM.exe process is running (PID $($again.Id)) - continuing to wait..."
            if (-not $again.WaitForExit($SwInstallTimeoutMin * 60 * 1000)) {
                $sw.Stop()
                Add-Result -Step 'SOLIDWORKS install' -Status 'Failed' -ExitCode 'timeout' -ElapsedSec ([math]::Round($sw.Elapsed.TotalSeconds))
                Stop-Deployment -Message "SOLIDWORKS install did not finish within $SwInstallTimeoutMin minutes (follow-on phase)." -ExitCode 30
            }
            try { $exitCode = $again.ExitCode } catch { }
        }
    } else {
        Write-Log "sldIM.exe never appeared within $SldimStartTimeoutSec seconds (StartSWInstall exit code: $($launcher.ExitCode)). Proceeding to validation in case the install completed synchronously." 'WARN'
    }
    $sw.Stop()
    $elapsed = [math]::Round($sw.Elapsed.TotalSeconds)
    Write-Log ("SOLIDWORKS installer phase finished: sldIM exit code {0}, elapsed {1}s" -f $exitCode, $elapsed)

    Write-Log 'Validating SOLIDWORKS installation...'
    if (Test-ProductInstalled -ProductLabel 'SOLIDWORKS' -DisplayNamePattern $SwDisplayNamePattern -VersionPattern $SwVersionPattern -FallbackInstallDir $SwFallbackInstallDir) {
        Add-Result -Step 'SOLIDWORKS install' -Status 'Installed' -ExitCode $exitCode -ElapsedSec $elapsed
        Add-Result -Step 'SOLIDWORKS validation' -Status 'Passed'
        Write-Log 'SOLIDWORKS 2024 installed and validated.' 'OK'
    } else {
        Add-Result -Step 'SOLIDWORKS install' -Status 'Failed' -ExitCode $exitCode -ElapsedSec $elapsed
        Add-Result -Step 'SOLIDWORKS validation' -Status 'Failed'
        Stop-Deployment -Message "SOLIDWORKS validation failed. Check the SOLIDWORKS Installation Manager logs under 'C:\ProgramData\SOLIDWORKS\Installation Logs' (per-version subfolder) for the root cause." -ExitCode 30
    }
}

# =============================================================================
#  MAIN
# =============================================================================

try {
    if (-not (Test-Path -LiteralPath $LogRoot)) {
        New-Item -ItemType Directory -Path $LogRoot -Force | Out-Null
    }
    $script:LogFile = Join-Path $LogRoot ('SWWC_Install_{0}.log' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

    Write-Log ('SWWC deployment starting on {0} (user: {1})' -f $env:COMPUTERNAME, [System.Security.Principal.WindowsIdentity]::GetCurrent().Name) 'STEP'
    Write-Log "Log file: $($script:LogFile)"
    if ($CheckOnly) { Write-Log 'Mode: CHECK ONLY - no changes will be made after prechecks.' 'WARN' }

    Test-Prechecks

    if ($CheckOnly) {
        Add-Result -Step 'WGM install' -Status 'Skipped' -Detail '-CheckOnly'
        Add-Result -Step 'SOLIDWORKS install' -Status 'Skipped' -Detail '-CheckOnly'
        Write-Summary -FinalExitCode 0
        exit 0
    }

    if (-not (Test-Path -LiteralPath $ExtractRoot)) {
        New-Item -ItemType Directory -Path $ExtractRoot -Force | Out-Null
    }

    Install-Wgm
    Install-SolidWorks

    Write-Summary -FinalExitCode 0
    exit 0
}
catch {
    $msg = $_.Exception.Message
    $pos = $_.InvocationInfo.PositionMessage
    Write-Log "UNEXPECTED ERROR: $msg`n$pos" 'ERROR'
    Add-Result -Step 'Script' -Status 'Failed' -Detail $msg
    Write-Summary -FinalExitCode 99
    exit 99
}
