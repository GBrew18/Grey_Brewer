<#
.SYNOPSIS
    Installs PTC Windchill Workgroup Manager 13.1 and SOLIDWORKS 2024 SP5.0
    from pre-staged zip payloads on Windows 11 workstations.

.DESCRIPTION
    Deployment order (hard requirement):
        1. Prechecks (fail fast, clear messages, nothing modified on failure)
        2. Extract + install Windchill Workgroup Manager, wait, validate
        3. Classify + extract + install SOLIDWORKS 2024, wait, validate
        4. Print summary table (also written to the log)

    If any step fails, the script stops safely at that step, reports why, and
    returns a distinct exit code. It never proceeds to SOLIDWORKS if WGM did
    not install and validate cleanly.

    SOLIDWORKS MEDIA POLICY
    -----------------------
    Unattended SYSTEM-context deployment is performed ONLY from a genuine
    SOLIDWORKS Administrative Image, using the vendor-documented launcher:

        <admin_image_root>\StartSWInstall.exe /install /now

    Ordinary SOLIDWORKS Installation Manager download media CANNOT be deployed
    unattended without vendor-required configuration (serial number, product
    selection, license/SNL server, install options). That configuration lives
    in an Administrative Image, not in the download media. When ordinary media
    is detected, this script REFUSES before extracting or installing and tells
    the administrator exactly what to create. It never launches an interactive
    Installation Manager session in Session 0 under SYSTEM.

    Designed to run as SYSTEM (e.g. deployed via SCCM/Intune/PDQ/psexec -s).

.PARAMETER CheckOnly
    Run all prechecks AND read-only payload structure inspection, but do not
    extract or install anything. Inspects the SOLIDWORKS zip's central
    directory in place - it does not expand the ~14.8 GB archive.

.PARAMETER ForceExtract
    Re-extract zips even if a previous extraction of the same zip is present.

.PARAMETER AllowNonSystem
    Testing escape hatch: permit running as an elevated administrator instead
    of SYSTEM. Production runs should NOT use this.

.PARAMETER SolidWorksAdminImagePath
    Path to a prepared SOLIDWORKS Administrative Image root (the directory
    containing StartSWInstall.exe and AdminDirector.xml), typically a UNC
    share. When supplied, this image is deployed directly and the SOLIDWORKS
    zip payload is ignored. This is the supported production path once an
    Administrative Image has been created.

.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File C:\SWWC\InstallSWWC2024.ps1 -CheckOnly

.EXAMPLE
    psexec -s -accepteula powershell.exe -ExecutionPolicy Bypass -File C:\SWWC\InstallSWWC2024.ps1 -SolidWorksAdminImagePath "\\server\swimage\SOLIDWORKS 2024 SP05"

.NOTES
    Exit codes:
         0  = success (check log/summary for "reboot required" flag)
        10  = precheck failure
        20  = WGM install or validation failure
        30  = SOLIDWORKS install or validation failure
        31  = SOLIDWORKS payload cannot support unattended deployment
              (ordinary media / ambiguous layout / missing admin image config)
        32  = SOLIDWORKS pre-existing installation in an ambiguous state
        99  = unexpected script error

    Requires Windows PowerShell 5.1 (64-bit). No PowerShell 7-only syntax.
    Win32_Product is never used.
#>

[CmdletBinding()]
param(
    [switch]$CheckOnly,
    [switch]$ForceExtract,
    [switch]$AllowNonSystem,
    [string]$SolidWorksAdminImagePath = ''
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
# NOTE: these patterns are UNVERIFIED against a real WGM 13.1 install. The
# validation failure message prints what was actually found so they can be
# corrected once, on the pilot machine.
$WgmDisplayNamePattern = '*Windchill Workgroup Manager*'
$WgmVersionPattern     = '13.1*'
$WgmFallbackInstallDir = 'C:\Program Files\PTC\Windchill Workgroup Manager'
$WgmInstallTimeoutMin  = 60

# --- SOLIDWORKS 2024 SP5.0 --------------------------------------------------
# SOLIDWORKS major version 32 == release year 2024. Second field == service
# pack. SP5.0 therefore expects a 32.5.x file/product version.
$SwExpectedMajorVersion = 32        # 2024
$SwExpectedSpMinor      = 5         # SP5
$SwDisplayNamePattern   = 'SOLIDWORKS 2024*'
$SwFallbackInstallDir   = 'C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS'
$SwPrimaryExeName       = 'SLDWORKS.exe'
$SwInstallTimeoutMin    = 240
$SldimStartTimeoutSec   = 300       # how long to wait for sldIM.exe to appear

# Uninstall-entry DisplayNames that contain "SOLIDWORKS 2024" but are NOT the
# primary SOLIDWORKS application. Used to avoid validating against a component.
$SwComponentExclusions = @(
    '*eDrawings*', '*Explorer*', '*Toolbox*', '*Composer*', '*Visualize*',
    '*Electrical*', '*Inspection*', '*Plastics*', '*Flow Simulation*',
    '*CAM*', '*PDM*', '*Manage*', '*Simulation*', '*Routing*', '*API SDK*',
    '*Language Pack*', '*File Utilities*', '*Document Manager*',
    '*Installation Manager*', '*Marketplace*', '*Sustainability*',
    '*Costing*', '*MBD*', '*Utilities*', '*Backup*', '*Driveworks*',
    '*Model Based Definition*', '*Manufacturing*'
)

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
        [Parameter(Mandatory)][string]$Status,
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
#  SOLIDWORKS MEDIA CLASSIFICATION
# =============================================================================
#
#  Why this exists: sldim\StartSWInstall.exe is present in ORDINARY SOLIDWORKS
#  download media, because sldIM ships that launcher as a payload component
#  which it later copies into any Administrative Image it creates. Finding
#  StartSWInstall.exe anywhere in the tree therefore proves NOTHING about the
#  media type. A genuine Administrative Image is identified by its ROOT-level
#  configuration artifacts:
#
#      <root>\AdminDirector.xml        - admin image director/answer file
#      <root>\StartSWInstall.exe       - launcher AT THE ROOT (not in sldim\)
#      <root>\StartSWInstall.hta       - interactive counterpart of the above
#      <root>\sldAdminOptionEditor.exe - Administrative Image Option Editor
#
#  Ordinary Installation Manager media is identified by:
#
#      <root>\setup.exe                - IM bootstrapper
#      <root>\sldim\sldIM.exe          - Installation Manager itself
#      <root>\CheckFile_*.dat          - media integrity manifests
#
#  Classification is evidence-based and refuses on ambiguity.

function Get-SwMediaClassification {
    <# Classifies a SOLIDWORKS payload from a flat list of forward-slash
       relative paths. Works identically for zip entries and extracted files,
       so -CheckOnly and the real deployment share one code path.
       Returns: MediaRootPrefix, MediaType, Evidence, Reason. #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][string[]]$RelativePaths
    )

    $evidence = New-Object System.Collections.ArrayList
    $norm = @()
    foreach ($p in $RelativePaths) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        $n = $p -replace '\\', '/'
        $n = $n.TrimStart('/')
        # ignore this script's own extraction marker
        if ($n -like '*.swwc_extract_ok') { continue }
        if ($n.EndsWith('/')) { continue }   # directory entries
        $norm += $n
    }

    if ($norm.Count -eq 0) {
        return [pscustomobject]@{
            MediaRootPrefix = $null; MediaType = 'Empty'
            Evidence = @('payload contains no files'); Reason = 'Payload contains no file entries.'
        }
    }
    [void]$evidence.Add("$($norm.Count) file entries examined")

    # --- Determine the media root prefix ------------------------------------
    # Peel single-wrapper directories until root-level marker files appear.
    $prefix = ''
    for ($depth = 0; $depth -lt 4; $depth++) {
        $scoped = @()
        foreach ($n in $norm) {
            if ($prefix -eq '') { $scoped += $n }
            elseif ($n.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                $scoped += $n.Substring($prefix.Length)
            }
        }
        # Root-level files present at this level?
        $rootFiles = @($scoped | Where-Object { $_ -notmatch '/' })
        $markerHere = @($rootFiles | Where-Object {
            $_ -match '^(setup\.exe|StartSWInstall\.(exe|hta)|AdminDirector\.xml|sldAdminOptionEditor\.exe)$'
        })
        if ($markerHere.Count -gt 0) { break }

        # No markers here - descend if there is exactly one top-level directory.
        $topDirs = @()
        foreach ($s in $scoped) {
            if ($s -match '^([^/]+)/') {
                $d = $matches[1]
                if ($topDirs -notcontains $d) { $topDirs += $d }
            }
        }
        if ($topDirs.Count -eq 1) {
            $prefix = $prefix + $topDirs[0] + '/'
            [void]$evidence.Add("descended into single top-level directory '$($topDirs[0])'")
            continue
        }
        if ($topDirs.Count -eq 0) { break }
        return [pscustomobject]@{
            MediaRootPrefix = $null; MediaType = 'Ambiguous'
            Evidence = $evidence
            Reason   = "Could not identify a single media root: $($topDirs.Count) top-level directories with no root marker files ($($topDirs -join ', '))."
        }
    }

    # Re-scope to the resolved root.
    $scoped = @()
    foreach ($n in $norm) {
        if ($prefix -eq '') { $scoped += $n }
        elseif ($n.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            $scoped += $n.Substring($prefix.Length)
        }
    }
    $rootLabel = if ($prefix -eq '') { '<payload root>' } else { $prefix.TrimEnd('/') }
    [void]$evidence.Add("media root resolved to '$rootLabel'")

    function Test-HasPath {
        param([string[]]$Set, [string]$Pattern)
        return (@($Set | Where-Object { $_ -like $Pattern }).Count -gt 0)
    }

    # --- Administrative Image indicators (root level only) ------------------
    $hasAdminDirector = Test-HasPath $scoped 'AdminDirector.xml'
    $hasRootStartExe  = Test-HasPath $scoped 'StartSWInstall.exe'
    $hasRootStartHta  = Test-HasPath $scoped 'StartSWInstall.hta'
    $hasOptionEditor  = Test-HasPath $scoped 'sldAdminOptionEditor.exe'

    # --- Ordinary Installation Manager media indicators ---------------------
    $hasRootSetup     = Test-HasPath $scoped 'setup.exe'
    $hasSldimExe      = Test-HasPath $scoped 'sldim/sldIM.exe'
    $checkFileCount   = @($scoped | Where-Object { $_ -like 'CheckFile_*.dat' }).Count
    $nestedStartExe   = Test-HasPath $scoped 'sldim/StartSWInstall.exe'

    if ($hasAdminDirector) { [void]$evidence.Add('FOUND AdminDirector.xml at media root (admin image marker)') }
    if ($hasRootStartExe)  { [void]$evidence.Add('FOUND StartSWInstall.exe at media root (admin image marker)') }
    if ($hasRootStartHta)  { [void]$evidence.Add('FOUND StartSWInstall.hta at media root (admin image marker)') }
    if ($hasOptionEditor)  { [void]$evidence.Add('FOUND sldAdminOptionEditor.exe at media root (admin image marker)') }
    if ($hasRootSetup)     { [void]$evidence.Add('FOUND setup.exe at media root (IM media marker)') }
    if ($hasSldimExe)      { [void]$evidence.Add('FOUND sldim/sldIM.exe (IM media marker)') }
    if ($checkFileCount -gt 0) { [void]$evidence.Add("FOUND $checkFileCount CheckFile_*.dat manifests (IM media marker)") }
    if ($nestedStartExe -and -not $hasRootStartExe) {
        [void]$evidence.Add('NOTE: StartSWInstall.exe exists ONLY under sldim\ - this is a redistributable component of ordinary media and is NOT evidence of an administrative image')
    }

    # --- Verdict ------------------------------------------------------------
    # Two independent Administrative Image signatures. Both require artifacts
    # that ordinary Installation Manager media does not place at its root.
    #   A: AdminDirector.xml (the answer/config file) + a root launcher.
    #   B: BOTH root launchers (StartSWInstall.exe AND .hta). The .hta is
    #      generated into an admin image; media does not carry one at root.
    if ($hasAdminDirector -and ($hasRootStartExe -or $hasRootStartHta)) {
        return [pscustomobject]@{
            MediaRootPrefix = $prefix; MediaType = 'AdministrativeImage'
            Evidence = $evidence
            Reason   = 'Administrative Image (signature A): root-level AdminDirector.xml plus a root-level StartSWInstall launcher.'
        }
    }
    if ($hasRootStartExe -and $hasRootStartHta) {
        [void]$evidence.Add('NOTE: AdminDirector.xml absent, but both root-level StartSWInstall launchers are present')
        return [pscustomobject]@{
            MediaRootPrefix = $prefix; MediaType = 'AdministrativeImage'
            Evidence = $evidence
            Reason   = 'Administrative Image (signature B): root-level StartSWInstall.exe and StartSWInstall.hta. Ordinary media carries neither at its root.'
        }
    }
    if ($hasRootSetup -and ($hasSldimExe -or $checkFileCount -gt 0)) {
        return [pscustomobject]@{
            MediaRootPrefix = $prefix; MediaType = 'InstallationMedia'
            Evidence = $evidence
            Reason   = 'Ordinary SOLIDWORKS Installation Manager media: root setup.exe with sldim/sldIM.exe and/or CheckFile manifests, and no root-level AdminDirector.xml.'
        }
    }
    return [pscustomobject]@{
        MediaRootPrefix = $prefix; MediaType = 'Unknown'
        Evidence = $evidence
        Reason   = 'Payload matched neither an Administrative Image nor Installation Manager media signature.'
    }
}

function Get-ZipRelativePaths {
    <# Reads a zip's central directory WITHOUT extracting. PS 5.1 compatible. #>
    param([Parameter(Mandatory)][string]$ZipPath)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = $null
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
        $paths = New-Object System.Collections.ArrayList
        foreach ($entry in $archive.Entries) { [void]$paths.Add($entry.FullName) }
        return , $paths.ToArray()
    }
    finally {
        if ($archive) { $archive.Dispose() }
    }
}

function Get-DirectoryRelativePaths {
    <# Enumerates files under a directory as root-relative paths.
       Uses .NET enumeration (no Get-ChildItem -Depth) for PS 5.1 certainty. #>
    param([Parameter(Mandatory)][string]$Root)
    $full = [System.IO.Path]::GetFullPath($Root)
    if (-not $full.EndsWith([System.IO.Path]::DirectorySeparatorChar)) {
        $full = $full + [System.IO.Path]::DirectorySeparatorChar
    }
    $paths = New-Object System.Collections.ArrayList
    foreach ($f in [System.IO.Directory]::EnumerateFiles($full, '*', [System.IO.SearchOption]::AllDirectories)) {
        [void]$paths.Add($f.Substring($full.Length))
    }
    return , $paths.ToArray()
}

function Write-MediaClassification {
    param([Parameter(Mandatory)][object]$Classification, [string]$Label = 'SOLIDWORKS payload')
    Write-Log "$Label media type: $($Classification.MediaType)"
    Write-Log "$Label classification reason: $($Classification.Reason)"
    foreach ($e in $Classification.Evidence) { Write-Log "  evidence: $e" }
}

function Get-UnsupportedMediaMessage {
    param([Parameter(Mandatory)][object]$Classification)
    return @"
SOLIDWORKS payload cannot be deployed unattended as SYSTEM.

Detected media type : $($Classification.MediaType)
Reason              : $($Classification.Reason)

This script deploys SOLIDWORKS only from a genuine Administrative Image, using
the vendor-documented launcher '<admin_image_root>\StartSWInstall.exe /install /now'.

Ordinary Installation Manager media does not contain the deployment data an
unattended install requires - serial number, product selection, license/SNL
server, and install options. Those values are supplied when an Administrative
Image is created; they are NOT present in the download media, and this script
will not invent them or drive the interactive Installation Manager in Session 0.

REQUIRED ADMINISTRATOR ACTION (one time, on a staging machine):
  1. Extract the SOLIDWORKS media and run 'sldim\sldIM.exe' interactively as a
     normal administrator (NOT as SYSTEM).
  2. Choose 'Administrative image' when asked for an installation type.
  3. Supply the serial number, product selection, and license/SNL server.
  4. Let it build the image, then confirm the image ROOT contains BOTH
     'AdminDirector.xml' and 'StartSWInstall.exe'.
  5. Optionally run 'sldAdminOptionEditor.exe' to set per-machine options.
  6. Re-run this script pointing at that image, e.g.:
       -SolidWorksAdminImagePath "\\server\share\SOLIDWORKS 2024 SP05"
     (or re-zip the image to the configured payload path).

No files were extracted or installed for SOLIDWORKS.
"@
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

    $usingExternalImage = -not [string]::IsNullOrWhiteSpace($SolidWorksAdminImagePath)
    $requiredZips = @($WgmZip)
    if (-not $usingExternalImage) { $requiredZips += $SwZip }

    $zipBytesTotal = 0
    foreach ($zip in $requiredZips) {
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

    if ($usingExternalImage) {
        Write-Log "SOLIDWORKS zip payload bypassed; using -SolidWorksAdminImagePath: $SolidWorksAdminImagePath" 'WARN'
        if (-not (Test-Path -LiteralPath $SolidWorksAdminImagePath -PathType Container)) {
            [void]$failures.Add("-SolidWorksAdminImagePath does not exist or is not a directory: $SolidWorksAdminImagePath")
        }
    }

    # -- Disk space ----------------------------------------------------------
    $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'"
    $freeGB = [math]::Round($disk.FreeSpace / 1GB, 1)
    $dynamicNeedGB = [math]::Ceiling(($zipBytesTotal * 2.5) / 1GB) + 10
    $requiredGB = [math]::Max([double]$MinimumFreeSpaceGB, [double]$dynamicNeedGB)
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

function Test-PayloadStructure {
    <# Read-only structural inspection. Never extracts. Runs in every mode so
       an unsupported payload is rejected BEFORE any 14.8 GB extraction. #>
    param([switch]$FailFast)

    Write-Log 'PAYLOAD STRUCTURE INSPECTION (read-only)' 'STEP'

    # --- WGM zip ------------------------------------------------------------
    try {
        $wgmPaths = Get-ZipRelativePaths -ZipPath $WgmZip
        Write-Log "WGM zip opened OK: $($wgmPaths.Count) entries" 'OK'
        $wgmMsis = @($wgmPaths | Where-Object { $_ -match '(?i)wgm.*\.msi$' })
        if ($wgmMsis.Count -gt 0) {
            Write-Log "WGM candidate MSI entries: $($wgmMsis.Count)" 'OK'
            foreach ($m in ($wgmMsis | Select-Object -First 5)) { Write-Log "  candidate: $m" }
        } else {
            $anyInstaller = @($wgmPaths | Where-Object { $_ -match '(?i)(setup\.exe|\.msi)$' } | Select-Object -First 8)
            Write-Log 'No *wgm*.msi found in the WGM zip. Installer candidates:' 'WARN'
            foreach ($m in $anyInstaller) { Write-Log "  candidate: $m" 'WARN' }
            Write-Log 'Set $WgmInstallerOverride in the config block if auto-detection fails at install time.' 'WARN'
        }
        Add-Result -Step 'WGM payload inspect' -Status 'Passed' -Detail "$($wgmPaths.Count) entries, $($wgmMsis.Count) wgm msi"
    } catch {
        Add-Result -Step 'WGM payload inspect' -Status 'Failed' -Detail $_.Exception.Message
        Stop-Deployment -Message "Could not read the WGM zip '$WgmZip': $($_.Exception.Message)" -ExitCode 10
    }

    # --- SOLIDWORKS: zip or external admin image ----------------------------
    $classification = $null
    if (-not [string]::IsNullOrWhiteSpace($SolidWorksAdminImagePath)) {
        try {
            $paths = Get-DirectoryRelativePaths -Root $SolidWorksAdminImagePath
            $classification = Get-SwMediaClassification -RelativePaths $paths
            Write-MediaClassification -Classification $classification -Label 'SOLIDWORKS admin image path'
        } catch {
            Add-Result -Step 'SOLIDWORKS payload inspect' -Status 'Failed' -Detail $_.Exception.Message
            Stop-Deployment -Message "Could not read -SolidWorksAdminImagePath '$SolidWorksAdminImagePath': $($_.Exception.Message)" -ExitCode 31
        }
    } else {
        try {
            $swPaths = Get-ZipRelativePaths -ZipPath $SwZip
            Write-Log "SOLIDWORKS zip opened OK: $($swPaths.Count) entries (central directory read; archive NOT extracted)" 'OK'
            $classification = Get-SwMediaClassification -RelativePaths $swPaths
            Write-MediaClassification -Classification $classification -Label 'SOLIDWORKS zip'
        } catch {
            Add-Result -Step 'SOLIDWORKS payload inspect' -Status 'Failed' -Detail $_.Exception.Message
            Stop-Deployment -Message "Could not read the SOLIDWORKS zip '$SwZip': $($_.Exception.Message)" -ExitCode 31
        }
    }

    $script:SwClassification = $classification

    if ($classification.MediaType -eq 'AdministrativeImage') {
        Add-Result -Step 'SOLIDWORKS payload inspect' -Status 'Passed' -Detail 'Administrative Image'
        Write-Log 'SOLIDWORKS payload supports unattended deployment.' 'OK'
        return
    }

    Add-Result -Step 'SOLIDWORKS payload inspect' -Status 'Failed' -Detail "$($classification.MediaType) - not deployable unattended"
    if ($FailFast) {
        Stop-Deployment -Message (Get-UnsupportedMediaMessage -Classification $classification) -ExitCode 31
    } else {
        Write-Log (Get-UnsupportedMediaMessage -Classification $classification) 'ERROR'
    }
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
    <# Scans 64-bit and 32-bit uninstall hives. Registry only - never
       Win32_Product (which reconfigures packages as a side effect). #>
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
    <# Generic validation: registry uninstall entry matches name+version
       patterns AND the install directory exists. Used for WGM. #>
    param(
        [Parameter(Mandatory)][string]$ProductLabel,
        [Parameter(Mandatory)][string]$DisplayNamePattern,
        [Parameter(Mandatory)][string]$VersionPattern,
        [Parameter(Mandatory)][string]$FallbackInstallDir
    )

    $match = @(Get-InstalledProduct -DisplayNamePattern $DisplayNamePattern -VersionPattern $VersionPattern) | Select-Object -First 1
    if (-not $match) {
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

function Get-SolidWorksPrimaryInstall {
    <# Identifies the PRIMARY SOLIDWORKS application, not a 2024 component.
       Returns an object describing what is installed, or $null.
       State is derived from BOTH registry evidence and SLDWORKS.exe on disk. #>

    $candidates = @(Get-InstalledProduct -DisplayNamePattern 'SOLIDWORKS *')
    $primary = @()
    foreach ($c in $candidates) {
        $excluded = $false
        foreach ($x in $SwComponentExclusions) {
            if ($c.DisplayName -like $x) { $excluded = $true; break }
        }
        if (-not $excluded) { $primary += $c }
    }

    # Locate SLDWORKS.exe: prefer a registry InstallLocation that actually
    # contains it, else the configured fallback directory.
    $exePath = $null
    foreach ($p in $primary) {
        if ($p.PSObject.Properties['InstallLocation'] -and $p.InstallLocation) {
            $candidate = Join-Path $p.InstallLocation.TrimEnd('\') $SwPrimaryExeName
            if (Test-Path -LiteralPath $candidate -PathType Leaf) { $exePath = $candidate; break }
        }
    }
    if (-not $exePath) {
        $candidate = Join-Path $SwFallbackInstallDir $SwPrimaryExeName
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { $exePath = $candidate }
    }

    if ($primary.Count -eq 0 -and -not $exePath) { return $null }

    $fileMajor = $null; $fileMinor = $null; $fileVersion = $null
    if ($exePath) {
        try {
            $vi = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($exePath)
            $fileVersion = $vi.FileVersion
            $fileMajor = $vi.FileMajorPart
            $fileMinor = $vi.FileMinorPart
        } catch {
            Write-Log "Could not read version info from '$exePath': $($_.Exception.Message)" 'WARN'
        }
    }

    return [pscustomobject]@{
        RegistryEntries = $primary
        DisplayName     = if ($primary.Count -gt 0) { $primary[0].DisplayName } else { $null }
        DisplayVersion  = if ($primary.Count -gt 0) { $primary[0].DisplayVersion } else { $null }
        ExePath         = $exePath
        FileVersion     = $fileVersion
        FileMajor       = $fileMajor
        FileMinor       = $fileMinor
    }
}

function Get-SolidWorksInstallState {
    <# Classifies the existing installation into an explicit state so the
       script never guesses: NotInstalled / TargetInstalled /
       DifferentServicePack / DifferentMajorVersion / Ambiguous. #>
    $info = Get-SolidWorksPrimaryInstall
    if (-not $info) {
        return [pscustomobject]@{ State = 'NotInstalled'; Info = $null; Detail = 'No primary SOLIDWORKS application found in registry or on disk.' }
    }

    $desc = "registry='$($info.DisplayName)' regVersion='$($info.DisplayVersion)' exe='$($info.ExePath)' fileVersion='$($info.FileVersion)'"

    if ($null -eq $info.FileMajor) {
        return [pscustomobject]@{ State = 'Ambiguous'; Info = $info; Detail = "SOLIDWORKS evidence present but $SwPrimaryExeName version could not be determined. $desc" }
    }
    if ($info.FileMajor -ne $SwExpectedMajorVersion) {
        return [pscustomobject]@{ State = 'DifferentMajorVersion'; Info = $info; Detail = "Installed SOLIDWORKS major version $($info.FileMajor) != expected $SwExpectedMajorVersion (2024). $desc" }
    }
    if ($info.FileMinor -ne $SwExpectedSpMinor) {
        return [pscustomobject]@{ State = 'DifferentServicePack'; Info = $info; Detail = "SOLIDWORKS 2024 present at SP$($info.FileMinor), expected SP$SwExpectedSpMinor. $desc" }
    }
    return [pscustomobject]@{ State = 'TargetInstalled'; Info = $info; Detail = "SOLIDWORKS 2024 SP$SwExpectedSpMinor confirmed. $desc" }
}

function Test-SolidWorksValidated {
    <# Post-install validation: primary application identified, correct
       major+SP version, and the real SLDWORKS.exe present on disk. #>
    $state = Get-SolidWorksInstallState
    Write-Log "SOLIDWORKS validation state: $($state.State)"
    Write-Log "SOLIDWORKS validation evidence: $($state.Detail)"
    if ($state.State -ne 'TargetInstalled') { return $false }
    if (-not $state.Info.ExePath) {
        Write-Log "SOLIDWORKS: $SwPrimaryExeName was not found on disk; an install directory alone is not accepted as valid." 'ERROR'
        return $false
    }
    Write-Log "SOLIDWORKS: primary executable OK - $($state.Info.ExePath) ($($state.Info.FileVersion))" 'OK'
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
    Write-Log "$Label launched as PID $($proc.Id)"
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
    Write-Log 'STEP 1: WINDCHILL WORKGROUP MANAGER 13.1' 'STEP'

    if (Test-ProductInstalled -ProductLabel 'WGM (pre-scan)' -DisplayNamePattern $WgmDisplayNamePattern -VersionPattern $WgmVersionPattern -FallbackInstallDir $WgmFallbackInstallDir) {
        Write-Log 'Target WGM version already installed - skipping install.' 'OK'
        Add-Result -Step 'WGM install' -Status 'AlreadyInstalled'
        Add-Result -Step 'WGM validation' -Status 'Passed'
        return
    }

    $cdRoot = Expand-Payload -ZipPath $WgmZip

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
        $allFiles = Get-DirectoryRelativePaths -Root $cdRoot
        $msiRel = @($allFiles | Where-Object { $_ -match '(?i)wgm.*\.msi$' })
        $chosenRel = $msiRel | Where-Object { $_ -match '(?i)solidworks|swx' } | Select-Object -First 1
        if (-not $chosenRel) { $chosenRel = $msiRel | Select-Object -First 1 }
        if (-not $chosenRel) {
            $found = @($allFiles | Where-Object { $_ -match '(?i)(setup\.exe|setup_win64\.exe|\.msi)$' } | Select-Object -First 10)
            Write-Log ("No WGM client MSI auto-detected. Installer candidates found on the CD:`n  " + ($found -join "`n  ")) 'ERROR'
            Add-Result -Step 'WGM install' -Status 'Failed' -Detail 'No installer auto-detected'
            Stop-Deployment -Message 'Could not auto-detect the WGM installer. Set $WgmInstallerOverride (and args if not an MSI) in the config block to the correct path from the list above.' -ExitCode 20
        }
        $chosenFull = Join-Path $cdRoot $chosenRel
        Write-Log "Auto-detected WGM installer: $chosenFull"
        $installerArgs = @('/i', "`"$chosenFull`"", '/qn', '/norestart', '/l*v', "`"$msiLog`"")
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
#  STEP 2: SOLIDWORKS 2024 SP5.0
# =============================================================================

function Install-SolidWorks {
    Write-Log 'STEP 2: SOLIDWORKS 2024 SP5.0' 'STEP'

    # --- Existing installation state ---------------------------------------
    $state = Get-SolidWorksInstallState
    Write-Log "Pre-existing SOLIDWORKS state: $($state.State)"
    Write-Log "Pre-existing SOLIDWORKS evidence: $($state.Detail)"

    switch ($state.State) {
        'TargetInstalled' {
            Write-Log 'Target SOLIDWORKS 2024 SP5.0 already installed - skipping install.' 'OK'
            Add-Result -Step 'SOLIDWORKS install' -Status 'AlreadyInstalled' -Detail "SP$SwExpectedSpMinor present"
            Add-Result -Step 'SOLIDWORKS validation' -Status 'Passed'
            return
        }
        'DifferentServicePack' {
            Add-Result -Step 'SOLIDWORKS install' -Status 'Failed' -Detail 'Different SP already installed'
            Stop-Deployment -Message @"
A different SOLIDWORKS 2024 service pack is already installed on this machine.

$($state.Detail)

This script does not automatically upgrade or uninstall an existing SOLIDWORKS
installation - that behaviour has not been designed, tested, or authorised for
this fleet. Decide the upgrade path deliberately (in-place SP upgrade via the
Installation Manager, or uninstall-then-install), then re-run.
"@ -ExitCode 32
        }
        'DifferentMajorVersion' {
            Add-Result -Step 'SOLIDWORKS install' -Status 'Failed' -Detail 'Different major version installed'
            Stop-Deployment -Message @"
A different major version of SOLIDWORKS is already installed on this machine.

$($state.Detail)

Side-by-side major versions and in-place major upgrades are deliberate
decisions with licensing and file-format consequences. This script will not
make that choice automatically. Resolve it, then re-run.
"@ -ExitCode 32
        }
        'Ambiguous' {
            Add-Result -Step 'SOLIDWORKS install' -Status 'Failed' -Detail 'Ambiguous existing install'
            Stop-Deployment -Message @"
SOLIDWORKS appears to be present but its state could not be determined.

$($state.Detail)

Failing safely rather than installing over an unknown installation. Inspect the
machine manually and re-run once the state is clear.
"@ -ExitCode 32
        }
        'NotInstalled' {
            Write-Log 'No existing SOLIDWORKS installation detected - proceeding.' 'OK'
        }
    }

    # --- Media classification gate (already computed read-only) -------------
    $classification = $script:SwClassification
    if (-not $classification) {
        Stop-Deployment -Message 'Internal error: SOLIDWORKS media classification was not performed before install.' -ExitCode 99
    }
    if ($classification.MediaType -ne 'AdministrativeImage') {
        Add-Result -Step 'SOLIDWORKS install' -Status 'Refused' -Detail "$($classification.MediaType)"
        Stop-Deployment -Message (Get-UnsupportedMediaMessage -Classification $classification) -ExitCode 31
    }

    # --- Resolve the admin image root ---------------------------------------
    if (-not [string]::IsNullOrWhiteSpace($SolidWorksAdminImagePath)) {
        $imageRoot = $SolidWorksAdminImagePath.TrimEnd('\')
        Write-Log "Using external administrative image: $imageRoot"
    } else {
        $extracted = Expand-Payload -ZipPath $SwZip
        $prefix = $classification.MediaRootPrefix
        if ([string]::IsNullOrEmpty($prefix)) {
            $imageRoot = $extracted
        } else {
            $imageRoot = Join-Path $extracted ($prefix.TrimEnd('/') -replace '/', '\')
        }
        Write-Log "Extraction root : $extracted"
        Write-Log "Detected zip media root prefix: '$prefix'"
        Write-Log "Resolved administrative image root: $imageRoot"
    }

    $launcher = Join-Path $imageRoot 'StartSWInstall.exe'
    if (-not (Test-Path -LiteralPath $launcher -PathType Leaf)) {
        Add-Result -Step 'SOLIDWORKS install' -Status 'Failed' -Detail 'Launcher missing at resolved root'
        Stop-Deployment -Message "Administrative image was classified as valid, but '$launcher' does not exist on disk. Refusing to search elsewhere for a launcher, because StartSWInstall.exe also ships inside ordinary media under sldim\ and running that copy would not be an administrative-image deployment." -ExitCode 31
    }
    Write-Log "Deployment mechanism: StartSWInstall.exe /install /now (administrative image)" 'OK'
    Write-Log "Selected launcher: $launcher"

    # --- Launch, tracking only processes WE started -------------------------
    # Record pre-existing sldIM PIDs so we can never attach to an unrelated
    # Installation Manager belonging to another session or deployment.
    $preExisting = @(Get-Process -Name 'sldIM' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    if ($preExisting.Count -gt 0) {
        Write-Log "Pre-existing sldIM PIDs (will be ignored): $($preExisting -join ', ')" 'WARN'
    }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $launchProc = Start-Process -FilePath $launcher -ArgumentList '/install', '/now' -PassThru -WorkingDirectory $imageRoot
    $null = $launchProc.Handle
    Write-Log "StartSWInstall.exe launched as PID $($launchProc.Id)"
    $launchProc.WaitForExit(120000) | Out-Null
    $launcherExit = 'running'
    if ($launchProc.HasExited) { $launcherExit = $launchProc.ExitCode }
    Write-Log "StartSWInstall.exe launcher exit code: $launcherExit"

    # --- Find OUR Installation Manager process ------------------------------
    Write-Log "Waiting up to $SldimStartTimeoutSec s for a NEW sldIM.exe process..."
    $ourSldim = $null
    $deadline = (Get-Date).AddSeconds($SldimStartTimeoutSec)
    while ((Get-Date) -lt $deadline) {
        $cands = @(Get-Process -Name 'sldIM' -ErrorAction SilentlyContinue | Where-Object { $preExisting -notcontains $_.Id })
        if ($cands.Count -gt 0) {
            $ourSldim = $cands[0]
            try { $null = $ourSldim.Handle } catch { }
            $exeInfo = ''
            try {
                $cim = Get-CimInstance -ClassName Win32_Process -Filter "ProcessId=$($ourSldim.Id)" -ErrorAction SilentlyContinue
                if ($cim -and $cim.ExecutablePath) { $exeInfo = " path='$($cim.ExecutablePath)'" }
            } catch { }
            Write-Log "Tracking NEW sldIM.exe PID $($ourSldim.Id)$exeInfo" 'OK'
            break
        }
        Start-Sleep -Seconds 5
    }

    $exitCode = 'n/a'
    if ($ourSldim) {
        Write-Log "Waiting for SOLIDWORKS install to complete (timeout $SwInstallTimeoutMin min)..."
        if (-not $ourSldim.WaitForExit($SwInstallTimeoutMin * 60 * 1000)) {
            $sw.Stop()
            Add-Result -Step 'SOLIDWORKS install' -Status 'Failed' -ExitCode 'timeout' -ElapsedSec ([math]::Round($sw.Elapsed.TotalSeconds))
            Stop-Deployment -Message "SOLIDWORKS install did not finish within $SwInstallTimeoutMin minutes. sldIM.exe (PID $($ourSldim.Id)) is still running - investigate before retrying. No process was killed." -ExitCode 30
        }
        try { $exitCode = $ourSldim.ExitCode } catch { $exitCode = 'unavailable' }
        Write-Log "Tracked sldIM.exe PID $($ourSldim.Id) exited with code $exitCode"

        # sldIM can respawn for elevation/phases. Keep waiting on any further
        # NEW instances, never on the pre-existing ones.
        while ($true) {
            $more = @(Get-Process -Name 'sldIM' -ErrorAction SilentlyContinue | Where-Object { $preExisting -notcontains $_.Id })
            if ($more.Count -eq 0) { break }
            $next = $more[0]
            try { $null = $next.Handle } catch { }
            Write-Log "Follow-on sldIM.exe PID $($next.Id) running - continuing to wait..."
            if (-not $next.WaitForExit($SwInstallTimeoutMin * 60 * 1000)) {
                $sw.Stop()
                Add-Result -Step 'SOLIDWORKS install' -Status 'Failed' -ExitCode 'timeout' -ElapsedSec ([math]::Round($sw.Elapsed.TotalSeconds))
                Stop-Deployment -Message "SOLIDWORKS install did not finish within $SwInstallTimeoutMin minutes (follow-on phase, PID $($next.Id))." -ExitCode 30
            }
            try { $exitCode = $next.ExitCode } catch { }
        }
    } else {
        Write-Log "No NEW sldIM.exe appeared within $SldimStartTimeoutSec s (launcher exit code: $launcherExit). Proceeding to validation, which is authoritative." 'WARN'
    }
    $sw.Stop()
    $elapsed = [math]::Round($sw.Elapsed.TotalSeconds)
    Write-Log ("SOLIDWORKS installer phase finished: last tracked exit code {0}, elapsed {1}s" -f $exitCode, $elapsed)

    # --- Validation ---------------------------------------------------------
    Write-Log 'Validating SOLIDWORKS installation...'
    if (Test-SolidWorksValidated) {
        Add-Result -Step 'SOLIDWORKS install' -Status 'Installed' -ExitCode $exitCode -ElapsedSec $elapsed
        Add-Result -Step 'SOLIDWORKS validation' -Status 'Passed'
        Write-Log 'SOLIDWORKS 2024 SP5.0 installed and validated.' 'OK'
    } else {
        Add-Result -Step 'SOLIDWORKS install' -Status 'Failed' -ExitCode $exitCode -ElapsedSec $elapsed
        Add-Result -Step 'SOLIDWORKS validation' -Status 'Failed'
        Stop-Deployment -Message "SOLIDWORKS validation failed. Check the SOLIDWORKS Installation Manager logs under 'C:\ProgramData\SOLIDWORKS\Installation Logs' for the root cause. Validation evidence is logged above." -ExitCode 30
    }
}

# =============================================================================
#  MAIN
# =============================================================================

$script:SwClassification = $null

try {
    if (-not (Test-Path -LiteralPath $LogRoot)) {
        New-Item -ItemType Directory -Path $LogRoot -Force | Out-Null
    }
    $script:LogFile = Join-Path $LogRoot ('SWWC_Install_{0}.log' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

    Write-Log ('SWWC deployment starting on {0} (user: {1})' -f $env:COMPUTERNAME, [System.Security.Principal.WindowsIdentity]::GetCurrent().Name) 'STEP'
    Write-Log "Log file: $($script:LogFile)"
    Write-Log "PowerShell: $($PSVersionTable.PSVersion) | 64-bit process: $([Environment]::Is64BitProcess)"
    if ($CheckOnly) { Write-Log 'Mode: CHECK ONLY - no changes will be made.' 'WARN' }

    Test-Prechecks

    # Structure inspection runs in BOTH modes. In a real run it fails fast so
    # an undeployable 14.8 GB payload is never extracted.
    if ($CheckOnly) {
        Test-PayloadStructure
    } else {
        Test-PayloadStructure -FailFast
    }

    if ($CheckOnly) {
        Add-Result -Step 'WGM install' -Status 'Skipped' -Detail '-CheckOnly'
        Add-Result -Step 'SOLIDWORKS install' -Status 'Skipped' -Detail '-CheckOnly'
        $swReady = ($script:SwClassification -and $script:SwClassification.MediaType -eq 'AdministrativeImage')
        if ($swReady) {
            Write-Summary -FinalExitCode 0
            exit 0
        }
        Write-Log 'CHECK ONLY verdict: prechecks passed, but the SOLIDWORKS payload cannot be deployed unattended (see the media classification above).' 'ERROR'
        Write-Summary -FinalExitCode 31
        exit 31
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
