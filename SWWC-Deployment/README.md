# SWWC Deployment — Windchill Workgroup Manager 13.1 + SOLIDWORKS 2024 SP5.0

Unattended installer for ~20 HP workstations running Windows 11, driven from
pre-staged zip payloads. Targets Windows PowerShell 5.1 (64-bit), running as
`NT AUTHORITY\SYSTEM` under a software deployment platform.

Repo filename is `Install-SWWC-2024.ps1`; it is typically staged on the
workstation as `C:\SWWC\InstallSWWC2024.ps1`. Use whichever path you staged.

| Payload | Staged path |
|---|---|
| Windchill Workgroup Manager 13.1.2.2 CD image | `C:\SWWC\MED-60707-CD-131_13-1-2-2_Windchill-Workgroup-Managers.zip` |
| SOLIDWORKS 2024 SP5.0 | `C:\SWWC\Solidworks_2024\SOLIDWORKS_2024_SP5.0.zip` (~14.8 GB) |

## SOLIDWORKS media policy — read this first

**Unattended deployment is performed only from a genuine SOLIDWORKS
Administrative Image**, using the vendor-documented launcher:

```
<admin_image_root>\StartSWInstall.exe /install /now
```

Ordinary Installation Manager download media **cannot** be installed unattended.
It does not contain the deployment data the install requires — serial number,
product selection, license/SNL server, install options. Those values are
supplied when an Administrative Image is created. This script will not invent
them, will not use undocumented switches, and will not launch the interactive
Installation Manager in Session 0 under SYSTEM. When ordinary media is
detected it **refuses before extracting anything** and exits `31`.

### Why `sldim\StartSWInstall.exe` is not proof of an admin image

The confirmed payload (2,472 entries, one top-level `SOLIDWORKS_2024_SP5.0\`
directory) contains `sldim\startswinstall.exe`. That file ships inside ordinary
media because sldIM copies it into any Administrative Image it later creates.
Classification therefore ignores it and keys on **root-level** artifacts:

| Signature | Meaning |
|---|---|
| root `AdminDirector.xml` + root `StartSWInstall.exe`/`.hta` | Administrative Image (A) |
| root `StartSWInstall.exe` **and** root `StartSWInstall.hta` | Administrative Image (B) |
| root `setup.exe` + `sldim\sldIM.exe` and/or `CheckFile_*.dat` | Ordinary media → refuse |
| several top-level dirs, no root markers | Ambiguous → refuse |

The media root is resolved by peeling single wrapper directories until root
markers appear, so both `payload\setup.exe` and
`payload\SOLIDWORKS_2024_SP5.0\setup.exe` resolve correctly without hardcoding
the folder name. Ambiguous layouts are rejected rather than guessed.

## What the script does

1. **Prechecks** (nothing modified if any fail): PowerShell 5.1+, 64-bit host,
   SYSTEM identity, `C:\SWWC` present, payloads present and non-truncated, disk
   space, no blocking processes, Windows Installer idle (`_MSIExecute` mutex),
   pending-reboot warning.
2. **Read-only payload inspection** — opens both zips' central directories
   *without extracting*, classifies the SOLIDWORKS media, logs the evidence.
   In a real run this **fails fast**, so an undeployable 14.8 GB archive is
   never expanded.
3. **WGM**: extract → install silently → wait → validate. Exit `20` on failure;
   SOLIDWORKS is never attempted.
4. **SOLIDWORKS**: existing-install state check → media gate → extract →
   `StartSWInstall.exe /install /now` → track only *newly started* `sldIM.exe`
   processes → validate.
5. **Summary table** (step / status / exit code / elapsed) on console and log.

### Existing SOLIDWORKS installations

The script classifies what is already present and never silently upgrades or
uninstalls:

| State | Behavior |
|---|---|
| `NotInstalled` | proceed |
| `TargetInstalled` (2024 SP5) | skip, report `AlreadyInstalled` |
| `DifferentServicePack` | stop, exit `32` |
| `DifferentMajorVersion` | stop, exit `32` |
| `Ambiguous` | stop, exit `32` |

### Validation

Validation identifies the **primary** SOLIDWORKS application, not an unrelated
2024 component — component display names (eDrawings, Toolbox, Composer, PDM,
Visualize, …) are excluded. It requires:

- a non-component registry uninstall entry, **and**
- the real `SLDWORKS.exe` on disk, **and**
- `SLDWORKS.exe` file version major `32` (2024) and minor `5` (SP5).

An otherwise-empty install directory is never accepted. Registry reads use the
uninstall hives only — `Win32_Product` is never used.

Logs: `C:\SWWC\Logs\SWWC_Install_<timestamp>.log`, plus a verbose MSI log for
WGM. Extractions are cached under `C:\SWWC\Extract` and reused when the zip is
unchanged (`-ForceExtract` overrides).

## Exit codes

| Code | Meaning |
|---|---|
| 0 | success (check summary for reboot-required flag) |
| 10 | precheck failure |
| 20 | WGM install/validation failure |
| 30 | SOLIDWORKS install/validation failure |
| 31 | SOLIDWORKS payload cannot support unattended deployment |
| 32 | pre-existing SOLIDWORKS in an ambiguous/conflicting state |
| 99 | unexpected script error |

Installer exit codes 3010/1641 are treated as success with a **reboot required**
flag in the summary.

## Test commands

**1. Syntax/parser validation (no execution):**

```powershell
$errs = $null; $toks = $null
[System.Management.Automation.Language.Parser]::ParseFile(
    'C:\SWWC\InstallSWWC2024.ps1', [ref]$toks, [ref]$errs) | Out-Null
if ($errs) { $errs | ForEach-Object { '{0}: {1}' -f $_.Extent.StartLineNumber, $_.Message } }
else { 'PARSE OK' }
```

**2. Read-only precheck (elevated admin):**

```powershell
powershell.exe -ExecutionPolicy Bypass -File C:\SWWC\InstallSWWC2024.ps1 -CheckOnly -AllowNonSystem
```

**3. Read-only precheck under SYSTEM:**

```cmd
psexec -s -accepteula powershell.exe -ExecutionPolicy Bypass -File C:\SWWC\InstallSWWC2024.ps1 -CheckOnly
```

With the current media this is **expected to exit 31** and print the media
classification plus the remediation steps. That is correct behavior, not a bug.

**4. Real deployment — only once an Administrative Image exists:**

```cmd
psexec -s -accepteula powershell.exe -ExecutionPolicy Bypass -File C:\SWWC\InstallSWWC2024.ps1 -SolidWorksAdminImagePath "\\server\share\SOLIDWORKS 2024 SP05"
```

Always run `-CheckOnly` against the image first and confirm it reports
`AdministrativeImage`.

## Still required before a production run

1. **Create the SOLIDWORKS Administrative Image** (serial, product selection,
   license/SNL server) — the current zip is ordinary media and cannot deploy.
2. **Verify the WGM registry patterns.** `*Windchill Workgroup Manager*` /
   `13.1*` and `C:\Program Files\PTC\Windchill Workgroup Manager` are
   unverified; the validation failure message prints what was actually
   registered so they can be corrected once on the pilot machine.
3. **Confirm the WGM silent-install switches.** Auto-detection targets a
   `*wgm*.msi` installed via `msiexec /qn /norestart`. If the CD uses a
   `setup.exe`, set `$WgmInstallerOverride` / `$WgmInstallerOverrideArgs`.
4. **Confirm SP5 minor-version encoding.** Validation expects
   `SLDWORKS.exe` file version `32.5.x`. If your SP5.0 build reports a
   different minor field, adjust `$SwExpectedSpMinor` — do not loosen the
   check to make it pass.

**Pilot on one workstation** before touching the other nineteen.
