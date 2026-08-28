# SWWC Deployment — Windchill Workgroup Manager + SOLIDWORKS 2024 SP5.0

Unattended installer for ~20 HP workstations running Windows 11, driven from
pre-staged zip payloads:

| Payload | Staged path |
|---|---|
| Windchill Workgroup Manager 13.1.2.2 CD image | `C:\SWWC\MED-60707-CD-131_13-1-2-2_Windchill-Workgroup-Managers.zip` |
| SOLIDWORKS 2024 SP5.0 **administrative image** | `C:\SWWC\Solidworks_2024\SOLIDWORKS_2024_SP5.0.zip` |

## What the script does

1. **Prechecks** (nothing is modified if any fail):
   - PowerShell 5.1+, 64-bit host
   - Running as SYSTEM (`-AllowNonSystem` permits an elevated admin, for testing)
   - `C:\SWWC` and both zip payloads present and non-truncated
   - Free disk space on C: (max of a configured floor and a payload-based estimate)
   - No blocking processes (SOLIDWORKS, eDrawings, WGM client, Creo, installers)
   - Windows Installer not busy (checks the real `_MSIExecute` mutex, not just for `msiexec.exe`)
   - Warns (does not block) on a pending Windows reboot
2. **WGM**: extract → install silently → wait → validate. Stops with exit code 20 on failure; SOLIDWORKS is never attempted.
3. **SOLIDWORKS**: extract → `StartSWInstall.exe /install /now` → wait for `sldIM.exe` to finish → validate. Exit code 30 on failure.
4. **Summary table** (step / status / exit code / elapsed seconds) on console and in the log.

Validation for each product = matching registry uninstall entry (name +
version pattern) **and** install directory present on disk. Installer exit
codes and elapsed times are captured for every step. Both products are
skipped as `AlreadyInstalled` if the target version is already present, so
re-running after a partial failure is safe.

Logs: `C:\SWWC\Logs\SWWC_Install_<timestamp>.log`, plus a verbose MSI log for
the WGM install. Extractions are cached under `C:\SWWC\Extract` and reused on
re-runs when the zip is unchanged (`-ForceExtract` overrides).

## Usage

```powershell
# Dry run (prechecks only), from an elevated prompt on a test machine:
powershell.exe -ExecutionPolicy Bypass -File .\Install-SWWC-2024.ps1 -CheckOnly -AllowNonSystem

# Production (as SYSTEM, e.g. via PsExec or your management tool):
psexec -s powershell.exe -ExecutionPolicy Bypass -File C:\SWWC\Install-SWWC-2024.ps1
```

Exit codes: `0` success · `10` precheck failure · `20` WGM failure ·
`30` SOLIDWORKS failure · `99` unexpected error. Exit codes 3010/1641 from an
installer are treated as success with a **reboot required** flag in the summary.

## Assumptions to verify before rollout (see config block at top of script)

1. **The SOLIDWORKS zip must be an administrative image** (contains
   `StartSWInstall.exe`). Admin images embed the serial number and install
   options, which is the supported way to script SOLIDWORKS. If the zip is
   raw download media (`setup.exe`), the script stops with instructions to
   build an admin image first.
2. **WGM installer auto-detection**: the script looks for a `*wgm*.msi` in
   the extracted CD (preferring a SOLIDWORKS-named path, since the CD ships
   workgroup managers for several CAD systems) and installs it with
   `msiexec /qn /norestart` + verbose log. If your CD uses a `setup.exe`
   instead, set `$WgmInstallerOverride` / `$WgmInstallerOverrideArgs`.
3. **Version patterns** used for validation: WGM `13.1*`,
   SOLIDWORKS DisplayName `SOLIDWORKS 2024*` with DisplayVersion `32.5*`
   (2024 SP5.0). If validation fails on the pilot machine, the log prints
   exactly what the registry contains so the patterns can be corrected once.

**Pilot on one workstation first** (with `-CheckOnly`, then a full run) before
touching the other nineteen.
