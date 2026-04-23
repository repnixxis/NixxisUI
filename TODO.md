# TODO — IrmNixxis Project

> Generated: 2026-04-23  
> Repo: `f:\NCS\Integrations\IrmNixxis`  
> Branch: `main`

---

## 1. Unimplemented Parameters / Missing Phases — `NixxisInstall.ps1`

These parameters are **declared and advertised in help output but have zero implementation**.

### 1.1 `$LicenseKey` — License Key Input
- **Declared at:** `NixxisInstall.ps1:11`
- **Help text at:** `NixxisInstall.ps1:43`
- **Status:** Parameter accepted, stored, never used
- **Expected behavior:** Pass a Nixxis license key and apply it during install (likely writing it to a file or calling a license activation endpoint)
- **Action needed:** Implement a phase that takes `$LicenseKey` and applies it (e.g. writes to `C:\Nixxis\*.lic`, or calls a Nixxis license API)

```powershell
# Line 11
[string]$LicenseKey = "",
# Line 43
Write-Host "  -LicenseKey <string>              Nixxis license key"
```

---

### 1.2 `$SkipLicenseJobs` — License Jobs (Event 903)
- **Declared at:** `NixxisInstall.ps1:14`
- **Help text at:** `NixxisInstall.ps1:46`
- **Status:** Switch defined, no corresponding phase exists
- **Expected behavior:** Configure NCS license jobs (specifically "Event 903") in the Nixxis scheduler or job table
- **Action needed:** Add a new phase (e.g. Phase 9) that configures Event 903 license jobs; gate it with `if (-not $SkipLicenseJobs)`

```powershell
# Line 14
[switch]$SkipLicenseJobs,
# Line 46
Write-Host "  -SkipLicenseJobs                  Skip license job configuration"
```

---

### 1.3 `$SkipEScriptRunner` — eScript Runner Installation
- **Declared at:** `NixxisInstall.ps1:15`
- **Help text at:** `NixxisInstall.ps1:47`
- **Status:** Switch defined, no corresponding phase exists
- **Expected behavior:** Install and configure the eScript Runner component
- **Action needed:** Add a phase (e.g. Phase 10) that installs eScript Runner; gate it with `if (-not $SkipEScriptRunner)`

```powershell
# Line 15
[switch]$SkipEScriptRunner,
# Line 47
Write-Host "  -SkipEScriptRunner                Skip eScript Runner installation"
```

---

## 2. Commented-Out Code — `NixxisUpdate.ps1`

### 2.1 License File Deletion Skipped
- **File:** `NixxisUpdate.ps1:798`
- **Status:** Intentionally (?) commented out — reason unknown
- **Risk:** If `.lic` file should be cleaned up on updates but isn't, stale license files may persist
- **Action needed:** Decide and document: should `CrAppServer.exe.lic` be deleted during updates? If yes, uncomment. If no, add an inline comment explaining why it's excluded.

```powershell
# Line 795-799
"C:\Nixxis\CrAppServer\Twitterizer2.dll",
"C:\Nixxis\CrAppServer\UscNetTools.dll",
"C:\Nixxis\CrAppServer\CRReportingServer.exe",
#"C:\Nixxis\CrAppServer\CrAppServer.exe.lic",   ← WHY IS THIS COMMENTED OUT?
"C:\Nixxis\CrAppServer\ManitermSettings.dll",
```

---

## 3. Manual Interventions — Candidates for Automation

These phases currently require interactive user input (Read-Host / manual steps). They work but break unattended/CI-style installs.

### 3.1 Reporting User Creation — Phase 6
- **File:** `NixxisInstall.ps1:1100–1140`
- **Current behavior:** Prompts "Create Reporting user? (Y/N)" and creates a local Windows account interactively
- **Default credentials:** `Reporting` / `Rep0rting`
- **Action needed:** Add a `-SkipReportingUser` switch (or `-ReportingPassword` param) to allow silent creation; fall back to interactive only when not supplied

```powershell
# Line 1106-1115
Write-Host "  A local user account 'Reporting' can be created for SQL Reporting Services."
Write-Host "  Default credentials: Username=Reporting  Password=Rep0rting"
...
$userChoice = (Read-Host "  Create Reporting user? (Y/N)").Trim()
```

---

### 3.2 Transcription Client — Temp Folder Cleanup
- **File:** `NixxisInstall.ps1:1288`
- **Current behavior:** Prompts "Remove temporary download folder? [Y/N, default Y]"
- **Action needed:** Auto-remove unless a `-KeepTempFiles` switch is passed (default should be auto-clean for unattended runs)

```powershell
# Line 1288
$tcCleanup = (Read-Host "  Remove temporary download folder? [Y/N, default Y]").Trim()
```

---

### 3.3 Reporting Server — Post-Install Manual Copy
- **File:** `NixxisInstall.ps1:1329–1337`
- **Current behavior:** Prints instructions telling the admin to manually copy report files into the SQL Reporting Services path
- **Problem:** Requires human intervention; easy to forget or get wrong
- **Action needed:** Automate the copy using `Copy-Item` targeting the detected SSRS path (MSRS version can be discovered from the registry or by globbing `Program Files\Microsoft SQL Server\MSRS*`)

```powershell
# Line 1329-1337
Write-Host "  POST-INSTALL ACTION REQUIRED - REPORTING SERVER"
Write-Host "  Copy the CONTENTS of the folder:"
Write-Host "  $nixxisInstallPath\Reporting\ToCopyReportingServer"
Write-Host "  INTO (accept folder merge when prompted):"
Write-Host "  [Drive]:\Program Files\Microsoft SQL Server\MSRS.?\"
```

---

### 3.4 MoveFiles Service — Manual XML Editing
- **File:** `NixxisInstall.ps1:1084–1087`
- **Current behavior:** Notifies operator that XML service config must be edited manually
- **Action needed:** Inject values programmatically using `[xml]` or `Select-Xml` / `Set-Content`

---

## 4. Minor / Low-Priority

| # | File | Line | Issue |
|---|------|------|-------|
| 4.1 | `NixxisInstall.ps1` | 118 | `NOTE:` in user-facing output is cosmetic — not a code issue, but consider rephrasing to match the rest of the output style |
| 4.2 | `NixxisUpdate.ps1` | 798 | Add inline comment explaining why `.lic` deletion is suppressed (see §2.1) |
| 4.3 | All scripts | — | No machine-readable version constant — consider adding `$ScriptVersion = "x.y"` at top of each file for easier support/debugging |

---

## 5. Summary Table

| Priority | Item | File | Effort |
|----------|------|------|--------|
| High | Implement `$LicenseKey` usage | `NixxisInstall.ps1` | M |
| High | Implement License Jobs phase (`$SkipLicenseJobs`) | `NixxisInstall.ps1` | L |
| High | Implement eScript Runner phase (`$SkipEScriptRunner`) | `NixxisInstall.ps1` | L |
| Medium | Clarify/uncomment `.lic` file deletion | `NixxisUpdate.ps1` | S |
| Medium | Automate SSRS report file copy | `NixxisInstall.ps1` | M |
| Medium | Automate MoveFiles XML config | `NixxisInstall.ps1` | M |
| Low | Silent Reporting user creation | `NixxisInstall.ps1` | S |
| Low | Silent temp-folder cleanup | `NixxisInstall.ps1` | S |
| Low | Add `$ScriptVersion` constants | All | S |

> **Effort key:** S = < 1h · M = 1–4h · L = half-day+
