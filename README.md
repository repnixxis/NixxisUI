# Nixxis Maintenance Tool — GitHub Repo

A WPF/XAML GUI for Nixxis update and deployment operations, inspired by [Chris Titus WinUtil](https://github.com/ChrisTitusTech/winutil).

## Quick Launch

Open **PowerShell as Administrator** and run:

```powershell
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; irm "https://raw.githubusercontent.com/repnixxis/NixxisUI/main/launch.ps1" | iex
```

> The `[Net.ServicePointManager]` line is required on **Windows Server / PowerShell 5.1** — it defaults to TLS 1.0 which GitHub rejects. This is harmless to include on all systems.

## Features

| Feature | Description |
|---|---|
| Online (Auto) | Auto-discovers the latest version from `update.nixxis.net` |
| Online (Custom URLs) | Manually specify ZIP download URLs |
| Offline Mode | Use pre-downloaded ZIPs from a local folder |
| Full Update | One-click: Download → Extract → Stop → Backup → Cleanup → Deploy |
| Individual Steps | Run any operation in isolation |
| Live Log | Color-coded, scrollable log panel with save-to-file |
| Service Status | Real-time crappserver status indicator |

## Manual Steps Exposed

1. **Download ZIPs** — pulls `ClientProvisioning.zip`, `ClientSoftware.zip`, `NCS.zip`
2. **Extract / Prepare** — builds the `NixxisApplicationServer` staging folder
3. **Stop Service** — stops `crappserver` and kills `nixxisclientdesktop.exe`
4. **Backup** — copies `C:\Nixxis\CrAppServer` and `C:\Nixxis\ClientSoftware` to dated backup folder
5. **Cleanup** — removes stale DLLs, `.pdb` files, and the old `Provisioning` contents
6. **Deploy** — copies new files to `C:\Nixxis\*`
7. **Start Service** — starts `crappserver`

## Requirements

- Windows 10/11 with PowerShell 5.1+
- **Run as Administrator** (auto-elevates via UAC prompt)
- Internet access for online mode; local ZIPs for offline mode

## File Structure

```
NixxisUI/
├── NixxisUI.ps1       # Main script — WPF GUI + all operations
├── launch.ps1         # irm | iex entry point
└── README.md
```

## Setup — Publishing to GitHub

1. Create a **public** GitHub repository named `NixxisUI`
2. Push all files
3. Replace `repnixxis` in both `launch.ps1` and `NixxisUI.ps1` with your GitHub username
4. Share the `irm` command above

## Paths Used

| Path | Purpose |
|---|---|
| `C:\NixxisMaintenance\Logs\` | Log files (one per run) |
| `C:\NixxisMaintenance\BackUp\YYYY\YYYYMMDD\` | Dated backups |
| `C:\Nixxis\` | Live Nixxis installation |
| Script directory | Staging: downloaded ZIPs + `NixxisApplicationServer\` |
