# GPSelector

GhostPractice Database Launcher — a polished WPF UI for switching between GP databases, even across different SQL Server instances.

**Beta 0.0.8** | Created by AZ

## What it does

- Dark, brand-themed selector with one button per saved connection
- Each connection pairs a **SQL Server\Instance** with a **database**, so you can switch between different servers / SQL instances as well as different databases
- One click → updates the relevant fields in GP's `user.config` and launches GhostPractice
- Auto-detects GP version path — survives GP upgrades without reconfiguration
- Supports launching multiple GP instances simultaneously
- Sorts connections by most-used; marks the most recently launched one with a green dot
- Right-click any connection for quick Edit / Remove / Reset usage
- Search box appears automatically when you have more than 8 connections
- Keyboard: `Esc` closes, `F2` opens Settings
- First-run wizard when no connections exist
- Old configs (single `server` + flat `databases[]`) auto-migrate to the new schema
- Click the email in the footer to open your default mail client; click *"Created by AZ"* for a small surprise

## Setup

1. Copy the `GPSelector` folder to any location on the target machine
2. Double-click `Install.bat` to create a desktop shortcut
3. Launch the shortcut and (if first-run) add your first connection

## Logo

The launcher ships with the official **GhostPractice logo** as `gp-logo.png` in the project folder, loaded automatically at runtime. If `gp-logo.png` is missing the launcher falls back to a brand-colored "GP" monogram so the header never looks broken.

The original `gp-logo.svg` is included in the folder for reference. To swap in a different logo, drop a transparent PNG named `gp-logo.png` next to `GPLauncher.ps1` (any reasonable aspect ratio works — `Stretch=Uniform` preserves it).

## Configuration

Edit via the in-app Settings window (recommended), or hand-edit `config.json`:

```json
{
  "gpPath": "C:\\Program Files (x86)\\Korbicom\\GhostPractice\\GhostPractice.exe",
  "connections": [
    { "name": "GD (Express 2022)",  "server": "10.10.10.21\\SQLEXPRESS2022", "database": "GD",       "regenerateReport": true  },
    { "name": "KLS (Express 2017)", "server": "10.10.10.21\\SQLEXPRESS2017", "database": "KLS",      "regenerateReport": true  },
    { "name": "Korbicom (Prod)",    "server": "10.10.10.21",                "database": "Korbicom", "regenerateReport": false }
  ],
  "usage":        { "GD (Express 2022)": 0, "KLS (Express 2017)": 0, "Korbicom (Prod)": 0 },
  "lastLaunched": ""
}
```

The `server` field accepts anything GhostPractice itself accepts: a hostname/IP, a `HOST\INSTANCE` named instance string, or `HOST,PORT`.

## How the GP `user.config` is updated

When you click a connection, the launcher does **not** overwrite the entire `user.config` file. It performs a surgical XML edit, changing **only** these four fields and leaving every other GP user preference (form geometry, column widths, `ApplicationDeviceID`, `HasUpgraded`, etc.) exactly as GP wrote them:

| Field                | Value                                                            |
|----------------------|------------------------------------------------------------------|
| `Server`             | The selected connection's server\instance                        |
| `Database`           | The selected connection's database                               |
| `ApplicationVersion` | The auto-detected GP version (folder name under `Korbitec`)      |
| `RegenerateReport`   | `True` or `False`, taken from each connection's per-DB checkbox  |

The **Regenerate reports on launch** checkbox lives in the Add/Edit Connection dialog in Settings. Default for new connections is on (matches the pre-0.0.8 always-regenerate behaviour). Turn it off for production databases that don't need reports rebuilt every launch.

A backup is taken once per launch session as `user.config.bak` next to the original, before the first edit, so you can roll back if anything goes sideways. If `user.config` doesn't exist yet (brand-new GP install that hasn't run), the launcher falls back to writing a full template once.

## Requirements

- Windows 10/11 with PowerShell 5.1+ (built-in)
- GhostPractice installed and run at least once (so its config folder exists)

## Keyboard shortcuts

| Key   | Where        | Action                              |
|-------|--------------|-------------------------------------|
| `Esc` | any window   | Close (Settings/dialog: discard)    |
| `F2`  | main window  | Open Settings                       |

## Diagnostic log

Each launch writes `gplauncher.log` next to the script with timestamps for opens, edits, saves, `user.config` updates, and any errors. **Send this file with bug reports.**

## File overview

```
GPSelector/
  GPLauncher.ps1     # main script (~1700 lines, fully commented)
  config.json        # your saved connections + usage counts (created on first run)
  gp-logo.png        # logo shown in the header
  gp-logo.svg        # original logo for reference / swapping
  icon.ico           # taskbar / window icon
  Install.bat        # creates a desktop shortcut
  Launch.bat         # launches the script directly
  CreateShortcut.ps1 # helper used by Install.bat
  gplauncher.log     # diagnostic log, rewritten on each launch (gitignored)
  README.md          # this file
```

## Report issues

aubrey.zemba@dyedurham.com
