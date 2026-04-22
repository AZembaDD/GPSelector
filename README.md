# GPSelector

GhostPractice Database Launcher — a polished WPF UI for switching between GP databases, even across different SQL Server instances.

**Beta 0.0.6** | Created by AZ

## What it does

- Dark, brand-themed selector with one button per saved connection
- Each connection pairs a **SQL Server\Instance** with a **database**, so you can switch between different servers/instances as well as different databases
- One click → writes the correct `user.config` and launches GhostPractice
- Auto-detects GP version path — survives GP upgrades without reconfiguration
- Supports launching multiple GP instances simultaneously
- Sorts connections by most-used; marks the most recently launched one with a green dot
- Right-click any connection for quick Edit / Remove / Reset usage
- Search box appears automatically when you have more than 8 connections
- Keyboard: `Esc` closes, `F2` opens Settings
- First-run wizard when no connections exist
- Old configs (single `server` + flat `databases[]`) auto-migrate to the new schema

## Setup

1. Copy the `GPSelector` folder to any location on the target machine
2. Double-click `Install.bat` to create a desktop shortcut
3. Launch the shortcut and (if first-run) add your first connection

## Logo

The launcher ships with a brand-colored "GP" monogram in the header. To use the **official GhostPractice logo** instead, drop a transparent PNG named `gp-logo.png` (recommended size: ~300×60) next to `GPLauncher.ps1`. The launcher loads it automatically on next launch.

The original SVG (`gp-logo.svg`) is included in the folder for reference. Convert it to PNG using any tool (e.g. https://cloudconvert.com/svg-to-png) at 300×60 with transparent background.

## Configuration

Edit via the in-app Settings window (recommended), or hand-edit `config.json`:

```json
{
  "gpPath": "C:\\Program Files (x86)\\Korbicom\\GhostPractice\\GhostPractice.exe",
  "connections": [
    { "name": "GD (Express 2022)",  "server": "10.10.10.21\\SQLEXPRESS2022", "database": "GD" },
    { "name": "KLS (Express 2017)", "server": "10.10.10.21\\SQLEXPRESS2017", "database": "KLS" },
    { "name": "Korbicom",           "server": "10.10.10.21",                "database": "Korbicom" }
  ],
  "usage":        { "GD (Express 2022)": 0, "KLS (Express 2017)": 0, "Korbicom": 0 },
  "lastLaunched": ""
}
```

The `server` field accepts anything GhostPractice itself accepts: a hostname/IP, a `HOST\INSTANCE` named instance string, or `HOST,PORT`.

## Requirements

- Windows 10/11 with PowerShell 5.1+ (built-in)
- GhostPractice installed and run at least once (so its config folder exists)

## Diagnostic log

Each launch writes `gplauncher.log` next to the script with timestamps for opens, edits, saves, and any errors. Send this file with bug reports.

## Report Issues

aubrey.zemba@dyedurham.com
