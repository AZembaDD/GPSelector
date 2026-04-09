# GPSelector

GhostPractice Database Launcher — a minimal WPF UI for switching between multiple GP databases on the same machine.

**Beta 0.0.1** | Created by AZ

## What it does

- Presents a clean dark-themed selector with your configured databases
- Writes the correct `user.config` and launches GhostPractice in one click
- Auto-detects GP version path and Windows user — no manual updates after GP upgrades
- Supports launching multiple GP instances for different databases simultaneously
- Sorts databases by most-used automatically

## Setup

1. Copy the `GPSelector` folder to any location on the target machine
2. Double-click `Install.bat` to create a desktop shortcut
3. Click the shortcut to launch the selector

## Configuration

Edit `config.json` or use the Settings button in the UI:

```json
{
  "server": "10.10.10.21",
  "gpPath": "C:\\Program Files (x86)\\Korbicom\\GhostPractice\\GhostPractice.exe",
  "databases": ["GD", "GhostpracticeCanada", "KLS", "Korbicom", "KorbicomSA"]
}
```

## Requirements

- Windows 10/11 with PowerShell 5.1+ (built-in)
- GhostPractice installed and run at least once

## Report Issues

aubrey.zemba@dyedurham.com
