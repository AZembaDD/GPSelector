@echo off
echo Creating desktop shortcut for GhostPractice Launcher...
powershell -ExecutionPolicy Bypass -File "%~dp0CreateShortcut.ps1"
echo Done.
pause
