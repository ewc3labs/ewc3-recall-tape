@echo off
REM ============================================================================
REM  RecallTape installer - double-click this.
REM
REM  WHY THIS FILE EXISTS: a .ps1 has no "Run as administrator" on its right-click
REM  menu, and Windows blocks scripts that came out of a downloaded zip. Between
REM  them, that is two obstacles a student should never have to solve to install a
REM  study tool. This wrapper elevates itself, unblocks the scripts, and runs the
REM  installer with a policy override scoped to this one process.
REM
REM  Nothing here weakens the machine: -ExecutionPolicy Bypass applies to this
REM  single PowerShell process and nothing else.
REM ============================================================================

setlocal
cd /d "%~dp0"

REM Already elevated? Do the work. Otherwise relaunch ourselves elevated.
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Requesting administrator access...
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
        "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

echo.
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "Get-ChildItem -Path '%~dp0*' -Include *.ps1,*.dll,*.exe -File | Unblock-File"

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"

echo.
echo Press any key to close...
pause >nul
