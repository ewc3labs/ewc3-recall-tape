<#
.SYNOPSIS
    Remove RecallTape completely.

.DESCRIPTION
    Undoes everything install.ps1 did and leaves nothing behind. Needs Administrator, because the
    COM class was registered machine-wide.

    This does NOT touch your notes. Tape already applied to pages stays there -- it is ordinary
    OneNote content. Use "Remove All Tape" on any page you want cleared BEFORE uninstalling, or
    the black boxes will still be sitting there with nothing to take them off.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Continue'

$Clsid = '{AA568A3C-2A53-479B-B188-2367D2E27CE4}'
$dll   = Join-Path $PSScriptRoot 'RecallTape.OneNote.dll'

$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "ERROR: Administrator required. Right-click uninstall.ps1 and 'Run with PowerShell as Administrator'." -ForegroundColor Red
    exit 1
}

if (Get-Process ONENOTE -ErrorAction SilentlyContinue) {
    Write-Host "WARNING: OneNote is running. Close it first, or files may stay locked." -ForegroundColor Yellow
}

Write-Host "Removing RecallTape..."

$regasm = Join-Path $env:WinDir 'Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'
if ((Test-Path $regasm) -and (Test-Path $dll)) {
    & $regasm $dll /unregister | Out-Null
    Write-Host "  COM class unregistered"
}

# HKLM add-in key is included even though install.ps1 never creates one. Anything that ever put
# RecallTape in HKLM - a diagnostic, an older build, a future machine-wide install - would otherwise
# survive uninstall AND be un-tickable from OneNote's COM Add-ins dialog without elevation, because
# clearing that checkbox writes LoadBehavior to whichever hive the key lives in. Sweeping both is
# cheap; leaving a machine-wide registration behind is not.
<#
    Remove a registry key that belongs to RecallTape - and nothing else.

    WHY THIS EXISTS: on 2026-08-13 this script deleted HKCU\SOFTWARE\Classes\CLSID and
    HKCU\SOFTWARE\Classes\AppID in their entirety, on a real machine. The paths were built as
    'HKCU:\SOFTWARE\Classes\CLSID\' + $Clsid inside an array literal, where PowerShell's comma
    operator binds tighter than +, so the GUID was never appended. Remove-Item -Recurse -Force then
    received the PARENT key.

    The string bug was the trigger. The absence of a guard is why it was destructive instead of a
    no-op, so the guard is the actual fix: every key we delete must name RecallTape or our CLSID.
    A path that does not is a bug by construction, and we refuse rather than proceed.
#>
function Remove-RecallTapeKey {
    param([Parameter(Mandatory)][string] $Path)

    if ($Path -notmatch 'RecallTape|recalltape|AA568A3C-2A53-479B-B188-2367D2E27CE4') {
        Write-Host "  REFUSED (does not belong to RecallTape): $Path" -ForegroundColor Red
        return
    }
    if (Test-Path $Path) {
        Remove-Item $Path -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "  removed $Path"
    }
}

foreach ($key in @(
    "HKLM:\SOFTWARE\Classes\AppID\$Clsid",
    'HKLM:\SOFTWARE\Microsoft\Office\OneNote\AddIns\RecallTape.AddIn',
    'HKCU:\SOFTWARE\Microsoft\Office\OneNote\AddIns\RecallTape.AddIn',
    'HKCU:\SOFTWARE\Classes\recalltape',
    'HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Security\Trusted Protocols\All Applications\recalltape:'
)) {
    Remove-RecallTapeKey $key
}

# --- the installed files ---------------------------------------------------------------------
# The installer copies into Program Files, so uninstall has to take that folder away too or we
# leave a dead copy of the add-in on disk forever. Only ever OUR folder, and only if it looks like
# ours - the same refusal rule the registry cleanup uses, for the same reason.
$installed = Join-Path $env:ProgramFiles 'EWC3 Labs\RecallTape'
if (Test-Path $installed) {
    $running = Get-Process ONENOTE -ErrorAction SilentlyContinue
    if ($running) {
        Write-Host ""
        Write-Host "  OneNote is running, so the program files are still locked." -ForegroundColor Yellow
        Write-Host "  Close OneNote completely and delete this folder by hand:"
        Write-Host "    $installed"
    } else {
        try {
            Remove-Item $installed -Recurse -Force -ErrorAction Stop
            Write-Host "  removed $installed"
            # Take the vendor folder too, but ONLY if nothing else of ours lives in it.
            $vendor = Split-Path -Parent $installed
            if ((Test-Path $vendor) -and -not (Get-ChildItem $vendor -Force)) {
                Remove-Item $vendor -Force -ErrorAction SilentlyContinue
            }
        } catch {
            Write-Host "  could not remove $installed - $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

Write-Host ""
Write-Host "Done. Restart OneNote." -ForegroundColor Green
Write-Host "Your notes are untouched. Any tape still on a page stays there as ordinary content."
