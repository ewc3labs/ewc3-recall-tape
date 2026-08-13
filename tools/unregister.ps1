<#
    Undo tools/register.ps1 — the dev-loop twin of uninstall.ps1.

    These two MUST remove the same set of keys. They diverged once: unregister.ps1 was cleaning only
    the COM class and the add-in key, leaving the AppID surrogate, the recalltape:// protocol and its
    Office trust entry behind. A "removal" that leaves half the registration is worse than none,
    because the machine looks clean and is not.

    Requires an elevated shell (RegAsm and the HKLM keys).
#>
param([string] $Configuration = 'Release')
$ErrorActionPreference = 'Continue'

$Clsid = '{AA568A3C-2A53-479B-B188-2367D2E27CE4}'
$repo  = Split-Path -Parent $PSScriptRoot
$dll   = Join-Path $repo ('src\RecallTape.OneNote\bin\' + $Configuration + '\net48\RecallTape.OneNote.dll')

if (Get-Process ONENOTE -ErrorAction SilentlyContinue) {
    Write-Host "WARNING: OneNote is still running. Close it first, or the DLL stays locked." -ForegroundColor Yellow
}

$regasm = Join-Path $env:WinDir 'Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'
if ((Test-Path $regasm) -and (Test-Path $dll)) {
    & $regasm $dll /unregister | Out-Null
    Write-Host "  COM class unregistered"
}

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
    "HKLM:\SOFTWARE\Classes\CLSID\$Clsid",
    "HKLM:\SOFTWARE\Classes\AppID\$Clsid",
    'HKLM:\SOFTWARE\Microsoft\Office\OneNote\AddIns\RecallTape.AddIn',
    'HKCU:\SOFTWARE\Microsoft\Office\OneNote\AddIns\RecallTape.AddIn',
    'HKCU:\SOFTWARE\Classes\recalltape',
    "HKCU:\SOFTWARE\Classes\CLSID\$Clsid",
    "HKCU:\SOFTWARE\Classes\AppID\$Clsid",
    'HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Security\Trusted Protocols\All Applications\recalltape:'
)) {
    Remove-RecallTapeKey $key
}

Write-Host ""
Write-Host "Unregistered. Restart OneNote." -ForegroundColor Green
