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
foreach ($key in @(
    "HKLM:\SOFTWARE\Classes\AppID\$Clsid",
    'HKLM:\SOFTWARE\Microsoft\Office\OneNote\AddIns\RecallTape.AddIn',
    'HKCU:\SOFTWARE\Microsoft\Office\OneNote\AddIns\RecallTape.AddIn',
    'HKCU:\SOFTWARE\Classes\recalltape',
    'HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Security\Trusted Protocols\All Applications\recalltape:'
)) {
    if (Test-Path $key) {
        Remove-Item $key -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "  removed $key"
    }
}

Write-Host ""
Write-Host "Done. Restart OneNote." -ForegroundColor Green
Write-Host "Your notes are untouched. Any tape still on a page stays there as ordinary content."
