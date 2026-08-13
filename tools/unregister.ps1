<# Undo tools/register.ps1. Leaves no trace in OneNote or the COM registry. #>
param([string] $Configuration = 'Release')
$ErrorActionPreference = 'Continue'
$repo = Split-Path -Parent $PSScriptRoot
$dll  = Join-Path $repo ('src\RecallTape.OneNote\bin\' + $Configuration + '\net48\RecallTape.OneNote.dll')
& 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe' $dll /unregister | Write-Host
Remove-Item 'HKCU:\SOFTWARE\Microsoft\Office\OneNote\AddIns\RecallTape.AddIn' -Recurse -Force -EA SilentlyContinue
# HKLM too - we never create it, but a diagnostic once did, and an add-in registered machine-wide
# cannot be disabled from OneNote's dialog without elevation.
Remove-Item 'HKLM:\SOFTWARE\Microsoft\Office\OneNote\AddIns\RecallTape.AddIn' -Recurse -Force -EA SilentlyContinue
Write-Host 'Unregistered.'
